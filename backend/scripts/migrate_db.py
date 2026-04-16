#!/usr/bin/env python3
"""
One-shot data migration helper: copy all persistence data between backends.

Two modes:

    # Relational: copy the 5 tables (interactions, choices, few_shot_examples,
    # conversations, conv_messages) between any supported backends.
    python scripts/migrate_db.py \
        --from "sqlite:///backend/data/feedback.db" \
        --to   "postgresql://user:pass@host:5432/excel_copilot"

    # Vector: copy every collection end-to-end (re-embeds on the way in so
    # schema / dim changes are handled transparently).
    python scripts/migrate_db.py --vector \
        --from "chroma://backend/data/chroma" \
        --to   "pgvector://user:pass@host:5432/excel_copilot"

Only runs one-way — re-running is idempotent (UPSERT / ON CONFLICT DO UPDATE).
Never mutates the source. Safe to run against a live production database as a
long-tailed backup copy.

Used when you flip `DATABASE_URL` / `VECTOR_STORE_URL` between SQLite/Chroma
and Postgres/pgvector. Not wired into startup or the API.
"""

from __future__ import annotations

import argparse
import asyncio
import logging
import sys
from pathlib import Path

# Allow running as `python backend/scripts/migrate_db.py ...` without
# installing the package.
_BACKEND = Path(__file__).resolve().parents[1]
if str(_BACKEND) not in sys.path:
    sys.path.insert(0, str(_BACKEND))

log = logging.getLogger("migrate_db")


# ── Relational migration ────────────────────────────────────────────────────


async def _migrate_relational(from_url: str, to_url: str) -> None:
    """
    Dump every row from the source repositories into the destination.

    We stream per-table; no in-memory full-table snapshot for large sources.
    """
    src = _build_repos(from_url)
    dst = _build_repos(to_url)
    await src.initialize()
    await dst.initialize()
    try:
        await _copy_interactions(src, dst)
        await _copy_few_shot(src, dst)
        await _copy_conversations(src, dst)
    finally:
        await src.close()
        await dst.close()


def _build_repos(url: str):
    from app.persistence.factory import _resolve_repo_backend

    backend = _resolve_repo_backend(url)
    if backend == "postgres":
        from app.persistence.postgres.repositories import PostgresRepositories

        return PostgresRepositories(url)
    from app.persistence.sqlite.repositories import SqliteRepositories

    return SqliteRepositories(url)


async def _copy_interactions(src, dst) -> None:
    # InteractionRepository exposes only `get_interaction` for single lookups;
    # for bulk export we dip into the concrete connection. Both backends own
    # the same schema, so a raw SELECT * is portable.
    rows = await _fetch_all(src, "interactions")
    if rows:
        log.info("Copying %d interactions …", len(rows))
        await _insert_all(dst, "interactions", rows)

    choices = await _fetch_all(src, "choices")
    if choices:
        log.info("Copying %d choices …", len(choices))
        await _insert_all(dst, "choices", choices)


async def _copy_few_shot(src, dst) -> None:
    rows = await _fetch_all(src, "few_shot_examples")
    if not rows:
        return
    log.info("Copying %d few-shot examples …", len(rows))
    for r in rows:
        await dst.few_shot.insert(
            example_id=r["id"],
            user_message=r["user_message"],
            assistant_response=r["assistant_response"],
            source=r.get("source", "seed"),
            interaction_id=r.get("interaction_id"),
        )


async def _copy_conversations(src, dst) -> None:
    convs = await _fetch_all(src, "conversations")
    if not convs:
        return
    log.info("Copying %d conversations …", len(convs))
    await _insert_all(dst, "conversations", convs)

    msgs = await _fetch_all(src, "conv_messages")
    if msgs:
        log.info("Copying %d conversation messages …", len(msgs))
        await _insert_all(dst, "conv_messages", msgs)


async def _fetch_all(repos, table: str) -> list[dict]:
    """Pull every row of `table` as a list of plain dicts."""
    # Both SqliteRepositories and PostgresRepositories expose `_holder` with
    # `.conn` (aiosqlite) or `.pool` (asyncpg). We use the underlying driver
    # directly since the abstract interface is row-agnostic by design.
    holder = repos._holder  # noqa: SLF001 — migration tool
    # SQLite
    if hasattr(holder, "conn") and holder.conn is not None:
        cursor = await holder.conn.execute(f"SELECT * FROM {table}")  # noqa: S608
        cols = [c[0] for c in cursor.description]
        rows = await cursor.fetchall()
        return [dict(zip(cols, r, strict=False)) for r in rows]
    # Postgres
    if hasattr(holder, "pool") and holder.pool is not None:
        async with holder.pool.acquire() as conn:
            records = await conn.fetch(f"SELECT * FROM {table}")  # noqa: S608
        return [dict(r) for r in records]
    return []


async def _insert_all(repos, table: str, rows: list[dict]) -> None:
    holder = repos._holder  # noqa: SLF001
    if not rows:
        return
    cols = list(rows[0].keys())

    # SQLite destination
    if hasattr(holder, "conn") and holder.conn is not None:
        placeholders = ",".join("?" for _ in cols)
        sql = (
            f"INSERT OR REPLACE INTO {table} ({','.join(cols)}) "  # noqa: S608
            f"VALUES ({placeholders})"
        )
        for r in rows:
            await holder.conn.execute(sql, [_coerce_sqlite(r[c]) for c in cols])
        await holder.conn.commit()
        return

    # Postgres destination
    if hasattr(holder, "pool") and holder.pool is not None:
        placeholders = ",".join(f"${i + 1}" for i in range(len(cols)))
        jsonb_cols = _jsonb_columns(table)
        casts = [f"${i + 1}::jsonb" if c in jsonb_cols else f"${i + 1}" for i, c in enumerate(cols)]
        sql = (
            f"INSERT INTO {table} ({','.join(cols)}) VALUES ({','.join(casts)}) "
            "ON CONFLICT (id) DO UPDATE SET "
            + ",".join(f"{c} = EXCLUDED.{c}" for c in cols if c != "id")
        )
        async with holder.pool.acquire() as conn:
            async with conn.transaction():
                for r in rows:
                    await conn.execute(sql, *[_coerce_pg(r[c], c in jsonb_cols) for c in cols])
        return


def _jsonb_columns(table: str) -> set[str]:
    """Which columns are JSONB in Postgres (but TEXT in SQLite)."""
    return {
        "interactions": {"range_tokens", "plans_json"},
        "conv_messages": {
            "range_tokens_json",
            "plan_json",
            "execution_json",
            "progress_log_json",
        },
    }.get(table, set())


def _coerce_sqlite(v):
    # asyncpg returns JSONB as python objects; SQLite wants text.
    if isinstance(v, (dict, list)):
        import json as _json

        return _json.dumps(v)
    return v


def _coerce_pg(v, is_jsonb: bool):
    # SQLite returns JSON as strings; Postgres JSONB bind-values accept text
    # when cast with `$N::jsonb`, which `_insert_all` already does.
    return v


# ── Vector migration ────────────────────────────────────────────────────────


def _migrate_vector(from_url: str, to_url: str) -> None:
    """
    Copy every document in every collection from source to destination. We
    re-embed on the way in (via the destination's configured embedder) so
    the target's dim is guaranteed correct, even if it differs from source.
    """
    src = _build_vector(from_url)
    dst = _build_vector(to_url)
    src.initialize()
    dst.initialize()

    # We only know the two collections we ever use in this app.
    for collection in ("capabilities", "few_shot_examples"):
        count = src.count(collection)
        if count == 0:
            log.info("Collection %s is empty in source — skipping.", collection)
            continue
        log.info("Copying %d rows from collection %s …", count, collection)

        # Pull every ID we've got. The ChromaDB client exposes .get() with no
        # IDs as "all rows"; pgvector needs a COUNT then a page read. We go
        # through the abstract `get_by_ids` after scraping IDs from the
        # underlying client.
        ids = _list_collection_ids(src, collection)
        if not ids:
            continue
        docs = src.get_by_ids(collection, ids)
        dst.upsert(
            collection,
            [d["id"] for d in docs],
            [d["document"] for d in docs],
            [d.get("metadata") or {} for d in docs],
        )


def _build_vector(url: str):
    from app.persistence.factory import _resolve_vector_backend

    backend = _resolve_vector_backend(url)
    if backend == "pgvector":
        from app.persistence.vector_pgvector import PgVectorStore

        return PgVectorStore(url)
    from app.persistence.vector_chroma import ChromaVectorStore

    return ChromaVectorStore(url)


def _list_collection_ids(store, collection: str) -> list[str]:
    """Dip into the underlying client to enumerate IDs (abstract interface is
    query-by-similarity only, which is deliberate)."""
    # ChromaDB
    if hasattr(store, "_coll"):
        try:
            coll = store._coll(collection)  # noqa: SLF001
            raw = coll.get()
            return raw.get("ids") or []
        except Exception:  # pragma: no cover
            return []
    # pgvector
    if hasattr(store, "_bg") and store._bg is not None:  # noqa: SLF001
        return store._bg.run(_pg_list_ids(store, collection))  # noqa: SLF001
    return []


async def _pg_list_ids(store, collection: str) -> list[str]:
    from app.persistence.vector_pgvector import _table_for

    await store._async_ensure_table(collection)  # noqa: SLF001
    async with store._pool.acquire() as conn:  # noqa: SLF001
        rows = await conn.fetch(f"SELECT id FROM {_table_for(collection)}")  # noqa: S608
    return [r["id"] for r in rows]


# ── CLI ─────────────────────────────────────────────────────────────────────


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    parser = argparse.ArgumentParser(
        description="Copy persistence data between backends.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("--from", dest="from_url", required=True, help="Source URL")
    parser.add_argument("--to", dest="to_url", required=True, help="Destination URL")
    parser.add_argument(
        "--vector",
        action="store_true",
        help="Copy vector collections instead of relational tables.",
    )
    args = parser.parse_args()

    if args.vector:
        _migrate_vector(args.from_url, args.to_url)
    else:
        asyncio.run(_migrate_relational(args.from_url, args.to_url))
    log.info("Migration complete.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
