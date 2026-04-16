"""
pgvector-backed vector store.

Requires the `vector` extension in the target Postgres database (the
`initialize()` call runs `CREATE EXTENSION IF NOT EXISTS vector`). Stores
each collection in its own table with a fixed-dim `VECTOR(N)` column
matching the embedding model's dimensionality.

Embedding is done in-process via `persistence/embedding.embed()` — the
same SentenceTransformer model the ChromaDB backend uses, so backend
swaps don't change retrieval quality.

Design choices:
- One table per collection, named `vec_<collection>`. Easier to nuke a
  collection with `TRUNCATE` and keeps the row schema simple.
- `id TEXT PRIMARY KEY` matches the ChromaDB backend so callers can use
  the same string IDs on both sides.
- `metadata JSONB` — passed through as opaque dict.
- Cosine distance (`<=>`) for retrieval; lower = more similar.
- A separate `asyncpg` pool is lazily created on first call to avoid
  synchronous interpreter deadlocks — the abstract contract is sync
  (like ChromaDB's), so we run the coroutines on a bounded event loop.
"""

from __future__ import annotations

import asyncio
import json
import logging
import re
import threading
from typing import Any, Optional

import asyncpg

from app.persistence.base import VectorStore

log = logging.getLogger(__name__)


def _normalise_pgvector_url(url: str) -> str:
    """
    `pgvector://user:pass@host:5432/db` → `postgresql://user:pass@host:5432/db`
    Leave anything else alone — the caller may hand us a full `postgresql://`.
    """
    if not url:
        raise ValueError("pgvector store requires a connection URL")
    if url.startswith("pgvector://"):
        return "postgresql://" + url[len("pgvector://") :]
    if url.startswith("postgresql+asyncpg://"):
        return "postgresql://" + url[len("postgresql+asyncpg://") :]
    return url


_VALID_NAME = re.compile(r"^[a-z_][a-z0-9_]*$", re.IGNORECASE)


def _table_for(collection: str) -> str:
    """Protect against SQL injection via the collection name."""
    if not _VALID_NAME.match(collection):
        raise ValueError(f"Invalid collection name: {collection!r}")
    return f"vec_{collection}"


class _BackgroundLoop:
    """
    Dedicated asyncio loop on a background thread so the sync `VectorStore`
    interface can drive asyncpg coroutines without hijacking the caller's
    loop. Created once per `PgVectorStore` instance.
    """

    def __init__(self) -> None:
        self.loop = asyncio.new_event_loop()
        self._thread = threading.Thread(
            target=self.loop.run_forever, daemon=True, name="pgvector-loop"
        )
        self._thread.start()

    def run(self, coro):
        future = asyncio.run_coroutine_threadsafe(coro, self.loop)
        return future.result()

    def close(self) -> None:
        self.loop.call_soon_threadsafe(self.loop.stop)


class PgVectorStore(VectorStore):
    """Postgres + pgvector implementation of the VectorStore contract."""

    def __init__(self, url: str) -> None:
        self._url = _normalise_pgvector_url(url)
        self._pool: Optional[asyncpg.Pool] = None
        self._dim: Optional[int] = None
        self._bg: Optional[_BackgroundLoop] = None
        self._initialised_tables: set[str] = set()

    # ── Lifecycle ────────────────────────────────────────────────────────────

    def initialize(self) -> None:
        from app.persistence.embedding import embedding_dimensions

        self._dim = embedding_dimensions()
        self._bg = _BackgroundLoop()
        self._bg.run(self._async_open())
        log.info(
            "pgvector store ready (dim=%d) at %s",
            self._dim,
            _redact(self._url),
        )

    async def _async_open(self) -> None:
        self._pool = await asyncpg.create_pool(dsn=self._url, min_size=1, max_size=10)
        async with self._pool.acquire() as conn:
            await conn.execute("CREATE EXTENSION IF NOT EXISTS vector")

    async def _async_ensure_table(self, collection: str) -> None:
        assert self._pool is not None and self._dim is not None
        table = _table_for(collection)
        if table in self._initialised_tables:
            return
        async with self._pool.acquire() as conn:
            await conn.execute(
                f"""
                CREATE TABLE IF NOT EXISTS {table} (
                    id        TEXT PRIMARY KEY,
                    document  TEXT NOT NULL,
                    metadata  JSONB,
                    embedding VECTOR({self._dim}) NOT NULL
                )
                """
            )
            # IVFFlat is a good default for mid-size corpora; cheap to (re)build.
            await conn.execute(
                f"CREATE INDEX IF NOT EXISTS {table}_emb_idx ON {table} "
                f"USING ivfflat (embedding vector_cosine_ops) WITH (lists = 50)"
            )
        self._initialised_tables.add(table)

    # ── Contract ─────────────────────────────────────────────────────────────

    def upsert(
        self,
        collection: str,
        ids: list[str],
        documents: list[str],
        metadatas: list[dict[str, Any]],
    ) -> None:
        if not ids:
            return
        from app.persistence.embedding import embed

        vectors = embed(documents)
        assert self._bg is not None
        self._bg.run(self._async_upsert(collection, ids, documents, metadatas, vectors))

    async def _async_upsert(
        self,
        collection: str,
        ids: list[str],
        documents: list[str],
        metadatas: list[dict[str, Any]],
        vectors: list[list[float]],
    ) -> None:
        await self._async_ensure_table(collection)
        # Fail loudly if someone hands us vectors that don't match the table's
        # VECTOR(N) declaration — Postgres would otherwise coerce/truncate
        # silently and poison the index. This catches a classic footgun when
        # the embedding model gets swapped mid-flight without `recreate()`.
        assert self._dim is not None
        bad = [i for i, v in enumerate(vectors) if len(v) != self._dim]
        if bad:
            raise ValueError(
                f"Vector dimension mismatch: collection={collection!r} "
                f"expects dim={self._dim}, got {len(vectors[bad[0]])} "
                f"at index {bad[0]} (and {len(bad) - 1} more). "
                f"Did the embedding model change? Run `recreate({collection!r})` "
                f"to drop and rebuild the table."
            )
        table = _table_for(collection)
        rows = [
            (ids[i], documents[i], json.dumps(metadatas[i] or {}), _pg_vector_literal(vectors[i]))
            for i in range(len(ids))
        ]
        assert self._pool is not None
        async with self._pool.acquire() as conn:
            async with conn.transaction():
                await conn.executemany(
                    f"""
                    INSERT INTO {table} (id, document, metadata, embedding)
                    VALUES ($1, $2, $3::jsonb, $4::vector)
                    ON CONFLICT (id) DO UPDATE
                      SET document = EXCLUDED.document,
                          metadata = EXCLUDED.metadata,
                          embedding = EXCLUDED.embedding
                    """,
                    rows,
                )

    def query(
        self,
        collection: str,
        text: str,
        top_k: int,
        where: Optional[dict[str, Any]] = None,
    ) -> list[dict[str, Any]]:
        from app.persistence.embedding import embed

        vec = embed([text])
        if not vec:
            return []
        assert self._bg is not None
        return self._bg.run(self._async_query(collection, vec[0], top_k, where or {}))

    async def _async_query(
        self,
        collection: str,
        vector: list[float],
        top_k: int,
        where: dict[str, Any],
    ) -> list[dict[str, Any]]:
        await self._async_ensure_table(collection)
        table = _table_for(collection)
        assert self._pool is not None

        # Optional metadata filter — only supports equality on top-level keys
        # (matches the subset of ChromaDB `where` we actually use).
        #
        # Build parameters in order: [$1 vector, ($key, $val) pairs per filter, $last top_k].
        # Each filter uses TWO placeholders: $N for the JSON key name, $N+1 for the value.
        args: list[Any] = [_pg_vector_literal(vector)]
        where_clauses: list[str] = []
        for k, v in (where or {}).items():
            key_idx = len(args) + 1
            val_idx = key_idx + 1
            args.extend([k, v])
            where_clauses.append(f"metadata->>${key_idx} = ${val_idx}")
        where_sql = "WHERE " + " AND ".join(where_clauses) if where_clauses else ""
        args.append(top_k)

        sql = (
            f"SELECT id, document, metadata, embedding <=> $1::vector AS distance "
            f"FROM {table} {where_sql} "
            f"ORDER BY embedding <=> $1::vector LIMIT ${len(args)}"
        )

        async with self._pool.acquire() as conn:
            rows = await conn.fetch(sql, *args)

        return [
            {
                "id": r["id"],
                "document": r["document"],
                "metadata": _load_jsonb(r["metadata"]),
                "distance": float(r["distance"]),
            }
            for r in rows
        ]

    def get_by_ids(self, collection: str, ids: list[str]) -> list[dict[str, Any]]:
        if not ids:
            return []
        assert self._bg is not None
        return self._bg.run(self._async_get_by_ids(collection, ids))

    async def _async_get_by_ids(
        self, collection: str, ids: list[str]
    ) -> list[dict[str, Any]]:
        await self._async_ensure_table(collection)
        table = _table_for(collection)
        assert self._pool is not None
        async with self._pool.acquire() as conn:
            rows = await conn.fetch(
                f"SELECT id, document, metadata FROM {table} WHERE id = ANY($1::text[])",
                ids,
            )
        row_map = {
            r["id"]: {
                "id": r["id"],
                "document": r["document"],
                "metadata": _load_jsonb(r["metadata"]),
                "distance": None,
            }
            for r in rows
        }
        return [row_map[i] for i in ids if i in row_map]

    def delete(self, collection: str, ids: Optional[list[str]] = None) -> None:
        assert self._bg is not None
        self._bg.run(self._async_delete(collection, ids))

    async def _async_delete(
        self, collection: str, ids: Optional[list[str]]
    ) -> None:
        await self._async_ensure_table(collection)
        table = _table_for(collection)
        assert self._pool is not None
        async with self._pool.acquire() as conn:
            if ids is None:
                await conn.execute(f"TRUNCATE {table}")
            elif ids:
                await conn.execute(
                    f"DELETE FROM {table} WHERE id = ANY($1::text[])", ids
                )

    def count(self, collection: str) -> int:
        assert self._bg is not None
        return self._bg.run(self._async_count(collection))

    async def _async_count(self, collection: str) -> int:
        await self._async_ensure_table(collection)
        table = _table_for(collection)
        assert self._pool is not None
        async with self._pool.acquire() as conn:
            return await conn.fetchval(f"SELECT COUNT(*) FROM {table}")

    def recreate(self, collection: str) -> None:
        assert self._bg is not None
        self._bg.run(self._async_recreate(collection))

    async def _async_recreate(self, collection: str) -> None:
        table = _table_for(collection)
        assert self._pool is not None
        async with self._pool.acquire() as conn:
            await conn.execute(f"DROP TABLE IF EXISTS {table}")
        self._initialised_tables.discard(table)
        await self._async_ensure_table(collection)

    # ── Shutdown (for tests) ─────────────────────────────────────────────────

    def close(self) -> None:  # pragma: no cover — factory caches for life of proc
        if self._bg is not None and self._pool is not None:
            self._bg.run(self._pool.close())
            self._bg.close()


def _pg_vector_literal(v: list[float]) -> str:
    """pgvector accepts text literals like `[0.1, 0.2, 0.3]`."""
    return "[" + ",".join(f"{x:.8f}" for x in v) + "]"


def _load_jsonb(value: object) -> dict[str, Any]:
    if value is None:
        return {}
    if isinstance(value, str):
        try:
            return json.loads(value)
        except (ValueError, TypeError):
            return {}
    if isinstance(value, dict):
        return value
    return {}


def _redact(url: str) -> str:
    return re.sub(r"://([^:]+):([^@]+)@", r"://\1:***@", url)
