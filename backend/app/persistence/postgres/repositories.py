"""
Postgres-backed repositories — full implementation, not a stub.

Uses asyncpg with a connection pool. Creates the same 5 tables as the
SQLite backend; behavioural parity is enforced by the shared test suite.

Schema differences vs SQLite:
- `JSONB` instead of `TEXT` for JSON blobs (better querying later if needed)
- `TIMESTAMPTZ` timestamps generated in Python (kept as ISO strings for
  wire-format symmetry with SQLite — the test suite compares ISO strings)
- ON DELETE CASCADE on `conv_messages.conversation_id`

Connection string accepts both `postgres://` and `postgresql://` prefixes;
asyncpg canonicalises internally.
"""

from __future__ import annotations

import json
import logging
import uuid
from datetime import datetime, timezone
from typing import TYPE_CHECKING, Optional

import asyncpg

from app.persistence.base import (
    ConversationRepository,
    FewShotRepository,
    InteractionRepository,
    Repositories,
)

if TYPE_CHECKING:
    from app.models.chat import ChatResponse, RangeTokenRef as RangeToken

log = logging.getLogger(__name__)


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


_SCHEMA = """
CREATE TABLE IF NOT EXISTS interactions (
    id              TEXT PRIMARY KEY,
    created_at      TEXT NOT NULL,
    user_message    TEXT NOT NULL,
    active_sheet    TEXT,
    workbook_name   TEXT,
    range_tokens    JSONB,
    response_type   TEXT NOT NULL,
    message         TEXT NOT NULL,
    plans_json      JSONB,
    plan_count      INTEGER DEFAULT 0,
    model_used      TEXT,
    latency_ms      INTEGER
);

CREATE TABLE IF NOT EXISTS choices (
    id              TEXT PRIMARY KEY,
    interaction_id  TEXT NOT NULL REFERENCES interactions(id) ON DELETE CASCADE,
    created_at      TEXT NOT NULL,
    chosen_plan_id  TEXT,
    action          TEXT NOT NULL,
    UNIQUE(interaction_id)
);

CREATE TABLE IF NOT EXISTS few_shot_examples (
    id                  TEXT PRIMARY KEY,
    created_at          TEXT NOT NULL,
    user_message        TEXT NOT NULL,
    assistant_response  TEXT NOT NULL,
    source              TEXT NOT NULL DEFAULT 'seed',
    interaction_id      TEXT,
    quality_score       REAL DEFAULT 1.0
);

CREATE TABLE IF NOT EXISTS conversations (
    id              TEXT PRIMARY KEY,
    title           TEXT NOT NULL,
    created_at      TEXT NOT NULL,
    updated_at      TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS conv_messages (
    id              TEXT PRIMARY KEY,
    conversation_id TEXT NOT NULL REFERENCES conversations(id) ON DELETE CASCADE,
    role            TEXT NOT NULL,
    content         TEXT NOT NULL,
    range_tokens_json   JSONB,
    plan_json           JSONB,
    execution_json      JSONB,
    progress_log_json   JSONB,
    created_at      TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_conv_messages_conv_id
    ON conv_messages(conversation_id, created_at);
CREATE INDEX IF NOT EXISTS idx_conversations_updated
    ON conversations(updated_at DESC);
"""


def _normalise_pg_url(url: str) -> str:
    """
    Normalise user-supplied URL to the form asyncpg expects.
    asyncpg rejects the `postgresql+asyncpg://` SQLAlchemy-style prefix.
    """
    if url.startswith("postgresql+asyncpg://"):
        return "postgresql://" + url[len("postgresql+asyncpg://") :]
    return url


class _PgConnection:
    """Shared asyncpg pool — all repos below share one pool."""

    def __init__(self, url: str) -> None:
        self.url = _normalise_pg_url(url)
        self.pool: Optional[asyncpg.Pool] = None

    async def open(self) -> None:
        # Conservative pool: 1-10 connections is plenty for a single-node API.
        self.pool = await asyncpg.create_pool(dsn=self.url, min_size=1, max_size=10)
        async with self.pool.acquire() as conn:
            await conn.execute(_SCHEMA)
        log.info("Postgres persistence ready at %s", _redact(self.url))

    async def close(self) -> None:
        if self.pool:
            await self.pool.close()
            self.pool = None


def _redact(url: str) -> str:
    """Hide password in log output."""
    import re

    return re.sub(r"://([^:]+):([^@]+)@", r"://\1:***@", url)


def _dump_jsonb(value: object) -> Optional[str]:
    """asyncpg JSONB bind: pass raw JSON text, not Python objects."""
    if value is None:
        return None
    return json.dumps(value)


def _load_jsonb(value: object) -> object:
    """asyncpg returns JSONB as a Python object already — but may be a str on older drivers."""
    if value is None:
        return None
    if isinstance(value, str):
        try:
            return json.loads(value)
        except (ValueError, TypeError):
            return None
    return value


# ── Interactions ───────────────────────────────────────────────────────────


class PostgresInteractionRepository(InteractionRepository):
    def __init__(self, holder: _PgConnection) -> None:
        self._holder = holder

    async def log_interaction(
        self,
        *,
        interaction_id: str,
        user_message: str,
        active_sheet: Optional[str],
        workbook_name: Optional[str],
        range_tokens: Optional[list["RangeToken"]],
        response: "ChatResponse",
        model_used: str,
        latency_ms: int,
    ) -> None:
        pool = self._holder.pool
        if pool is None:
            return

        tokens_json: Optional[str] = None
        if range_tokens:
            tokens_json = json.dumps(
                [{"address": t.address, "sheetName": t.sheetName} for t in range_tokens]
            )

        plans_json: Optional[str] = None
        plan_count = 0
        if response.plans:
            plans_json = json.dumps(
                [
                    {"optionLabel": o.optionLabel, "plan": o.plan.model_dump(mode="json")}
                    for o in response.plans
                ]
            )
            plan_count = len(response.plans)
        elif response.plan:
            plans_json = json.dumps(
                [{"optionLabel": "Plan", "plan": response.plan.model_dump(mode="json")}]
            )
            plan_count = 1

        async with pool.acquire() as conn:
            await conn.execute(
                """INSERT INTO interactions
                   (id, created_at, user_message, active_sheet, workbook_name,
                    range_tokens, response_type, message, plans_json, plan_count,
                    model_used, latency_ms)
                   VALUES ($1, $2, $3, $4, $5, $6::jsonb, $7, $8, $9::jsonb, $10, $11, $12)""",
                interaction_id,
                _now_iso(),
                user_message,
                active_sheet,
                workbook_name,
                tokens_json,
                response.responseType,
                response.message,
                plans_json,
                plan_count,
                model_used,
                latency_ms,
            )

    async def log_choice(
        self,
        *,
        interaction_id: str,
        chosen_plan_id: Optional[str],
        action: str,
    ) -> None:
        pool = self._holder.pool
        if pool is None:
            return
        async with pool.acquire() as conn:
            await conn.execute(
                """INSERT INTO choices (id, interaction_id, created_at, chosen_plan_id, action)
                   VALUES ($1, $2, $3, $4, $5)
                   ON CONFLICT (interaction_id)
                   DO UPDATE SET chosen_plan_id = EXCLUDED.chosen_plan_id,
                                 action = EXCLUDED.action,
                                 created_at = EXCLUDED.created_at""",
                str(uuid.uuid4()),
                interaction_id,
                _now_iso(),
                chosen_plan_id,
                action,
            )

    async def get_interaction(self, interaction_id: str) -> Optional[dict]:
        pool = self._holder.pool
        if pool is None:
            return None
        async with pool.acquire() as conn:
            row = await conn.fetchrow(
                "SELECT id, user_message, plans_json::text AS plans_json FROM interactions WHERE id = $1",
                interaction_id,
            )
        if not row:
            return None
        return {"id": row["id"], "user_message": row["user_message"], "plans_json": row["plans_json"]}


# ── Conversations ──────────────────────────────────────────────────────────


class PostgresConversationRepository(ConversationRepository):
    def __init__(self, holder: _PgConnection) -> None:
        self._holder = holder

    async def create(self, title: str) -> str:
        pool = self._holder.pool
        if pool is None:
            raise RuntimeError("Postgres not initialised")
        conv_id = str(uuid.uuid4())
        now = _now_iso()
        async with pool.acquire() as conn:
            await conn.execute(
                "INSERT INTO conversations (id, title, created_at, updated_at) VALUES ($1, $2, $3, $4)",
                conv_id,
                title[:120],
                now,
                now,
            )
        return conv_id

    async def touch(self, conversation_id: str) -> None:
        pool = self._holder.pool
        if pool is None:
            return
        async with pool.acquire() as conn:
            await conn.execute(
                "UPDATE conversations SET updated_at = $1 WHERE id = $2",
                _now_iso(),
                conversation_id,
            )

    async def rename(self, conversation_id: str, title: str) -> bool:
        pool = self._holder.pool
        if pool is None:
            return False
        async with pool.acquire() as conn:
            result = await conn.execute(
                "UPDATE conversations SET title = $1, updated_at = $2 WHERE id = $3",
                title[:120],
                _now_iso(),
                conversation_id,
            )
        # asyncpg returns "UPDATE <count>" — parse the tail number.
        return _affected(result) > 0

    async def delete(self, conversation_id: str) -> bool:
        pool = self._holder.pool
        if pool is None:
            return False
        async with pool.acquire() as conn:
            # ON DELETE CASCADE handles conv_messages automatically.
            result = await conn.execute(
                "DELETE FROM conversations WHERE id = $1", conversation_id
            )
        return _affected(result) > 0

    async def list(self, limit: int = 100) -> list[dict]:
        pool = self._holder.pool
        if pool is None:
            return []
        async with pool.acquire() as conn:
            rows = await conn.fetch(
                """SELECT c.id, c.title, c.created_at, c.updated_at,
                          (SELECT COUNT(*) FROM conv_messages m WHERE m.conversation_id = c.id) AS msg_count
                   FROM conversations c
                   ORDER BY c.updated_at DESC
                   LIMIT $1""",
                limit,
            )
        return [
            {
                "id": r["id"],
                "title": r["title"],
                "createdAt": r["created_at"],
                "updatedAt": r["updated_at"],
                "messageCount": r["msg_count"],
            }
            for r in rows
        ]

    async def load(self, conversation_id: str) -> Optional[dict]:
        pool = self._holder.pool
        if pool is None:
            return None
        async with pool.acquire() as conn:
            conv_row = await conn.fetchrow(
                "SELECT id, title, created_at, updated_at FROM conversations WHERE id = $1",
                conversation_id,
            )
            if not conv_row:
                return None
            msg_rows = await conn.fetch(
                """SELECT id, role, content, range_tokens_json, plan_json,
                          execution_json, progress_log_json, created_at
                   FROM conv_messages
                   WHERE conversation_id = $1
                   ORDER BY created_at ASC""",
                conversation_id,
            )

        messages = [
            {
                "id": r["id"],
                "role": r["role"],
                "content": r["content"],
                "rangeTokens": _load_jsonb(r["range_tokens_json"]),
                "plan": _load_jsonb(r["plan_json"]),
                "execution": _load_jsonb(r["execution_json"]),
                "progressLog": _load_jsonb(r["progress_log_json"]),
                "timestamp": r["created_at"],
            }
            for r in msg_rows
        ]
        return {
            "id": conv_row["id"],
            "title": conv_row["title"],
            "createdAt": conv_row["created_at"],
            "updatedAt": conv_row["updated_at"],
            "messages": messages,
        }

    async def append_message(
        self,
        *,
        conversation_id: str,
        message_id: str,
        role: str,
        content: str,
        range_tokens: Optional[object] = None,
        plan: Optional[object] = None,
    ) -> None:
        pool = self._holder.pool
        if pool is None:
            return
        async with pool.acquire() as conn:
            async with conn.transaction():
                await conn.execute(
                    """INSERT INTO conv_messages
                       (id, conversation_id, role, content, range_tokens_json, plan_json, created_at)
                       VALUES ($1, $2, $3, $4, $5::jsonb, $6::jsonb, $7)""",
                    message_id,
                    conversation_id,
                    role,
                    content,
                    _dump_jsonb(range_tokens),
                    _dump_jsonb(plan),
                    _now_iso(),
                )
                await conn.execute(
                    "UPDATE conversations SET updated_at = $1 WHERE id = $2",
                    _now_iso(),
                    conversation_id,
                )

    async def update_message_execution(
        self,
        *,
        conversation_id: str,
        message_id: str,
        execution: Optional[object],
        progress_log: Optional[object],
    ) -> bool:
        pool = self._holder.pool
        if pool is None:
            return False
        async with pool.acquire() as conn:
            result = await conn.execute(
                """UPDATE conv_messages
                   SET execution_json = $1::jsonb, progress_log_json = $2::jsonb
                   WHERE id = $3 AND conversation_id = $4""",
                _dump_jsonb(execution),
                _dump_jsonb(progress_log),
                message_id,
                conversation_id,
            )
        return _affected(result) > 0

    async def pop_last_exchange(self, conversation_id: str) -> int:
        pool = self._holder.pool
        if pool is None:
            return 0
        async with pool.acquire() as conn:
            rows = await conn.fetch(
                """SELECT id FROM conv_messages
                   WHERE conversation_id = $1
                   ORDER BY created_at DESC LIMIT 2""",
                conversation_id,
            )
            ids_to_delete = [r["id"] for r in rows]
            if not ids_to_delete:
                return 0
            await conn.execute(
                "DELETE FROM conv_messages WHERE id = ANY($1::text[])",
                ids_to_delete,
            )
        return len(ids_to_delete)


# ── Few-shot examples ──────────────────────────────────────────────────────


class PostgresFewShotRepository(FewShotRepository):
    def __init__(self, holder: _PgConnection) -> None:
        self._holder = holder

    async def insert(
        self,
        *,
        example_id: str,
        user_message: str,
        assistant_response: str,
        source: str = "seed",
        interaction_id: Optional[str] = None,
    ) -> None:
        pool = self._holder.pool
        if pool is None:
            return
        async with pool.acquire() as conn:
            await conn.execute(
                """INSERT INTO few_shot_examples
                   (id, created_at, user_message, assistant_response, source, interaction_id)
                   VALUES ($1, $2, $3, $4, $5, $6)
                   ON CONFLICT (id) DO NOTHING""",
                example_id,
                _now_iso(),
                user_message,
                assistant_response,
                source,
                interaction_id,
            )

    async def get_by_ids(self, ids: list[str]) -> list[dict]:
        pool = self._holder.pool
        if pool is None or not ids:
            return []
        async with pool.acquire() as conn:
            rows = await conn.fetch(
                "SELECT id, user_message, assistant_response FROM few_shot_examples WHERE id = ANY($1::text[])",
                ids,
            )
        row_map = {r["id"]: {"id": r["id"], "user_message": r["user_message"], "assistant_response": r["assistant_response"]} for r in rows}
        return [row_map[i] for i in ids if i in row_map]


# ── Bundle ─────────────────────────────────────────────────────────────────


class PostgresRepositories(Repositories):
    """Postgres backend using asyncpg + connection pooling."""

    def __init__(self, database_url: str) -> None:
        self._holder = _PgConnection(database_url)
        self.interactions = PostgresInteractionRepository(self._holder)
        self.conversations = PostgresConversationRepository(self._holder)
        self.few_shot = PostgresFewShotRepository(self._holder)

    async def initialize(self) -> None:
        await self._holder.open()

    async def close(self) -> None:
        await self._holder.close()


def _affected(status: str) -> int:
    """Parse asyncpg's 'UPDATE 3' / 'DELETE 1' status string into a row count."""
    try:
        return int(status.split()[-1])
    except (ValueError, IndexError):
        return 0
