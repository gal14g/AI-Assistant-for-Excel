"""
SQLite-backed repositories.

Concrete implementation of the persistence interfaces using `aiosqlite`
with WAL mode for safe concurrent access. Behaviourally identical to the
original `backend/app/db.py` module — the code below is a straight migration
of that logic into the repository pattern.

This is the default backend: no deployment-time configuration needed.
The database file lives at `backend/data/feedback.db` unless overridden via
`database_url` / `feedback_db_path`.
"""

from __future__ import annotations

import json
import logging
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import TYPE_CHECKING, Optional

import aiosqlite

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


def _default_db_path() -> Path:
    # backend/app/persistence/sqlite/repositories.py → parents[3] = backend/
    return Path(__file__).resolve().parents[3] / "data" / "feedback.db"


def _normalise_sqlite_url(url: str) -> Path:
    """Accept `""`, a path, or `sqlite:///path` — return a filesystem Path."""
    if not url:
        return _default_db_path()
    if url.startswith("sqlite:///"):
        return Path(url[len("sqlite:///") :])
    if url.startswith("sqlite://"):
        return Path(url[len("sqlite://") :])
    return Path(url)


_SCHEMA = """
CREATE TABLE IF NOT EXISTS interactions (
    id              TEXT PRIMARY KEY,
    created_at      TEXT NOT NULL,
    user_message    TEXT NOT NULL,
    active_sheet    TEXT,
    workbook_name   TEXT,
    range_tokens    TEXT,
    response_type   TEXT NOT NULL,
    message         TEXT NOT NULL,
    plans_json      TEXT,
    plan_count      INTEGER DEFAULT 0,
    model_used      TEXT,
    latency_ms      INTEGER
);

CREATE TABLE IF NOT EXISTS choices (
    id              TEXT PRIMARY KEY,
    interaction_id  TEXT NOT NULL REFERENCES interactions(id),
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
    range_tokens_json   TEXT,
    plan_json           TEXT,
    execution_json      TEXT,
    progress_log_json   TEXT,
    created_at      TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_conv_messages_conv_id
    ON conv_messages(conversation_id, created_at);
CREATE INDEX IF NOT EXISTS idx_conversations_updated
    ON conversations(updated_at DESC);
"""


class _SqliteConnection:
    """
    Shared connection holder. All three repository classes below share a
    single `aiosqlite.Connection` to match the original module-level `_db`
    in `backend/app/db.py`.
    """

    def __init__(self, path: Path) -> None:
        self.path = path
        self.conn: Optional[aiosqlite.Connection] = None

    async def open(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.conn = await aiosqlite.connect(str(self.path))
        await self.conn.execute("PRAGMA journal_mode=WAL")
        await self.conn.executescript(_SCHEMA)
        await self.conn.commit()
        log.info("SQLite persistence ready at %s", self.path)

    async def close(self) -> None:
        if self.conn:
            await self.conn.close()
            self.conn = None


# ── Interactions ───────────────────────────────────────────────────────────


class SqliteInteractionRepository(InteractionRepository):
    def __init__(self, holder: _SqliteConnection) -> None:
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
        db = self._holder.conn
        if db is None:
            return

        tokens_json = None
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

        await db.execute(
            """INSERT INTO interactions
               (id, created_at, user_message, active_sheet, workbook_name,
                range_tokens, response_type, message, plans_json, plan_count,
                model_used, latency_ms)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
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
            ),
        )
        await db.commit()

    async def log_choice(
        self,
        *,
        interaction_id: str,
        chosen_plan_id: Optional[str],
        action: str,
    ) -> None:
        db = self._holder.conn
        if db is None:
            return
        await db.execute(
            """INSERT OR REPLACE INTO choices
               (id, interaction_id, created_at, chosen_plan_id, action)
               VALUES (?, ?, ?, ?, ?)""",
            (str(uuid.uuid4()), interaction_id, _now_iso(), chosen_plan_id, action),
        )
        await db.commit()

    async def get_interaction(self, interaction_id: str) -> Optional[dict]:
        db = self._holder.conn
        if db is None:
            return None
        cursor = await db.execute(
            "SELECT id, user_message, plans_json FROM interactions WHERE id = ?",
            (interaction_id,),
        )
        row = await cursor.fetchone()
        if not row:
            return None
        return {"id": row[0], "user_message": row[1], "plans_json": row[2]}


# ── Conversations ──────────────────────────────────────────────────────────


class SqliteConversationRepository(ConversationRepository):
    def __init__(self, holder: _SqliteConnection) -> None:
        self._holder = holder

    async def create(self, title: str) -> str:
        db = self._holder.conn
        if db is None:
            raise RuntimeError("SQLite not initialised")
        conv_id = str(uuid.uuid4())
        now = _now_iso()
        await db.execute(
            "INSERT INTO conversations (id, title, created_at, updated_at) VALUES (?, ?, ?, ?)",
            (conv_id, title[:120], now, now),
        )
        await db.commit()
        return conv_id

    async def touch(self, conversation_id: str) -> None:
        db = self._holder.conn
        if db is None:
            return
        await db.execute(
            "UPDATE conversations SET updated_at = ? WHERE id = ?",
            (_now_iso(), conversation_id),
        )
        await db.commit()

    async def rename(self, conversation_id: str, title: str) -> bool:
        db = self._holder.conn
        if db is None:
            return False
        cursor = await db.execute(
            "UPDATE conversations SET title = ?, updated_at = ? WHERE id = ?",
            (title[:120], _now_iso(), conversation_id),
        )
        await db.commit()
        return cursor.rowcount > 0

    async def delete(self, conversation_id: str) -> bool:
        db = self._holder.conn
        if db is None:
            return False
        await db.execute("DELETE FROM conv_messages WHERE conversation_id = ?", (conversation_id,))
        cursor = await db.execute("DELETE FROM conversations WHERE id = ?", (conversation_id,))
        await db.commit()
        return cursor.rowcount > 0

    async def list(self, limit: int = 100) -> list[dict]:
        db = self._holder.conn
        if db is None:
            return []
        cursor = await db.execute(
            """SELECT c.id, c.title, c.created_at, c.updated_at,
                      (SELECT COUNT(*) FROM conv_messages m WHERE m.conversation_id = c.id) AS msg_count
               FROM conversations c
               ORDER BY c.updated_at DESC
               LIMIT ?""",
            (limit,),
        )
        rows = await cursor.fetchall()
        return [
            {
                "id": r[0],
                "title": r[1],
                "createdAt": r[2],
                "updatedAt": r[3],
                "messageCount": r[4],
            }
            for r in rows
        ]

    async def load(self, conversation_id: str) -> Optional[dict]:
        db = self._holder.conn
        if db is None:
            return None
        cursor = await db.execute(
            "SELECT id, title, created_at, updated_at FROM conversations WHERE id = ?",
            (conversation_id,),
        )
        conv_row = await cursor.fetchone()
        if not conv_row:
            return None

        cursor = await db.execute(
            """SELECT id, role, content, range_tokens_json, plan_json,
                      execution_json, progress_log_json, created_at
               FROM conv_messages
               WHERE conversation_id = ?
               ORDER BY created_at ASC""",
            (conversation_id,),
        )
        msg_rows = await cursor.fetchall()

        def _parse(v: Optional[str]) -> object:
            if v is None:
                return None
            try:
                return json.loads(v)
            except (ValueError, TypeError):
                return None

        messages = [
            {
                "id": r[0],
                "role": r[1],
                "content": r[2],
                "rangeTokens": _parse(r[3]),
                "plan": _parse(r[4]),
                "execution": _parse(r[5]),
                "progressLog": _parse(r[6]),
                "timestamp": r[7],
            }
            for r in msg_rows
        ]
        return {
            "id": conv_row[0],
            "title": conv_row[1],
            "createdAt": conv_row[2],
            "updatedAt": conv_row[3],
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
        db = self._holder.conn
        if db is None:
            return
        await db.execute(
            """INSERT INTO conv_messages
               (id, conversation_id, role, content, range_tokens_json, plan_json, created_at)
               VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (
                message_id,
                conversation_id,
                role,
                content,
                json.dumps(range_tokens) if range_tokens is not None else None,
                json.dumps(plan) if plan is not None else None,
                _now_iso(),
            ),
        )
        await db.execute(
            "UPDATE conversations SET updated_at = ? WHERE id = ?",
            (_now_iso(), conversation_id),
        )
        await db.commit()

    async def update_message_execution(
        self,
        *,
        conversation_id: str,
        message_id: str,
        execution: Optional[object],
        progress_log: Optional[object],
    ) -> bool:
        db = self._holder.conn
        if db is None:
            return False
        cursor = await db.execute(
            """UPDATE conv_messages
               SET execution_json = ?, progress_log_json = ?
               WHERE id = ? AND conversation_id = ?""",
            (
                json.dumps(execution) if execution is not None else None,
                json.dumps(progress_log) if progress_log is not None else None,
                message_id,
                conversation_id,
            ),
        )
        await db.commit()
        return cursor.rowcount > 0

    async def pop_last_exchange(self, conversation_id: str) -> int:
        db = self._holder.conn
        if db is None:
            return 0
        cursor = await db.execute(
            """SELECT id, role FROM conv_messages
               WHERE conversation_id = ?
               ORDER BY created_at DESC LIMIT 2""",
            (conversation_id,),
        )
        rows = await cursor.fetchall()
        ids_to_delete = [r[0] for r in rows]
        if not ids_to_delete:
            return 0
        placeholders = ",".join("?" for _ in ids_to_delete)
        await db.execute(
            f"DELETE FROM conv_messages WHERE id IN ({placeholders})",  # noqa: S608
            ids_to_delete,
        )
        await db.commit()
        return len(ids_to_delete)


# ── Few-shot examples ──────────────────────────────────────────────────────


class SqliteFewShotRepository(FewShotRepository):
    def __init__(self, holder: _SqliteConnection) -> None:
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
        db = self._holder.conn
        if db is None:
            return
        await db.execute(
            """INSERT OR IGNORE INTO few_shot_examples
               (id, created_at, user_message, assistant_response, source, interaction_id)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (example_id, _now_iso(), user_message, assistant_response, source, interaction_id),
        )
        await db.commit()

    async def get_by_ids(self, ids: list[str]) -> list[dict]:
        db = self._holder.conn
        if db is None or not ids:
            return []
        placeholders = ",".join("?" for _ in ids)
        cursor = await db.execute(
            f"SELECT id, user_message, assistant_response FROM few_shot_examples WHERE id IN ({placeholders})",  # noqa: S608
            ids,
        )
        rows = await cursor.fetchall()
        row_map = {r[0]: {"id": r[0], "user_message": r[1], "assistant_response": r[2]} for r in rows}
        return [row_map[i] for i in ids if i in row_map]


# ── Bundle ─────────────────────────────────────────────────────────────────


class SqliteRepositories(Repositories):
    """Default persistence backend — one aiosqlite connection, three repos."""

    def __init__(self, database_url: str = "") -> None:
        self._holder = _SqliteConnection(_normalise_sqlite_url(database_url))
        self.interactions = SqliteInteractionRepository(self._holder)
        self.conversations = SqliteConversationRepository(self._holder)
        self.few_shot = SqliteFewShotRepository(self._holder)

    async def initialize(self) -> None:
        await self._holder.open()

    async def close(self) -> None:
        await self._holder.close()
