"""
Feedback Database – SQLite-backed interaction logging.

Stores every chat interaction (request + LLM response) and the user's
subsequent choice (applied / dismissed).  Used for future fine-tuning
and quality analysis.

Uses aiosqlite for async access with WAL mode for safe concurrent writes.
"""

from __future__ import annotations

import json
import logging
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import TYPE_CHECKING

import aiosqlite

from .config import settings

if TYPE_CHECKING:
    from .models.chat import ChatResponse, RangeToken

logger = logging.getLogger(__name__)

_db: aiosqlite.Connection | None = None


def _default_db_path() -> Path:
    return Path(__file__).resolve().parents[1] / "data" / "feedback.db"


async def init_db() -> None:
    """Open the database and create tables if they don't exist."""
    global _db  # noqa: PLW0603

    db_path = Path(settings.feedback_db_path or str(_default_db_path()))
    db_path.parent.mkdir(parents=True, exist_ok=True)

    _db = await aiosqlite.connect(str(db_path))
    await _db.execute("PRAGMA journal_mode=WAL")

    await _db.executescript("""
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
    """)
    await _db.commit()
    logger.info("Feedback database ready at %s", db_path)


async def _get_db() -> aiosqlite.Connection:
    """Return the active DB connection or raise if not initialised."""
    if _db is None:
        msg = "Database not initialised – call init_db() first"
        raise RuntimeError(msg)
    return _db


async def close_db() -> None:
    """Close the database connection."""
    global _db  # noqa: PLW0603
    if _db:
        await _db.close()
        _db = None


async def log_interaction(
    *,
    interaction_id: str,
    user_message: str,
    active_sheet: str | None,
    workbook_name: str | None,
    range_tokens: list[RangeToken] | None,
    response: ChatResponse,
    model_used: str,
    latency_ms: int,
) -> None:
    """Record a chat interaction (request + response) to the DB."""
    if not _db:
        return

    # Serialize range tokens
    tokens_json = None
    if range_tokens:
        tokens_json = json.dumps(
            [{"address": t.address, "sheetName": t.sheetName} for t in range_tokens]
        )

    # Serialize plans
    plans_json = None
    plan_count = 0
    if response.plans:
        plans_json = json.dumps(
            [{"optionLabel": o.optionLabel, "plan": o.plan.model_dump(mode="json")} for o in response.plans]
        )
        plan_count = len(response.plans)
    elif response.plan:
        plans_json = json.dumps([{"optionLabel": "Plan", "plan": response.plan.model_dump(mode="json")}])
        plan_count = 1

    await _db.execute(
        """INSERT INTO interactions
           (id, created_at, user_message, active_sheet, workbook_name,
            range_tokens, response_type, message, plans_json, plan_count,
            model_used, latency_ms)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
        (
            interaction_id,
            datetime.now(timezone.utc).isoformat(),
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
    await _db.commit()


async def log_choice(
    *,
    interaction_id: str,
    chosen_plan_id: str | None,
    action: str,
) -> None:
    """Record the user's choice for an interaction."""
    if not _db:
        return

    await _db.execute(
        """INSERT OR REPLACE INTO choices
           (id, interaction_id, created_at, chosen_plan_id, action)
           VALUES (?, ?, ?, ?, ?)""",
        (
            str(uuid.uuid4()),
            interaction_id,
            datetime.now(timezone.utc).isoformat(),
            chosen_plan_id,
            action,
        ),
    )
    await _db.commit()


async def insert_few_shot_example(
    *,
    example_id: str,
    user_message: str,
    assistant_response: str,
    source: str = "seed",
    interaction_id: str | None = None,
) -> None:
    """Insert a few-shot example. Skips silently if the ID already exists."""
    if not _db:
        return
    await _db.execute(
        """INSERT OR IGNORE INTO few_shot_examples
           (id, created_at, user_message, assistant_response, source, interaction_id)
           VALUES (?, ?, ?, ?, ?, ?)""",
        (
            example_id,
            datetime.now(timezone.utc).isoformat(),
            user_message,
            assistant_response,
            source,
            interaction_id,
        ),
    )
    await _db.commit()


async def get_few_shot_examples_by_ids(ids: list[str]) -> list[dict]:
    """Fetch few-shot examples by their IDs. Returns list of dicts."""
    if not _db or not ids:
        return []
    placeholders = ",".join("?" for _ in ids)
    cursor = await _db.execute(
        f"SELECT id, user_message, assistant_response FROM few_shot_examples WHERE id IN ({placeholders})",  # noqa: S608
        ids,
    )
    rows = await cursor.fetchall()
    # Return in the order of the input IDs for relevance ordering
    row_map = {r[0]: {"id": r[0], "user_message": r[1], "assistant_response": r[2]} for r in rows}
    return [row_map[i] for i in ids if i in row_map]


# ── Conversations ────────────────────────────────────────────────────────────

def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


async def create_conversation(title: str) -> str:
    """Create a new conversation and return its id."""
    if not _db:
        raise RuntimeError("DB not initialised")
    conv_id = str(uuid.uuid4())
    now = _now_iso()
    await _db.execute(
        "INSERT INTO conversations (id, title, created_at, updated_at) VALUES (?, ?, ?, ?)",
        (conv_id, title[:120], now, now),
    )
    await _db.commit()
    return conv_id


async def touch_conversation(conversation_id: str) -> None:
    if not _db:
        return
    await _db.execute(
        "UPDATE conversations SET updated_at = ? WHERE id = ?",
        (_now_iso(), conversation_id),
    )
    await _db.commit()


async def rename_conversation(conversation_id: str, title: str) -> bool:
    if not _db:
        return False
    cursor = await _db.execute(
        "UPDATE conversations SET title = ?, updated_at = ? WHERE id = ?",
        (title[:120], _now_iso(), conversation_id),
    )
    await _db.commit()
    return cursor.rowcount > 0


async def delete_conversation(conversation_id: str) -> bool:
    if not _db:
        return False
    await _db.execute("DELETE FROM conv_messages WHERE conversation_id = ?", (conversation_id,))
    cursor = await _db.execute("DELETE FROM conversations WHERE id = ?", (conversation_id,))
    await _db.commit()
    return cursor.rowcount > 0


async def list_conversations(limit: int = 100) -> list[dict]:
    """Return conversations ordered by most recently updated."""
    if not _db:
        return []
    cursor = await _db.execute(
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
            "id": r[0], "title": r[1], "createdAt": r[2],
            "updatedAt": r[3], "messageCount": r[4],
        }
        for r in rows
    ]


async def get_conversation(conversation_id: str) -> dict | None:
    """Fetch a conversation with all its messages."""
    if not _db:
        return None
    cursor = await _db.execute(
        "SELECT id, title, created_at, updated_at FROM conversations WHERE id = ?",
        (conversation_id,),
    )
    conv_row = await cursor.fetchone()
    if not conv_row:
        return None

    cursor = await _db.execute(
        """SELECT id, role, content, range_tokens_json, plan_json,
                  execution_json, progress_log_json, created_at
           FROM conv_messages
           WHERE conversation_id = ?
           ORDER BY created_at ASC""",
        (conversation_id,),
    )
    msg_rows = await cursor.fetchall()

    def _parse(v: str | None) -> object:
        if v is None:
            return None
        try:
            return json.loads(v)
        except (ValueError, TypeError):
            return None

    messages = [
        {
            "id": r[0], "role": r[1], "content": r[2],
            "rangeTokens": _parse(r[3]),
            "plan": _parse(r[4]),
            "execution": _parse(r[5]),
            "progressLog": _parse(r[6]),
            "timestamp": r[7],
        }
        for r in msg_rows
    ]
    return {
        "id": conv_row[0], "title": conv_row[1],
        "createdAt": conv_row[2], "updatedAt": conv_row[3],
        "messages": messages,
    }


async def append_conv_message(
    *,
    conversation_id: str,
    message_id: str,
    role: str,
    content: str,
    range_tokens: object | None = None,
    plan: object | None = None,
) -> None:
    """Append a message to a conversation."""
    if not _db:
        return
    await _db.execute(
        """INSERT INTO conv_messages
           (id, conversation_id, role, content, range_tokens_json, plan_json, created_at)
           VALUES (?, ?, ?, ?, ?, ?, ?)""",
        (
            message_id, conversation_id, role, content,
            json.dumps(range_tokens) if range_tokens is not None else None,
            json.dumps(plan) if plan is not None else None,
            _now_iso(),
        ),
    )
    await _db.execute(
        "UPDATE conversations SET updated_at = ? WHERE id = ?",
        (_now_iso(), conversation_id),
    )
    await _db.commit()


async def update_conv_message_execution(
    *, conversation_id: str, message_id: str,
    execution: object | None, progress_log: object | None,
) -> bool:
    """Attach an execution state + progress log to an existing message."""
    if not _db:
        return False
    cursor = await _db.execute(
        """UPDATE conv_messages
           SET execution_json = ?, progress_log_json = ?
           WHERE id = ? AND conversation_id = ?""",
        (
            json.dumps(execution) if execution is not None else None,
            json.dumps(progress_log) if progress_log is not None else None,
            message_id, conversation_id,
        ),
    )
    await _db.commit()
    return cursor.rowcount > 0


async def pop_last_exchange(conversation_id: str) -> int:
    """Remove the last user+assistant pair. Returns number of rows deleted."""
    if not _db:
        return 0
    cursor = await _db.execute(
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
    await _db.execute(
        f"DELETE FROM conv_messages WHERE id IN ({placeholders})",  # noqa: S608
        ids_to_delete,
    )
    await _db.commit()
    return len(ids_to_delete)


async def get_interaction(interaction_id: str) -> dict | None:
    """Fetch a single interaction by ID."""
    if not _db:
        return None
    cursor = await _db.execute(
        "SELECT id, user_message, plans_json FROM interactions WHERE id = ?",
        (interaction_id,),
    )
    row = await cursor.fetchone()
    if not row:
        return None
    return {"id": row[0], "user_message": row[1], "plans_json": row[2]}


