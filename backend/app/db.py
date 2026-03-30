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
    """)
    await _db.commit()
    logger.info("Feedback database ready at %s", db_path)


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
