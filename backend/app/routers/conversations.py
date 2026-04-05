"""
Conversations API — persisted chat history.

Conversations are created automatically when the /api/chat endpoint is
called without a conversationId. These endpoints let the frontend list,
load, rename, and delete past conversations, and attach post-hoc
execution state to specific messages for the execution timeline.
"""

from __future__ import annotations

import logging
from typing import Any, Optional

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel, Field
from slowapi import Limiter
from slowapi.util import get_remote_address

from ..db import (
    delete_conversation,
    get_conversation,
    list_conversations,
    pop_last_exchange,
    rename_conversation,
    update_conv_message_execution,
)

logger = logging.getLogger(__name__)
limiter = Limiter(key_func=get_remote_address)

router = APIRouter(prefix="/api/conversations", tags=["conversations"])


class ConversationSummary(BaseModel):
    id: str
    title: str
    createdAt: str
    updatedAt: str
    messageCount: int


class ConversationMessage(BaseModel):
    id: str
    role: str
    content: str
    timestamp: str
    rangeTokens: Optional[Any] = None
    plan: Optional[Any] = None
    execution: Optional[Any] = None
    progressLog: Optional[Any] = None


class ConversationDetail(BaseModel):
    id: str
    title: str
    createdAt: str
    updatedAt: str
    messages: list[ConversationMessage]


class RenameRequest(BaseModel):
    title: str = Field(..., min_length=1, max_length=120)


class ExecutionPatchRequest(BaseModel):
    execution: Optional[Any] = None
    progressLog: Optional[Any] = None


@router.get("", response_model=list[ConversationSummary])
async def list_all() -> list[ConversationSummary]:
    rows = await list_conversations(limit=100)
    return [ConversationSummary(**r) for r in rows]


@router.get("/{conversation_id}", response_model=ConversationDetail)
async def get_one(conversation_id: str) -> ConversationDetail:
    conv = await get_conversation(conversation_id)
    if not conv:
        raise HTTPException(status_code=404, detail="Conversation not found")
    return ConversationDetail(**conv)


@router.patch("/{conversation_id}", response_model=ConversationSummary)
async def rename_one(conversation_id: str, body: RenameRequest) -> ConversationSummary:
    ok = await rename_conversation(conversation_id, body.title)
    if not ok:
        raise HTTPException(status_code=404, detail="Conversation not found")
    # Return refreshed summary
    rows = await list_conversations(limit=1000)
    for r in rows:
        if r["id"] == conversation_id:
            return ConversationSummary(**r)
    raise HTTPException(status_code=404, detail="Conversation not found")


@router.delete("/{conversation_id}")
async def delete_one(conversation_id: str) -> dict[str, bool]:
    ok = await delete_conversation(conversation_id)
    if not ok:
        raise HTTPException(status_code=404, detail="Conversation not found")
    return {"deleted": True}


@router.patch("/{conversation_id}/messages/{message_id}")
async def patch_message_execution(
    conversation_id: str, message_id: str, body: ExecutionPatchRequest,
) -> dict[str, bool]:
    """Attach execution state + progress log to a specific message."""
    ok = await update_conv_message_execution(
        conversation_id=conversation_id,
        message_id=message_id,
        execution=body.execution,
        progress_log=body.progressLog,
    )
    if not ok:
        raise HTTPException(status_code=404, detail="Message not found")
    return {"updated": True}


@router.delete("/{conversation_id}/last")
async def pop_last(conversation_id: str) -> dict[str, int]:
    """Remove the last user+assistant exchange (used by undo)."""
    removed = await pop_last_exchange(conversation_id)
    return {"removed": removed}
