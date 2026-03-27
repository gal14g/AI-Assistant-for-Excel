"""
Chat request/response models for POST /api/chat.

The chat endpoint uses a single LLM call that either returns a
conversational reply (responseType="message") or a full execution
plan (responseType="plan"), depending on what the user asked.
"""

from __future__ import annotations

from typing import Literal, Optional
from pydantic import BaseModel

from .plan import ExecutionPlan
from .request import RangeTokenRef, ConversationMessage


class ChatRequest(BaseModel):
    userMessage: str
    rangeTokens: Optional[list[RangeTokenRef]] = None
    activeSheet: Optional[str] = None
    workbookName: Optional[str] = None
    conversationHistory: Optional[list[ConversationMessage]] = None


class ChatResponse(BaseModel):
    responseType: Literal["message", "plan"]
    message: str
    plan: Optional[ExecutionPlan] = None
