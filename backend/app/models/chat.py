"""
Chat request/response models for POST /api/chat.

The chat endpoint uses a single LLM call that either returns a
conversational reply (responseType="message") or multiple execution
plan options (responseType="plans") for the user to choose from.
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


class PlanOption(BaseModel):
    optionLabel: str       # e.g. "Option A: Use SUMIF formulas"
    plan: ExecutionPlan


class ChatResponse(BaseModel):
    responseType: Literal["message", "plan", "plans"]
    message: str
    plan: Optional[ExecutionPlan] = None           # single plan (backward compat)
    plans: Optional[list[PlanOption]] = None        # multiple options
    interactionId: Optional[str] = None             # DB interaction ID for feedback
