"""
Chat request/response models for POST /api/chat.

The chat endpoint uses a single LLM call that either returns a
conversational reply (responseType="message") or multiple execution
plan options (responseType="plans") for the user to choose from.
"""

from __future__ import annotations

from typing import Literal, Optional
from pydantic import BaseModel, Field

from .plan import ExecutionPlan
from .request import RangeTokenRef, ConversationMessage


class ChatRequest(BaseModel):
    userMessage: str = Field(..., min_length=1, max_length=5000)
    rangeTokens: Optional[list[RangeTokenRef]] = Field(None, max_length=50)
    activeSheet: Optional[str] = Field(None, max_length=255)
    workbookName: Optional[str] = Field(None, max_length=260)
    usedRangeEnd: Optional[str] = Field(None, max_length=20)
    locale: Optional[str] = Field(None, max_length=10)
    conversationHistory: Optional[list[ConversationMessage]] = Field(None, max_length=20)


class PlanOption(BaseModel):
    optionLabel: str       # e.g. "Option A: Use SUMIF formulas"
    plan: ExecutionPlan


class ChatResponse(BaseModel):
    responseType: Literal["message", "plan", "plans"]
    message: str
    plan: Optional[ExecutionPlan] = None           # single plan (backward compat)
    plans: Optional[list[PlanOption]] = None        # multiple options
    interactionId: Optional[str] = None             # DB interaction ID for feedback
