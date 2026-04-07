"""
Chat request/response models for POST /api/chat.

The chat endpoint uses a single LLM call that either returns a
conversational reply (responseType="message") or multiple execution
plan options (responseType="plans") for the user to choose from.
"""

from __future__ import annotations

from typing import Literal, Optional, Union
from pydantic import BaseModel, Field

from .plan import ExecutionPlan
from .request import RangeTokenRef, ConversationMessage

# Cell values from the frontend come in as strings, numbers, booleans, or null.
SnapshotCell = Optional[Union[str, int, float, bool]]


class SheetSnapshot(BaseModel):
    """Per-sheet snapshot of the workbook's real data."""
    sheetName: str = Field(..., max_length=255)
    rowCount: int = Field(0, ge=0)
    columnCount: int = Field(0, ge=0)
    headers: list[str] = Field(default_factory=list, max_length=100)
    # Sample rows (first 5 data rows after the header).
    sampleRows: list[list[SnapshotCell]] = Field(default_factory=list, max_length=10)
    # Inferred dtypes per column: "number" | "date" | "text" | "boolean" | "mixed" | "empty"
    dtypes: list[str] = Field(default_factory=list, max_length=100)
    # Top-left cell of the used range (e.g. "A1", "C5") — tables don't always start at A1.
    anchorCell: str = Field("A1", max_length=20)
    # Full used-range address (e.g. "Sheet1!C5:H40").
    usedRangeAddress: str = Field("", max_length=100)


class WorkbookSnapshot(BaseModel):
    """Lightweight snapshot of the whole workbook — lets the planner see
    real column names, dtypes, and sample values instead of guessing."""
    sheets: list[SheetSnapshot] = Field(default_factory=list, max_length=20)
    truncated: bool = False


class ChatRequest(BaseModel):
    userMessage: str = Field(..., min_length=1, max_length=5000)
    rangeTokens: Optional[list[RangeTokenRef]] = Field(None, max_length=50)
    activeSheet: Optional[str] = Field(None, max_length=255)
    workbookName: Optional[str] = Field(None, max_length=260)
    usedRangeEnd: Optional[str] = Field(None, max_length=20)
    locale: Optional[str] = Field(None, max_length=10)
    conversationHistory: Optional[list[ConversationMessage]] = Field(None, max_length=20)
    # Persisted-conversation plumbing. All optional: when omitted a new
    # conversation is created server-side and returned in the response.
    conversationId: Optional[str] = Field(None, max_length=64)
    userMessageId: Optional[str] = Field(None, max_length=64)
    workbookSnapshot: Optional[WorkbookSnapshot] = None


class PlanOption(BaseModel):
    optionLabel: str       # e.g. "Option A: Use SUMIF formulas"
    plan: ExecutionPlan


class ChatResponse(BaseModel):
    responseType: Literal["message", "plan", "plans"]
    message: str
    plan: Optional[ExecutionPlan] = None           # single plan (backward compat)
    plans: Optional[list[PlanOption]] = None        # multiple options
    interactionId: Optional[str] = None             # DB interaction ID for feedback
    conversationId: Optional[str] = None            # Persistent conversation id
    assistantMessageId: Optional[str] = None        # ID of the stored assistant message
