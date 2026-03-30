"""
Request/response models for the API endpoints.
"""

from __future__ import annotations

from typing import Optional
from pydantic import BaseModel

from .plan import ExecutionPlan


class RangeTokenRef(BaseModel):
    address: str
    sheetName: str


class ConversationMessage(BaseModel):
    role: str
    content: str


class PlanRequest(BaseModel):
    """Request body for POST /api/plan"""

    userMessage: str
    rangeTokens: Optional[list[RangeTokenRef]] = None
    activeSheet: Optional[str] = None
    # Workbook identity – used to qualify cross-sheet and cross-file references
    workbookName: Optional[str] = None
    workbookPath: Optional[str] = None
    conversationHistory: Optional[list[ConversationMessage]] = None


class PlanResponse(BaseModel):
    """Response body for POST /api/plan"""

    plan: ExecutionPlan
    explanation: str
    alternatives: Optional[list[str]] = None


class ValidationIssue(BaseModel):
    message: str
    code: str
    stepId: Optional[str] = None
    field: Optional[str] = None


class ValidationResponse(BaseModel):
    """Response body for POST /api/validate"""

    valid: bool
    errors: list[ValidationIssue]
    warnings: list[ValidationIssue]


class CapabilityInfo(BaseModel):
    action: str
    description: str
    mutates: bool
    affectsFormatting: bool
