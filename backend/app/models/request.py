"""
Shared request/response models.
"""

from __future__ import annotations

from typing import Optional
from pydantic import BaseModel


class RangeTokenRef(BaseModel):
    address: str
    sheetName: str


class ConversationMessage(BaseModel):
    role: str
    content: str


class ValidationIssue(BaseModel):
    message: str
    code: str
    stepId: Optional[str] = None
    field: Optional[str] = None


class ValidationResponse(BaseModel):
    valid: bool
    errors: list[ValidationIssue]
    warnings: list[ValidationIssue]
