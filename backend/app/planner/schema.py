"""
Re-exports and planner-specific context types.

The planner builds SheetContext objects from available SheetData and injects
them into the LLM prompt so the planner has accurate schema knowledge.
"""
from __future__ import annotations

from typing import Any
from pydantic import BaseModel

# Re-export the core analytical types for callers who import from here
from ..models.analytical_plan import (
    AnalyticalPlan,
    IntentType,
    StrategyType,
    OperationType,
    ColumnProfile,
    MatchabilityEstimate,
    MatchResult,
    SheetData,
)

__all__ = [
    "AnalyticalPlan",
    "IntentType",
    "StrategyType",
    "OperationType",
    "ColumnProfile",
    "MatchabilityEstimate",
    "MatchResult",
    "SheetData",
    "SheetContext",
    "PlannerContext",
]


class SheetContext(BaseModel):
    """Lightweight schema snapshot of one sheet — injected into the LLM prompt."""

    name: str
    columns: list[str]
    row_count: int
    column_profiles: list[dict[str, Any]] = []

    def describe(self) -> str:
        """Return a compact one-line description suitable for prompt injection."""
        cols = ", ".join(self.columns[:20])
        suffix = f" … (+{len(self.columns) - 20} more)" if len(self.columns) > 20 else ""
        return f"  '{self.name}': {self.row_count} rows | columns: [{cols}{suffix}]"


class PlannerContext(BaseModel):
    """Full context passed to the LLM planner."""

    user_message: str
    available_sheets: list[SheetContext]
    conversation_summary: str = ""

    def sheet_descriptions(self) -> str:
        return "\n".join(sc.describe() for sc in self.available_sheets)
