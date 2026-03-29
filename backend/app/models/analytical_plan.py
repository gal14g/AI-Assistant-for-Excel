"""
Pydantic models for the analytical pipeline (Layer 1 planner output).

These are distinct from ExecutionPlan (Office.js operations) — they
describe Python-side data analysis operations.
"""
from __future__ import annotations

from enum import Enum
from typing import Any, Optional
from pydantic import BaseModel, Field


class IntentType(str, Enum):
    answer_question = "answer_question"
    ask_clarification = "ask_clarification"
    filter_rows = "filter_rows"
    aggregate = "aggregate"
    group_and_summarize = "group_and_summarize"
    match_rows = "match_rows"
    find_duplicates = "find_duplicates"
    clean_data = "clean_data"
    compare_sheets = "compare_sheets"
    semantic_lookup = "semantic_lookup"
    profile_sheet = "profile_sheet"


class StrategyType(str, Enum):
    exact = "exact"
    fuzzy = "fuzzy"
    semantic = "semantic"
    hybrid = "hybrid"


class OperationType(str, Enum):
    list_sheets = "list_sheets"
    get_sheet_schema = "get_sheet_schema"
    preview_sheet = "preview_sheet"
    profile_columns = "profile_columns"
    clean_columns = "clean_columns"
    estimate_matchability = "estimate_matchability"
    run_exact_match = "run_exact_match"
    run_fuzzy_match = "run_fuzzy_match"
    run_semantic_match = "run_semantic_match"
    run_hybrid_match = "run_hybrid_match"
    find_duplicates = "find_duplicates"
    aggregate_values = "aggregate_values"
    filter_rows = "filter_rows"
    compare_sheets = "compare_sheets"
    explain_match_result = "explain_match_result"


class AnalyticalPlan(BaseModel):
    intent: IntentType
    confidence: float = Field(ge=0.0, le=1.0)
    needs_clarification: bool = False
    clarification_question: Optional[str] = None
    selected_tool_chain: list[OperationType]
    parameters: dict[str, Any] = Field(default_factory=dict)
    reasoning_summary: str


class ColumnProfile(BaseModel):
    name: str
    dtype: str  # "numeric" | "text" | "id" | "date" | "mixed" | "unknown"
    null_rate: float  # 0.0–1.0
    distinct_count: int
    uniqueness_ratio: float  # 0.0–1.0
    avg_text_length: Optional[float] = None
    sample_values: list[Any] = Field(default_factory=list)
    match_weight: float = 1.0
    suggested_strategy: Optional[StrategyType] = None
    warnings: list[str] = Field(default_factory=list)


class MatchabilityEstimate(BaseModel):
    overall_confidence: float
    recommended_strategy: StrategyType
    column_scores: dict[str, float]
    column_strategies: dict[str, StrategyType]
    warnings: list[str]
    needs_more_columns: bool
    suggested_additional_columns: list[str]


class MatchResult(BaseModel):
    left_indices: list[int]
    right_indices: list[int]
    scores: list[float]
    match_count: int
    unmatched_left: int
    unmatched_right: int
    strategy_used: StrategyType
    column_contributions: dict[str, list[float]] = Field(default_factory=dict)


class SheetData(BaseModel):
    """Sheet data sent from the frontend (Office.js readRange result)."""
    name: str
    data: list[list[Any]]  # raw 2D cell values
    headers: Optional[list[str]] = None  # if None, first data row is headers

    @property
    def header_row(self) -> list[str]:
        if self.headers:
            return self.headers
        if self.data:
            return [str(v) if v is not None else "" for v in self.data[0]]
        return []

    @property
    def data_rows(self) -> list[list[Any]]:
        if self.headers:
            return self.data
        return self.data[1:] if len(self.data) > 1 else []

    def to_dataframe(self):  # -> pd.DataFrame
        import pandas as pd
        headers = self.header_row
        rows = self.data_rows
        if not rows:
            return pd.DataFrame(columns=headers)
        # Pad short rows to match header count
        col_count = len(headers)
        padded = [
            r + [None] * (col_count - len(r)) if len(r) < col_count else r[:col_count]
            for r in rows
        ]
        return pd.DataFrame(padded, columns=headers)
