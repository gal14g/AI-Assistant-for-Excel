"""
Deterministic Python tools for spreadsheet data operations.

All tools:
- Accept typed inputs (SheetData, column lists, config)
- Return ToolOutput with success/error state
- Never modify the original data (immutable)
- Are fully synchronous (no I/O, no LLM)
"""

from .sheet_tools import list_sheets, get_sheet_schema, preview_sheet, profile_columns
from .cleaning_tools import clean_columns
from .matching_tools import (
    estimate_matchability,
    run_exact_match,
    run_fuzzy_match,
    run_hybrid_match,
    run_semantic_match,
)
from .aggregation_tools import aggregate_values, find_duplicates
from .comparison_tools import compare_sheets, filter_rows, explain_match_result

__all__ = [
    "list_sheets",
    "get_sheet_schema",
    "preview_sheet",
    "profile_columns",
    "clean_columns",
    "estimate_matchability",
    "run_exact_match",
    "run_fuzzy_match",
    "run_hybrid_match",
    "run_semantic_match",
    "aggregate_values",
    "find_duplicates",
    "compare_sheets",
    "filter_rows",
    "explain_match_result",
]
