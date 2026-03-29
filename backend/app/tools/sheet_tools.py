"""
Sheet inspection tools: list_sheets, get_sheet_schema, preview_sheet, profile_columns.

All functions are pure / synchronous — no I/O, no LLM calls.
"""
from __future__ import annotations

import re
from typing import Any

import pandas as pd

from ..models.analytical_plan import ColumnProfile, SheetData, StrategyType
from ..models.tool_output import ToolOutput

_TOOL = "sheet_tools"

# ── Null sentinels treated as "missing" during profiling ──────────────────────
_NULL_LIKE = {"", "n/a", "na", "null", "none", "-", "--", "nan"}


# ── Helpers ───────────────────────────────────────────────────────────────────


def _is_numeric_series(s: pd.Series) -> bool:
    """Return True if >80 % of non-null values parse as float."""
    non_null = s.dropna()
    if non_null.empty:
        return False
    def _try_float(v: Any) -> bool:
        try:
            float(str(v).replace(",", "").replace("_", ""))
            return True
        except (ValueError, TypeError):
            return False
    return non_null.map(_try_float).mean() >= 0.80


def _is_date_series(s: pd.Series) -> bool:
    """Return True if >70 % of non-null values parse as a date."""
    non_null = s.dropna()
    if non_null.empty:
        return False
    converted = pd.to_datetime(non_null.astype(str), errors="coerce")
    return converted.notna().mean() >= 0.70


_ID_PATTERN = re.compile(r"^[A-Za-z0-9_\-\.]{1,40}$")


def _is_id_series(s: pd.Series, uniqueness_ratio: float) -> bool:
    """
    Heuristic: short strings, high uniqueness and matches a simple id pattern.
    """
    if uniqueness_ratio < 0.70:
        return False
    non_null = s.dropna()
    if non_null.empty:
        return False
    avg_len = non_null.astype(str).map(len).mean()
    if avg_len > 30:
        return False
    matches = non_null.astype(str).map(lambda v: bool(_ID_PATTERN.match(v)))
    return matches.mean() >= 0.80


def _detect_dtype(s: pd.Series, uniqueness_ratio: float) -> str:
    """Classify a column series into one of the canonical dtype labels."""
    non_null = s.dropna()
    if non_null.empty:
        return "unknown"

    # 1. Numeric
    if _is_numeric_series(s):
        return "numeric"

    # 2. Date
    if _is_date_series(s):
        return "date"

    # 3. ID — checked before generic text because IDs are short/high-uniqueness
    if _is_id_series(s, uniqueness_ratio):
        return "id"

    # 4. Text vs mixed: check whether the majority of values are plain strings
    str_vals = non_null.astype(str)
    avg_len = str_vals.map(len).mean()
    if avg_len <= 60:
        # Could be short categorical text or IDs that didn't pass the pattern
        return "text"

    # 5. Mixed if we can't decide confidently
    return "mixed"


def _suggested_strategy(dtype: str, uniqueness_ratio: float) -> StrategyType:
    if dtype == "id":
        return StrategyType.exact
    if dtype == "numeric" or dtype == "date":
        return StrategyType.exact
    if uniqueness_ratio > 0.80:
        return StrategyType.fuzzy
    return StrategyType.semantic


# ── Public functions ──────────────────────────────────────────────────────────


def list_sheets(sheets: dict[str, SheetData]) -> ToolOutput:
    """Return a list of sheet names with their row and column counts."""
    result = []
    for name, sheet in sheets.items():
        df = sheet.to_dataframe()
        result.append(
            {
                "name": name,
                "row_count": len(df),
                "column_count": len(df.columns),
                "columns": list(df.columns),
            }
        )
    return ToolOutput.ok(
        tool_name="list_sheets",
        data={"sheets": result, "sheet_count": len(result)},
    )


def get_sheet_schema(sheet: SheetData) -> ToolOutput:
    """Return column names, row count, and column count for a sheet."""
    df = sheet.to_dataframe()
    return ToolOutput.ok(
        tool_name="get_sheet_schema",
        data={
            "name": sheet.name,
            "columns": list(df.columns),
            "row_count": len(df),
            "column_count": len(df.columns),
        },
    )


def preview_sheet(sheet: SheetData, n_rows: int = 20) -> ToolOutput:
    """Return the first *n_rows* rows as a list of dicts (column → value)."""
    df = sheet.to_dataframe()
    preview = df.head(n_rows).where(pd.notnull(df), other=None)
    records = preview.to_dict(orient="records")
    return ToolOutput.ok(
        tool_name="preview_sheet",
        data={
            "name": sheet.name,
            "rows": records,
            "returned_rows": len(records),
            "total_rows": len(df),
        },
    )


def profile_columns(sheet: SheetData, columns: list[str]) -> ToolOutput:
    """
    Profile each requested column and return a list of ColumnProfile dicts.

    For each column:
      - dtype: numeric | date | id | text | mixed | unknown
      - null_rate, distinct_count, uniqueness_ratio
      - avg_text_length (for non-numeric)
      - sample_values (up to 5 non-null values)
      - warnings: high null rate, low/high uniqueness
      - suggested_strategy
    """
    df = sheet.to_dataframe()
    available = set(df.columns)
    profiles: list[dict] = []
    tool_warnings: list[str] = []

    for col in columns:
        if col not in available:
            tool_warnings.append(f"Column '{col}' not found in sheet '{sheet.name}' — skipped.")
            continue

        series = df[col]
        total = len(series)

        # Treat null-like string sentinels as missing
        def _is_null(v: Any) -> bool:
            if v is None:
                return True
            return str(v).strip().lower() in _NULL_LIKE

        null_mask = series.map(_is_null)
        null_rate = float(null_mask.mean()) if total > 0 else 1.0
        non_null_series = series[~null_mask]

        distinct_count = int(non_null_series.nunique())
        non_null_count = len(non_null_series)
        uniqueness_ratio = (
            float(distinct_count / non_null_count) if non_null_count > 0 else 0.0
        )

        dtype = _detect_dtype(non_null_series, uniqueness_ratio)

        # avg_text_length only for non-numeric / non-date
        avg_text_length: float | None = None
        if dtype not in ("numeric",):
            lengths = non_null_series.dropna().astype(str).map(len)
            avg_text_length = float(lengths.mean()) if not lengths.empty else None

        # Sample values — up to 5 unique, non-null
        sample_values: list[Any] = (
            non_null_series.dropna()
            .drop_duplicates()
            .head(5)
            .tolist()
        )

        warnings: list[str] = []
        if null_rate > 0.3:
            warnings.append(
                f"High null rate ({null_rate:.0%}) — column may be unreliable for matching."
            )
        if uniqueness_ratio < 0.05 and non_null_count > 0:
            warnings.append(
                f"Very low uniqueness ({uniqueness_ratio:.2%}) — likely a categorical column."
            )
        if uniqueness_ratio > 0.95 and non_null_count > 0:
            warnings.append(
                f"Very high uniqueness ({uniqueness_ratio:.2%}) — potential key/ID column."
            )

        strategy = _suggested_strategy(dtype, uniqueness_ratio)

        profile = ColumnProfile(
            name=col,
            dtype=dtype,
            null_rate=null_rate,
            distinct_count=distinct_count,
            uniqueness_ratio=uniqueness_ratio,
            avg_text_length=avg_text_length,
            sample_values=sample_values,
            match_weight=1.0,
            suggested_strategy=strategy,
            warnings=warnings,
        )
        profiles.append(profile.model_dump())

    return ToolOutput.ok(
        tool_name="profile_columns",
        data={"profiles": profiles, "column_count": len(profiles)},
        warnings=tool_warnings,
    )
