"""
Sheet comparison, row filtering, and match-result explanation tools.
"""
from __future__ import annotations

from typing import Any

from ..models.analytical_plan import SheetData
from ..models.tool_output import ToolOutput

_TOOL_COMPARE = "compare_sheets"
_TOOL_FILTER = "filter_rows"
_TOOL_EXPLAIN = "explain_match_result"


def compare_sheets(
    left: SheetData,
    right: SheetData,
    key_column_left: str,
    key_column_right: str,
) -> ToolOutput:
    """
    Compare two sheets row-by-row on a shared key.

    Returns:
      - only_in_left: rows whose key appears only in the left sheet
      - only_in_right: rows whose key appears only in the right sheet
      - in_both: rows present in both (by normalised key)
    """
    try:
        import pandas as pd
    except ImportError:
        return ToolOutput.fail(
            tool_name=_TOOL_COMPARE,
            errors=["pandas is required for compare_sheets."],
        )

    left_df = left.to_dataframe()
    right_df = right.to_dataframe()

    if key_column_left not in left_df.columns:
        return ToolOutput.fail(
            tool_name=_TOOL_COMPARE,
            errors=[f"Key column '{key_column_left}' not found in sheet '{left.name}'."],
        )
    if key_column_right not in right_df.columns:
        return ToolOutput.fail(
            tool_name=_TOOL_COMPARE,
            errors=[f"Key column '{key_column_right}' not found in sheet '{right.name}'."],
        )

    left_keys = set(left_df[key_column_left].astype(str).str.lower().str.strip())
    right_keys = set(right_df[key_column_right].astype(str).str.lower().str.strip())

    left_df = left_df.copy()
    right_df = right_df.copy()
    left_df["_norm"] = left_df[key_column_left].astype(str).str.lower().str.strip()
    right_df["_norm"] = right_df[key_column_right].astype(str).str.lower().str.strip()

    only_left = left_df[~left_df["_norm"].isin(right_keys)].drop(columns=["_norm"])
    only_right = right_df[~right_df["_norm"].isin(left_keys)].drop(columns=["_norm"])
    in_both_left = left_df[left_df["_norm"].isin(right_keys)].drop(columns=["_norm"])

    def _to_records(df: Any) -> list[dict]:
        return df.where(pd.notnull(df), other=None).to_dict(orient="records")

    return ToolOutput.ok(
        tool_name=_TOOL_COMPARE,
        data={
            "only_in_left": _to_records(only_left),
            "only_in_right": _to_records(only_right),
            "in_both": _to_records(in_both_left),
            "only_left_count": len(only_left),
            "only_right_count": len(only_right),
            "in_both_count": len(in_both_left),
        },
    )


def filter_rows(
    sheet: SheetData,
    column: str,
    operator: str,
    value: Any,
) -> ToolOutput:
    """
    Filter rows in a sheet based on a simple condition.

    operator: eq, ne, gt, lt, gte, lte, contains, startswith, endswith
    """
    try:
        import pandas as pd
    except ImportError:
        return ToolOutput.fail(
            tool_name=_TOOL_FILTER,
            errors=["pandas is required for filter_rows."],
        )

    df = sheet.to_dataframe()

    if column not in df.columns:
        return ToolOutput.fail(
            tool_name=_TOOL_FILTER,
            errors=[f"Column '{column}' not found in sheet '{sheet.name}'."],
        )

    series = df[column]
    value_str = str(value).lower()

    op_map = {
        "eq": lambda s: s.astype(str).str.lower() == value_str,
        "ne": lambda s: s.astype(str).str.lower() != value_str,
        "contains": lambda s: s.astype(str).str.lower().str.contains(value_str, na=False),
        "startswith": lambda s: s.astype(str).str.lower().str.startswith(value_str, na=False),
        "endswith": lambda s: s.astype(str).str.lower().str.endswith(value_str, na=False),
    }

    numeric_ops = {"gt", "lt", "gte", "lte"}
    if operator in numeric_ops:
        numeric_series = pd.to_numeric(series, errors="coerce")
        try:
            float_val = float(value)
        except (TypeError, ValueError):
            return ToolOutput.fail(
                tool_name=_TOOL_FILTER,
                errors=[f"Cannot compare '{column}' with non-numeric value '{value}' using operator '{operator}'."],
            )
        if operator == "gt":
            mask = numeric_series > float_val
        elif operator == "lt":
            mask = numeric_series < float_val
        elif operator == "gte":
            mask = numeric_series >= float_val
        else:
            mask = numeric_series <= float_val
    elif operator in op_map:
        mask = op_map[operator](series)
    else:
        return ToolOutput.fail(
            tool_name=_TOOL_FILTER,
            errors=[f"Unknown operator '{operator}'."],
        )

    filtered = df[mask].where(pd.notnull(df[mask]), other=None).to_dict(orient="records")

    return ToolOutput.ok(
        tool_name=_TOOL_FILTER,
        data={
            "rows": filtered,
            "row_count": len(filtered),
            "total_rows": len(df),
            "filter": {"column": column, "operator": operator, "value": value},
        },
    )


def explain_match_result(
    match_data: dict,
    left_name: str,
    right_name: str,
) -> ToolOutput:
    """
    Generate a human-readable summary of a match result dict.
    """
    match_count = match_data.get("match_count", 0)
    unmatched_left = match_data.get("unmatched_left_count", 0)
    unmatched_right = match_data.get("unmatched_right_count", 0)
    strategy = match_data.get("match_result", {}).get("strategy_used", "unknown")

    lines = [
        f"Matched {match_count} rows between '{left_name}' and '{right_name}' "
        f"using {strategy} strategy.",
    ]
    if unmatched_left:
        lines.append(
            f"{unmatched_left} row(s) in '{left_name}' had no match in '{right_name}'."
        )
    if unmatched_right:
        lines.append(
            f"{unmatched_right} row(s) in '{right_name}' had no match in '{left_name}'."
        )
    if match_count == 0:
        lines.append(
            "No rows matched. Consider reviewing the key columns or lowering the fuzzy threshold."
        )

    return ToolOutput.ok(
        tool_name=_TOOL_EXPLAIN,
        data={
            "summary": " ".join(lines),
            "lines": lines,
            "match_count": match_count,
            "unmatched_left": unmatched_left,
            "unmatched_right": unmatched_right,
        },
    )
