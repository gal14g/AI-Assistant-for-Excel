"""
Aggregation and duplicate-detection tools.
"""
from __future__ import annotations

from ..models.analytical_plan import SheetData
from ..models.tool_output import ToolOutput

_TOOL_AGG = "aggregate_values"
_TOOL_DUP = "find_duplicates"


def aggregate_values(
    sheet: SheetData,
    group_by_columns: list[str],
    agg_column: str,
    agg_function: str = "sum",
) -> ToolOutput:
    """
    Group the sheet by *group_by_columns* and apply *agg_function* to
    *agg_column*.

    Supported functions: sum, mean, min, max, count, median.
    """
    try:
        import pandas as pd
    except ImportError:
        return ToolOutput.fail(
            tool_name=_TOOL_AGG,
            errors=["pandas is required for aggregate_values."],
        )

    df = sheet.to_dataframe()
    warnings: list[str] = []

    for col in group_by_columns:
        if col not in df.columns:
            return ToolOutput.fail(
                tool_name=_TOOL_AGG,
                errors=[f"Group-by column '{col}' not found in sheet '{sheet.name}'."],
            )
    if agg_column not in df.columns:
        return ToolOutput.fail(
            tool_name=_TOOL_AGG,
            errors=[f"Aggregation column '{agg_column}' not found in sheet '{sheet.name}'."],
        )

    # Coerce agg column to numeric for numeric aggregations
    numeric_funcs = {"sum", "mean", "min", "max", "median"}
    if agg_function in numeric_funcs:
        df[agg_column] = pd.to_numeric(df[agg_column], errors="coerce")
        null_count = df[agg_column].isna().sum()
        if null_count > 0:
            warnings.append(
                f"{null_count} value(s) in '{agg_column}' could not be converted to numeric and were ignored."
            )

    func_map = {
        "sum": "sum",
        "mean": "mean",
        "min": "min",
        "max": "max",
        "count": "count",
        "median": "median",
    }

    if agg_function not in func_map:
        return ToolOutput.fail(
            tool_name=_TOOL_AGG,
            errors=[
                f"Unknown aggregation function '{agg_function}'. "
                f"Supported: {', '.join(func_map.keys())}"
            ],
        )

    grouped = df.groupby(group_by_columns, as_index=False)[agg_column].agg(
        func_map[agg_function]
    )
    result_rows = grouped.where(grouped.notna(), other=None).to_dict(orient="records")

    return ToolOutput.ok(
        tool_name=_TOOL_AGG,
        data={
            "result": result_rows,
            "group_by": group_by_columns,
            "agg_column": agg_column,
            "agg_function": agg_function,
            "group_count": len(result_rows),
        },
        warnings=warnings,
    )


def find_duplicates(
    sheet: SheetData,
    key_columns: list[str] | None = None,
    keep: str = "first",
) -> ToolOutput:
    """
    Identify duplicate rows in a sheet.

    key_columns: columns to compare for duplicates (defaults to all columns).
    keep: 'first' marks all subsequent duplicates; 'last' marks all preceding;
          'none' marks all occurrences.
    """
    try:
        import pandas as pd
    except ImportError:
        return ToolOutput.fail(
            tool_name=_TOOL_DUP,
            errors=["pandas is required for find_duplicates."],
        )

    df = sheet.to_dataframe()
    warnings: list[str] = []

    subset = key_columns if key_columns else None
    if subset:
        missing = [c for c in subset if c not in df.columns]
        if missing:
            return ToolOutput.fail(
                tool_name=_TOOL_DUP,
                errors=[f"Column(s) not found: {missing}"],
            )

    keep_param: str | bool = keep if keep in ("first", "last") else False
    dup_mask = df.duplicated(subset=subset, keep=keep_param)  # type: ignore[arg-type]

    duplicates = df[dup_mask].where(pd.notnull(df[dup_mask]), other=None).to_dict(orient="records")
    unique = df[~dup_mask].where(pd.notnull(df[~dup_mask]), other=None).to_dict(orient="records")

    total_dup_groups = df[df.duplicated(subset=subset, keep=False)].shape[0]

    if total_dup_groups == 0:
        warnings.append("No duplicates found.")

    return ToolOutput.ok(
        tool_name=_TOOL_DUP,
        data={
            "duplicates": duplicates,
            "unique_rows": unique,
            "duplicate_count": len(duplicates),
            "unique_count": len(unique),
            "total_rows": len(df),
            "key_columns": key_columns,
        },
        warnings=warnings,
    )
