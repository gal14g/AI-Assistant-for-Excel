"""
Text / data cleaning tools.
"""
from __future__ import annotations

import re

from ..models.analytical_plan import SheetData
from ..models.tool_output import ToolOutput

_TOOL = "clean_columns"


def clean_columns(
    sheet: SheetData,
    columns: list[str],
    operations: list[str],
) -> ToolOutput:
    """
    Apply a sequence of cleaning operations to selected columns.

    Supported operations:
      trim, lowercase, uppercase, proper_case, remove_punctuation,
      normalize_whitespace, strip_leading_zeros, remove_non_ascii
    """
    try:
        import pandas as pd
    except ImportError:
        return ToolOutput.fail(
            tool_name=_TOOL,
            errors=["pandas is required for clean_columns."],
        )

    df = sheet.to_dataframe().copy()
    warnings: list[str] = []
    available = set(df.columns)

    for col in columns:
        if col not in available:
            warnings.append(f"Column '{col}' not found — skipped.")
            continue

        series = df[col].astype(str)

        for op in operations:
            if op == "trim":
                series = series.str.strip()
            elif op == "lowercase":
                series = series.str.lower()
            elif op == "uppercase":
                series = series.str.upper()
            elif op == "proper_case":
                series = series.str.title()
            elif op == "remove_punctuation":
                series = series.map(lambda v: re.sub(r"[^\w\s]", "", v))
            elif op == "normalize_whitespace":
                series = series.map(lambda v: re.sub(r"\s+", " ", v).strip())
            elif op == "strip_leading_zeros":
                series = series.str.lstrip("0")
            elif op == "remove_non_ascii":
                series = series.map(lambda v: v.encode("ascii", "ignore").decode("ascii"))
            else:
                warnings.append(f"Unknown operation '{op}' — skipped.")

        df[col] = series

    cleaned_rows = [df.columns.tolist()] + df.where(pd.notnull(df), other=None).values.tolist()

    return ToolOutput.ok(
        tool_name=_TOOL,
        data={
            "cleaned_data": cleaned_rows,
            "rows_processed": len(df),
            "columns_cleaned": [c for c in columns if c in available],
        },
        warnings=warnings,
    )
