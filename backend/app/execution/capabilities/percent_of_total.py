"""percentOfTotal — write =Cn/SUM($C$first:$C$last) for each row.

Produces a formula column showing each value as a fraction of the total.
When formatAsPercent is true (default), applies "0.0%" number format to the
output range.
"""

from __future__ import annotations

import re
from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


_COL_RE = re.compile(r"[A-Z]+", re.IGNORECASE)


def _column_letter(ref: str) -> str:
    stripped = ref.replace("$", "")
    if "!" in stripped:
        stripped = stripped.split("!", 1)[1]
    m = _COL_RE.search(stripped)
    return m.group(0).upper() if m else "A"


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    source = params.get("sourceRange")
    output = params.get("outputRange")
    has_headers = bool(params.get("hasHeaders", True))
    format_as_percent = bool(params.get("formatAsPercent", True))

    if not source or not output:
        return {
            "status": "error",
            "message": "percentOfTotal requires 'sourceRange' and 'outputRange'.",
        }

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would write percent-of-total formulas from {source} to {output}.",
        }

    try:
        src_rng = resolve_range(ctx.workbook_handle, source)
        sheet = src_rng.sheet

        try:
            used = src_rng.current_region
            start_row = used.row
            last_row = start_row + used.rows.count - 1
        except Exception:
            used = sheet.used_range
            last_row = used.row + used.rows.count - 1

        first_data_row = 2 if has_headers else 1
        if last_row < first_data_row:
            return {"status": "success", "message": "No data rows found.", "outputs": {}}

        src_col = _column_letter(source)
        out_col = _column_letter(output)

        if has_headers:
            sheet.range(f"{out_col}1").value = "% of Total"

        formulas = [
            [f"={src_col}{r}/SUM(${src_col}${first_data_row}:${src_col}${last_row})"]
            for r in range(first_data_row, last_row + 1)
        ]
        out_rng = sheet.range(f"{out_col}{first_data_row}:{out_col}{last_row}")
        out_rng.formula = formulas

        if format_as_percent:
            out_rng.number_format = "0.0%"

        row_count = last_row - first_data_row + 1
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"percentOfTotal failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Created {row_count} percent-of-total formulas in {output}.",
        "outputs": {"outputRange": output},
    }


registry.register("percentOfTotal", handler, mutates=True, affects_formatting=True)
