"""runningTotal — write a running (cumulative) SUM formula column.

For each data row, writes =SUM($<srcCol>$<firstDataRow>:<srcCol><row>) so the
total accumulates as you scroll down. Detects the last used row of the source
range and preserves the header row when hasHeaders is true.
"""

from __future__ import annotations

import re
from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


_COL_RE = re.compile(r"[A-Z]+", re.IGNORECASE)


def _column_letter(ref: str) -> str:
    """Return the first column-letter chunk from an A1 reference."""
    stripped = ref.replace("$", "")
    if "!" in stripped:
        stripped = stripped.split("!", 1)[1]
    m = _COL_RE.search(stripped)
    return m.group(0).upper() if m else "A"


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    source = params.get("sourceRange")
    output = params.get("outputRange")
    has_headers = bool(params.get("hasHeaders", True))

    if not source or not output:
        return {
            "status": "error",
            "message": "runningTotal requires 'sourceRange' and 'outputRange'.",
        }

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would write running total formulas from {source} to {output}.",
        }

    try:
        src_rng = resolve_range(ctx.workbook_handle, source)
        sheet = src_rng.sheet

        # Determine last used row from source range; fall back to the sheet's
        # used_range if the source range's used area can't be discovered.
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

        # Write header if applicable.
        if has_headers:
            sheet.range(f"{out_col}1").value = "Running Total"

        formulas = [
            [f"=SUM(${src_col}${first_data_row}:{src_col}{r})"]
            for r in range(first_data_row, last_row + 1)
        ]
        out_rng = sheet.range(f"{out_col}{first_data_row}:{out_col}{last_row}")
        out_rng.formula = formulas
        row_count = last_row - first_data_row + 1
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"runningTotal failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Created {row_count} running total formulas in {output}.",
        "outputs": {"outputRange": output},
    }


registry.register("runningTotal", handler, mutates=True)
