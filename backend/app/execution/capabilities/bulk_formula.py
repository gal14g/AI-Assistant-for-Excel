"""
bulkFormula — write a row-adjusted formula to a column up to the last data row.

Detects the last row of the data range (via xlwings' used-range resolution),
then replicates the template formula down the output column with row
references auto-incremented — exactly mirroring the TS implementation.

Template: "=A2*B2" at row 3 becomes "=A3*B3".
"""

from __future__ import annotations

import re
from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import parse_address, resolve_range


# Column-then-row references anywhere in the formula, e.g. A2, $B$17, AA100.
# We need to rewrite the row half only; leaving any `$` prefix intact.
_CELL_REF_RE = re.compile(r"([A-Z]+)(\d+)")


def _first_row_number(formula: str) -> int:
    """Return the row number of the first cell reference in the formula."""
    m = re.search(r"\d+", formula)
    return int(m.group(0)) if m else 1


def _strip_workbook_qualifier(address: str) -> str:
    """Remove `[Book.xlsx]` qualifier for pattern-based parsing."""
    return re.sub(r"\[[^\]]+\]", "", address)


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    formula = params.get("formula")
    output_range = params.get("outputRange")
    data_range = params.get("dataRange")
    has_headers = params.get("hasHeaders", True)

    if not formula or not output_range or not data_range:
        return {
            "status": "error",
            "message": "bulkFormula requires 'formula', 'outputRange', and 'dataRange'.",
        }

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would write formula {formula!r} to {output_range} based on data in {data_range}.",
        }

    try:
        data_rng = resolve_range(ctx.workbook_handle, data_range)
        sheet = data_rng.sheet

        # Determine the last row of actual data. Mirrors TS getUsedRange(false).
        try:
            used = data_rng.current_region if hasattr(data_rng, "current_region") else None
            if used is not None and used.count > 0:
                used_rows = used.shape[0]
                # Find the top row of data_rng (1-based).
                top_row = data_rng.row
                last_row = top_row + used_rows - 1
            else:
                raise AttributeError  # fall into fallback
        except Exception:  # noqa: BLE001 — fall back to whole-sheet used range
            whole = sheet.used_range
            last_row = whole.shape[0] if whole.shape else 1

        first_data_row = 2 if has_headers else 1
        if last_row < first_data_row:
            return {
                "status": "success",
                "message": "No data rows found.",
                "outputs": {},
            }

        row_count = last_row - first_data_row + 1

        # Determine output column letter.
        parsed_out = parse_address(output_range)
        out_cell = parsed_out.cell or "C1"
        out_cell_stripped = _strip_workbook_qualifier(out_cell)
        col_match = re.match(r"[A-Z]+", out_cell_stripped)
        out_col = col_match.group(0) if col_match else "C"

        # Build per-row formulas by offsetting every row-number digit.
        template_row = _first_row_number(formula)
        formulas: list[list[str]] = []
        for r in range(first_data_row, last_row + 1):
            offset = r - template_row
            if offset == 0:
                f = formula
            else:
                def _shift(m: re.Match[str]) -> str:
                    return f"{m.group(1)}{int(m.group(2)) + offset}"

                f = _CELL_REF_RE.sub(_shift, formula)
            formulas.append([f])

        out_addr = f"{out_col}{first_data_row}:{out_col}{last_row}"
        out_rng = sheet.range(out_addr)
        out_rng.formula = formulas
    except Exception as exc:  # noqa: BLE001
        return {
            "status": "error",
            "message": f"Bulk formula failed: {exc}",
            "error": str(exc),
        }

    return {
        "status": "success",
        "message": f"Wrote {formula!r} (adjusted) to {row_count} rows in column {out_col}.",
        "outputs": {"range": f"{out_col}{first_data_row}:{out_col}{last_row}"},
    }


registry.register("bulkFormula", handler, mutates=True)
