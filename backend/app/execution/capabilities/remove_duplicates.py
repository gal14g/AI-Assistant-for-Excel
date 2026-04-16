"""removeDuplicates — delete duplicate rows from a range in place.

Ported from `frontend/src/engine/capabilities/removeDuplicates.ts`. The Office.js
side uses `Range.removeDuplicates()` (ExcelApi 1.9); xlwings goes through the COM
API's `Range.RemoveDuplicates` method which takes a 1-based `Columns` parameter
(VBA array of column indexes within the range) and a `Header` flag.
"""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    column_indexes = params.get("columnIndexes")

    if not address:
        return {"status": "error", "message": "removeDuplicates requires a 'range' parameter."}

    rng = resolve_range(ctx.workbook_handle, address)

    if ctx.dry_run:
        return {"status": "preview", "message": f"Would remove duplicates from {rng.address}."}

    # Frontend passes 0-based indexes; Excel's COM RemoveDuplicates wants 1-based
    # indexes within the range. Translate defensively — also accept 1-based
    # indexes coming straight from the plan schema when the planner already
    # normalised them.
    row_count_before = rng.rows.count

    try:
        if column_indexes:
            # Heuristic: if any index is 0, assume 0-based (TS style) and shift.
            if any(i == 0 for i in column_indexes):
                one_based = [int(i) + 1 for i in column_indexes]
            else:
                one_based = [int(i) for i in column_indexes]
            # COM accepts a tuple/list as a VBA Array when there are multiple
            # columns, or a single int when only one column is specified.
            cols: Any = one_based[0] if len(one_based) == 1 else tuple(one_based)
            rng.api.RemoveDuplicates(Columns=cols, Header=1)
        else:
            # No columns given → compare every column in the range.
            total_cols = rng.columns.count
            if total_cols <= 0:
                return {"status": "error", "message": "Range has zero columns."}
            cols_all: Any = 1 if total_cols == 1 else tuple(range(1, total_cols + 1))
            rng.api.RemoveDuplicates(Columns=cols_all, Header=1)
    except Exception as exc:  # noqa: BLE001
        return {
            "status": "error",
            "message": f"removeDuplicates failed: {exc}",
            "error": str(exc),
        }

    # Excel deletes duplicate rows in place. We can no longer introspect the
    # "removed" count directly from COM, so compare the used-range row count of
    # the anchor sheet before and after. A cheaper proxy: inspect the range's
    # current row count — RemoveDuplicates replaces duplicate rows with blanks
    # at the bottom of the range rather than physically removing them, so row
    # count stays the same. To get a meaningful count we walk the first column.
    try:
        after_vals = rng.columns[0].value
        if after_vals is None:
            unique_remaining = 0
        elif isinstance(after_vals, list):
            unique_remaining = sum(1 for v in after_vals if v not in (None, ""))
        else:
            unique_remaining = 1 if after_vals not in (None, "") else 0
    except Exception:  # noqa: BLE001
        unique_remaining = row_count_before

    removed = max(0, row_count_before - unique_remaining)

    return {
        "status": "success",
        "message": f"Removed {removed} duplicate row(s); {unique_remaining} unique rows remain in {rng.address}.",
        "outputs": {"range": rng.address, "removedCount": removed},
    }


registry.register("removeDuplicates", handler, mutates=True)
