"""sortRange — sort a range by one or more columns."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    criteria = params.get("criteria") or []
    has_headers = bool(params.get("hasHeaders", True))
    if not address:
        return {"status": "error", "message": "sortRange requires 'range'."}
    if not criteria:
        return {"status": "error", "message": "sortRange requires at least one sort criterion."}

    rng = resolve_range(ctx.workbook_handle, address)
    if ctx.dry_run:
        return {"status": "preview", "message": f"Would sort {rng.address} by {len(criteria)} key(s)."}

    # xlwings exposes sort via .api for full parity with Excel's sort engine.
    # Build key1/order1, key2/order2, key3/order3 (Excel supports up to 3).
    # For more than 3 keys we do multi-pass sorts (stable: sort by least
    # important first so later sorts preserve prior ordering for ties).
    for crit in reversed(criteria[:8]):  # cap at 8 for safety
        key_col = crit.get("columnIndex")
        if key_col is None:
            continue
        ascending = (crit.get("order") or "asc").lower() in ("asc", "ascending")
        # xlwings' Range.sort() accepts a column key (1-based absolute Excel col).
        # Convert columnIndex (0-based within the range) to the absolute col.
        first_col = rng.column
        abs_col = first_col + int(key_col)
        key_range = rng.sheet.range((rng.row, abs_col))
        try:
            rng.api.Sort(
                Key1=key_range.api,
                Order1=1 if ascending else 2,
                Header=1 if has_headers else 2,
            )
        except Exception as exc:  # noqa: BLE001
            return {"status": "error", "message": f"Sort failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Sorted {rng.address} by {len(criteria)} key(s).",
        "outputs": {"range": rng.address},
    }


registry.register("sortRange", handler, mutates=True)
