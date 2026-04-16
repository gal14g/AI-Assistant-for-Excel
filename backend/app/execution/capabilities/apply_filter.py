"""applyFilter — set an AutoFilter on a range."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    filters = params.get("filters") or []
    if not address:
        return {"status": "error", "message": "applyFilter requires 'range'."}

    rng = resolve_range(ctx.workbook_handle, address)
    if ctx.dry_run:
        return {"status": "preview", "message": f"Would apply {len(filters)} filter(s) to {rng.address}."}

    # AutoFilter is set on the Range.api side because xlwings doesn't wrap it.
    sheet_api = rng.sheet.api
    # Clear any existing filter first.
    try:
        if getattr(sheet_api, "AutoFilterMode", False):
            sheet_api.AutoFilterMode = False
    except Exception:
        pass

    if not filters:
        # Toggle filter on only — no criteria.
        rng.api.AutoFilter()
        return {
            "status": "success",
            "message": f"Enabled filter on {rng.address}.",
            "outputs": {"range": rng.address},
        }

    for f in filters:
        col_idx = int(f.get("columnIndex", 0)) + 1  # Excel is 1-based
        criterion = f.get("criterion") or f.get("value")
        operator = f.get("operator", "equals")
        op_map = {"equals": "=", "notEquals": "<>", "greaterThan": ">", "lessThan": "<"}
        op_str = op_map.get(operator, "=")
        criterion_str = f"{op_str}{criterion}" if operator != "equals" else str(criterion)
        try:
            rng.api.AutoFilter(Field=col_idx, Criteria1=criterion_str)
        except Exception as exc:  # noqa: BLE001
            return {"status": "error", "message": f"AutoFilter failed on column {col_idx}: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Applied {len(filters)} filter(s) to {rng.address}.",
        "outputs": {"range": rng.address},
    }


registry.register("applyFilter", handler, mutates=True)
