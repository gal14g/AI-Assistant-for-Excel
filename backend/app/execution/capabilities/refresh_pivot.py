"""refreshPivot — refresh one or all PivotTables on a sheet."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    pivot_name = params.get("pivotName")
    sheet_name = params.get("sheetName")
    book = ctx.workbook_handle

    if ctx.dry_run:
        return {"status": "preview", "message": "Would refresh pivot(s)."}

    target = book.sheets[sheet_name] if sheet_name else book.sheets.active
    refreshed = 0
    try:
        pivots = target.api.PivotTables()
        # pivots.Count is 1-based; iterate explicitly because xlwings doesn't
        # wrap PivotTables iteration.
        for i in range(1, pivots.Count + 1):
            pt = pivots(i)
            if pivot_name and pt.Name != pivot_name:
                continue
            pt.RefreshTable()
            refreshed += 1
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Refresh failed: {exc}", "error": str(exc)}

    return {"status": "success", "message": f"Refreshed {refreshed} pivot(s) on {target.name}."}


registry.register("refreshPivot", handler, mutates=False)
