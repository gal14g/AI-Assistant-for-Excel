"""insertDeleteColumns — insert or delete worksheet columns.

Parity with insert_delete_rows.py. Accepts a column-letter range ("C:E") or
a cell-address range and uses its column span via COM's `EntireColumn`.
"""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    action = params.get("action")
    shift_direction = params.get("shiftDirection", "right")

    if not address:
        return {"status": "error", "message": "insertDeleteColumns requires 'range'."}
    if action not in ("insert", "delete"):
        return {"status": "error", "message": "'action' must be 'insert' or 'delete'."}

    rng = resolve_range(ctx.workbook_handle, address)

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would {action} columns {rng.address}.",
        }

    try:
        col_range = rng.api.EntireColumn
        if action == "insert":
            # xlShiftToRight = -4161, xlShiftDown = -4121 (irrelevant here).
            col_range.Insert()
        else:
            col_range.Delete()
    except Exception as exc:  # noqa: BLE001
        return {
            "status": "error",
            "message": f"{action.capitalize()} columns failed on {rng.address}: {exc}",
            "error": str(exc),
        }

    return {
        "status": "success",
        "message": f"{action.capitalize()}ed columns {rng.address} (shift {shift_direction}).",
        "outputs": {"range": rng.address, "columnCount": rng.columns.count},
    }


registry.register("insertDeleteColumns", handler, mutates=True)
