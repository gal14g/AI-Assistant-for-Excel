"""calculationMode — set workbook calculation mode (manual / automatic / automaticExceptTables)."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry


# Excel COM CalculationMode enum:
#   -4135 = xlCalculationAutomatic
#   -4105 = xlCalculationSemiautomatic (AutomaticExceptTables)
#   -4135 is wrong in some docs; the actual values are:
#   xlCalculationAutomatic = -4105, xlCalculationManual = -4135,
#   xlCalculationSemiautomatic = 2. Use xlwings constants where available.
_MODE_MAP = {
    "automatic": -4105,               # xlCalculationAutomatic
    "automaticexcepttables": 2,       # xlCalculationSemiautomatic
    "manual": -4135,                  # xlCalculationManual
}


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    mode = (params.get("mode") or "").strip().lower()
    if mode not in _MODE_MAP:
        return {
            "status": "error",
            "message": f"Unknown calc mode: {params.get('mode')!r}. Expected one of: manual, automatic, automaticExceptTables.",
        }

    if ctx.dry_run:
        return {"status": "preview", "message": f"Would set calculation mode to {mode}."}

    try:
        ctx.workbook_handle.app.api.Calculation = _MODE_MAP[mode]
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Setting calc mode failed: {exc}", "error": str(exc)}

    return {"status": "success", "message": f"Calculation mode set to {mode}."}


registry.register("calculationMode", handler, mutates=False)
