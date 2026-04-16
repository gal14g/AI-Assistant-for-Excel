"""autoFitColumns — auto-fit column widths on a range."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    if not address:
        return {"status": "error", "message": "autoFitColumns requires 'range'."}

    rng = resolve_range(ctx.workbook_handle, address)
    if ctx.dry_run:
        return {"status": "preview", "message": f"Would autofit columns on {rng.address}."}

    try:
        rng.columns.autofit()
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"AutoFit failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Autofit columns on {rng.address}.",
        "outputs": {"range": rng.address},
    }


registry.register("autoFitColumns", handler, mutates=False, affects_formatting=True)
