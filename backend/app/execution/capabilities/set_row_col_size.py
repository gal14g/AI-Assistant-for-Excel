"""setRowColSize — set row height (points) or column width (character widths) on a range."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    dimension = params.get("dimension")
    size = params.get("size")
    if not address or dimension not in ("rowHeight", "columnWidth") or size is None:
        return {
            "status": "error",
            "message": "setRowColSize requires 'range', 'dimension' (rowHeight|columnWidth), and 'size'.",
        }

    rng = resolve_range(ctx.workbook_handle, address)

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would set {dimension} to {size} on {rng.address}.",
        }

    try:
        if dimension == "rowHeight":
            rng.row_height = float(size)
        else:
            rng.column_width = float(size)
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Set size failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Set {dimension} to {size} on {rng.address}.",
        "outputs": {"range": rng.address, "dimension": dimension, "size": size},
    }


registry.register("setRowColSize", handler, mutates=False, affects_formatting=True)
