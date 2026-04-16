"""writeValues — bulk write a 2D array of values to a range."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    values = params.get("values")
    if not address or values is None:
        return {"status": "error", "message": "writeValues requires 'range' and 'values'."}

    # Accept flat lists — reshape into 2D to match the frontend semantics.
    if isinstance(values, list) and (not values or not isinstance(values[0], list)):
        values = [values] if values else [[]]

    rng = resolve_range(ctx.workbook_handle, address)

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would write {len(values)}×{len(values[0]) if values else 0} to {rng.address}.",
        }

    # Resize to match the incoming matrix so partial addresses like "A1"
    # auto-expand. Mirrors the frontend's `range.getResizedRange` behavior.
    rows = len(values)
    cols = max((len(r) for r in values), default=0)
    target = rng.resize(rows, cols) if hasattr(rng, "resize") else rng
    target.value = values

    return {
        "status": "success",
        "message": f"Wrote {rows}×{cols} to {target.address}.",
        "outputs": {"range": target.address, "rowCount": rows, "columnCount": cols},
    }


registry.register("writeValues", handler, mutates=True)
