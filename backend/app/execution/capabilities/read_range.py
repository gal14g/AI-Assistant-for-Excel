"""readRange — read values from a range. Non-mutating."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    if not address:
        return {"status": "error", "message": "readRange requires a 'range' parameter."}

    rng = resolve_range(ctx.workbook_handle, address)
    values = rng.value

    # Normalize: single-cell reads come back as a scalar; consumers of the
    # result (binding resolver, chat client) expect 2D data. Match the
    # frontend's shape so the same plan produces identical outputs.
    if rng.count == 1:
        values_2d = [[values]]
        row_count, col_count = 1, 1
    else:
        if not isinstance(values, list):
            values_2d = [[values]]
        elif values and not isinstance(values[0], list):
            values_2d = [values]
        else:
            values_2d = values
        row_count, col_count = rng.shape

    return {
        "status": "success",
        "message": f"Read {row_count}×{col_count} from {rng.address}.",
        "outputs": {
            "range": rng.address,
            "values": values_2d,
            "rowCount": row_count,
            "columnCount": col_count,
        },
    }


registry.register("readRange", handler, mutates=False)
