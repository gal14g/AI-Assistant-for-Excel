"""transpose — flip rows and columns of a range.

Reads the source range, transposes the 2D array, and writes the result to the
output range. Values-only by default (formatting is not copied).
"""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    source = params.get("sourceRange")
    output = params.get("outputRange")
    if not source or not output:
        return {
            "status": "error",
            "message": "transpose requires 'sourceRange' and 'outputRange'.",
        }

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would transpose {source} → {output}.",
        }

    try:
        src = resolve_range(ctx.workbook_handle, source)
        raw = src.value

        # Normalize to 2D matrix — xlwings returns scalar for 1 cell, list for
        # a single row/column, 2D list otherwise. Match the TS behavior.
        if raw is None:
            vals: list[list[Any]] = []
        elif not isinstance(raw, list):
            vals = [[raw]]
        elif raw and not isinstance(raw[0], list):
            # Single row or single column — disambiguate by rng shape.
            rows, cols = src.shape
            if rows == 1:
                vals = [list(raw)]
            else:
                vals = [[v] for v in raw]
        else:
            vals = [list(r) for r in raw]

        if not vals:
            return {"status": "success", "message": "No data to transpose.", "outputs": {}}

        rows = len(vals)
        cols = max((len(r) for r in vals), default=0)

        # Transpose: out[c][r] = vals[r][c]
        transposed: list[list[Any]] = [
            [vals[r][c] if c < len(vals[r]) else None for r in range(rows)]
            for c in range(cols)
        ]

        out_rows = len(transposed)
        out_cols = len(transposed[0]) if transposed else 0
        out_rng = resolve_range(ctx.workbook_handle, output)
        out_rng.resize(out_rows, out_cols).value = transposed
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Transpose failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Transposed {rows}×{cols} → {out_rows}×{out_cols} written to {output}.",
        "outputs": {"outputRange": output},
    }


registry.register("transpose", handler, mutates=True)
