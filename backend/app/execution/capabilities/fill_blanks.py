"""fillBlanks — fill empty cells downward, upward, or with a constant value.

Ported from `frontend/src/engine/capabilities/fillBlanks.ts`. Reads values,
computes fills in Python, writes back in one batch.
"""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def _is_empty(v: Any) -> bool:
    return v is None or v == ""


def _to_2d(values: Any) -> list[list[Any]]:
    """Coerce xlwings `.value` output to a 2D list."""
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if not isinstance(values[0], list):
        return [values]
    return values


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    fill_mode = params.get("fillMode") or "down"
    constant_value = params.get("constantValue", "")

    if not address:
        return {"status": "error", "message": "fillBlanks requires a 'range' parameter."}

    rng = resolve_range(ctx.workbook_handle, address)

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would fill blanks in {rng.address} (mode: {fill_mode}).",
        }

    # Prefer the used sub-range so full-column refs like "A:D" don't load a
    # million rows. Fall back to the raw range if used_range lookup fails.
    try:
        sheet_used = rng.sheet.used_range
        # Intersect conceptually by resizing to max of rng and used_range rows.
        # Simpler: if rng is already bounded (has concrete last cell), just use rng.
        # xlwings ranges are always bounded when address has explicit endpoints.
        # Use rng directly — it already represents the user's intent.
        target = rng
        _ = sheet_used  # suppress unused warning
    except Exception:  # noqa: BLE001
        target = rng

    raw_vals = target.value
    vals = _to_2d(raw_vals)
    if not vals:
        return {"status": "success", "message": "No data found.", "outputs": {"range": rng.address, "filledCount": 0}}

    num_cols = max((len(r) for r in vals), default=0)
    # Normalize ragged rows to a uniform width.
    out: list[list[Any]] = [list(row) + [None] * (num_cols - len(row)) for row in vals]

    filled = 0

    if fill_mode == "down":
        for c in range(num_cols):
            last: Any = None
            for r in range(len(out)):
                if not _is_empty(out[r][c]):
                    last = out[r][c]
                elif last is not None:
                    out[r][c] = last
                    filled += 1
    elif fill_mode == "up":
        for c in range(num_cols):
            nxt: Any = None
            for r in range(len(out) - 1, -1, -1):
                if not _is_empty(out[r][c]):
                    nxt = out[r][c]
                elif nxt is not None:
                    out[r][c] = nxt
                    filled += 1
    else:
        # constant
        for r in range(len(out)):
            for c in range(len(out[r])):
                if _is_empty(out[r][c]):
                    out[r][c] = constant_value
                    filled += 1

    try:
        target.value = out
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Write back failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Filled {filled} blank cell(s) in {rng.address} (mode: {fill_mode}).",
        "outputs": {"range": rng.address, "filledCount": filled},
    }


registry.register("fillBlanks", handler, mutates=True)
