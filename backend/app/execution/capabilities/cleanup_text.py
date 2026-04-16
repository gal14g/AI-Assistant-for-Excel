"""cleanupText — apply text cleanup operations (trim, case, whitespace, etc.).

Ported from `frontend/src/engine/capabilities/cleanupText.ts`. Reads values,
runs string transforms, writes back. Non-string cells pass through untouched.
"""

from __future__ import annotations

import re
from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


# Control characters: U+0000-U+001F except tab (0x09), LF (0x0A), CR (0x0D), and DEL (0x7F).
# Matches the TS regex /[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g exactly — preserves Unicode.
_NON_PRINTABLE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]")
_WHITESPACE_RUN = re.compile(r"\s+")
# (^|\s)(\S) — capitalises first char of every whitespace-delimited word.
_WORD_START = re.compile(r"(^|\s)(\S)")


def _to_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if not isinstance(values[0], list):
        return [values]
    return values


def _apply_op(value: str, operation: str) -> str:
    if operation == "trim":
        return value.strip()
    if operation == "lowercase":
        return value.lower()
    if operation == "uppercase":
        return value.upper()
    if operation == "properCase":
        return _WORD_START.sub(lambda m: m.group(1) + m.group(2).upper(), value)
    if operation == "removeNonPrintable":
        return _NON_PRINTABLE.sub("", value)
    if operation == "normalizeWhitespace":
        return _WHITESPACE_RUN.sub(" ", value).strip()
    return value


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    operations = params.get("operations") or []
    output_range = params.get("outputRange")

    if not address:
        return {"status": "error", "message": "cleanupText requires a 'range' parameter."}
    if not operations:
        return {"status": "error", "message": "cleanupText requires at least one operation."}

    rng = resolve_range(ctx.workbook_handle, address)

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would apply [{', '.join(operations)}] to {rng.address}.",
        }

    vals = _to_2d(rng.value)
    if not vals:
        return {"status": "success", "message": "No data to clean.", "outputs": {"outputRange": output_range or rng.address}}

    cleaned: list[list[Any]] = []
    for row in vals:
        new_row: list[Any] = []
        for cell in row:
            if isinstance(cell, str):
                val = cell
                for op in operations:
                    val = _apply_op(val, op)
                new_row.append(val)
            else:
                new_row.append(cell)
        cleaned.append(new_row)

    target_addr = output_range if output_range else address
    target = resolve_range(ctx.workbook_handle, target_addr)
    rows = len(cleaned)
    cols = max((len(r) for r in cleaned), default=0)

    try:
        if hasattr(target, "resize"):
            target.resize(rows, cols).value = cleaned
        else:
            target.value = cleaned
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Write failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Applied [{', '.join(operations)}] to {rows} rows in {target.address}.",
        "outputs": {"outputRange": target.address},
    }


registry.register("cleanupText", handler, mutates=True)
