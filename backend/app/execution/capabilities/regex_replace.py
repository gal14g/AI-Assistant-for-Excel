"""regexReplace — apply regex find-and-replace across string cells in a range.

Ported from `frontend/src/engine/capabilities/regexReplace.ts`. JS flags like
"gi" are translated to Python `re` flags; the `g` flag is implicit in
`re.sub` (it replaces all non-overlapping occurrences by default).

Capture group references ($1, $2) used by JS regex are translated to Python
backrefs (\\1, \\2) so the replacement string behaves identically.
"""

from __future__ import annotations

import re
from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


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


def _translate_flags(js_flags: str) -> int:
    """Convert JS regex flags (e.g. 'gi') to Python re flags."""
    flags = 0
    if not js_flags:
        return flags
    if "i" in js_flags:
        flags |= re.IGNORECASE
    if "m" in js_flags:
        flags |= re.MULTILINE
    if "s" in js_flags:
        flags |= re.DOTALL
    if "u" in js_flags:
        flags |= re.UNICODE
    # 'g' is implicit in re.sub; 'y' (sticky) has no direct equivalent.
    return flags


# Replace JS-style $1 / $2 backrefs with Python \1 / \2 so re.sub handles them.
_JS_BACKREF = re.compile(r"\$(\d+)")


def _translate_replacement(repl: str) -> str:
    # Escape Python's own backslash escapes first, then map $N → \N.
    # Users may also legitimately use \\ in their replacement; leave as-is.
    return _JS_BACKREF.sub(r"\\\1", repl)


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    pattern = params.get("pattern")
    replacement = params.get("replacement", "")
    flags_str = params.get("flags", "gi") or ""

    if not address or pattern is None:
        return {"status": "error", "message": "regexReplace requires 'range' and 'pattern'."}

    rng = resolve_range(ctx.workbook_handle, address)

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would apply regex /{pattern}/{flags_str} replacement on {rng.address}.",
        }

    try:
        regex = re.compile(pattern, _translate_flags(flags_str))
    except re.error as exc:
        return {
            "status": "error",
            "message": f"Invalid regex /{pattern}/{flags_str}: {exc}",
            "error": str(exc),
        }

    py_repl = _translate_replacement(replacement)

    vals = _to_2d(rng.value)
    if not vals:
        return {"status": "success", "message": "No data to process.", "outputs": {"range": rng.address, "replacementCount": 0}}

    replaced_cells = 0
    new_vals: list[list[Any]] = []
    for row in vals:
        new_row: list[Any] = []
        for cell in row:
            if not isinstance(cell, str):
                new_row.append(cell)
                continue
            if regex.search(cell) is None:
                new_row.append(cell)
                continue
            try:
                new_row.append(regex.sub(py_repl, cell))
                replaced_cells += 1
            except re.error as exc:
                return {
                    "status": "error",
                    "message": f"Replacement failed: {exc}",
                    "error": str(exc),
                }
        new_vals.append(new_row)

    try:
        rng.value = new_vals
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Write back failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Replaced {replaced_cells} cell(s) across {len(vals)} rows in {rng.address}.",
        "outputs": {"range": rng.address, "replacementCount": replaced_cells},
    }


registry.register("regexReplace", handler, mutates=True)
