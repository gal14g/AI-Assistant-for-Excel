"""extractPattern — extract emails, phones, URLs, dates, numbers, or a custom
regex pattern from text cells.

Ported from `frontend/src/engine/capabilities/extractPattern.ts`. Built-in
pattern names map to the same regexes used by the frontend so results match.
"""

from __future__ import annotations

import re
from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


# Mirrors the JS BUILT_IN table in extractPattern.ts. All patterns are compiled
# without flags (matching the TS `"g"` flag — `re.findall` is global by default).
BUILT_IN: dict[str, str] = {
    "email":    r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}",
    "phone":    r"(?:\+?[\d\s\-().]{7,15})",
    "url":      r"https?://[^\s\"'>]+",
    "date":     r"\b\d{1,4}[/\-.]\d{1,2}[/\-.]\d{1,4}\b",
    "number":   r"[-+]?\d+(?:[.,]\d+)*",
    "currency": r"[$€£₪¥]?\s*\d+(?:[.,]\d{2,})+",
}


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


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    source_range = params.get("sourceRange")
    pattern = params.get("pattern")
    output_range = params.get("outputRange")
    all_matches = bool(params.get("allMatches", False))

    if not source_range or pattern is None or not output_range:
        return {
            "status": "error",
            "message": "extractPattern requires 'sourceRange', 'pattern', and 'outputRange'.",
        }

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would extract '{pattern}' from {source_range}.",
        }

    src = resolve_range(ctx.workbook_handle, source_range)

    pattern_src = BUILT_IN.get(pattern, pattern)
    try:
        regex = re.compile(pattern_src)
    except re.error:
        # Mirror the TS fallback to an always-empty regex rather than erroring.
        regex = re.compile(r"(?:)")

    vals = _to_2d(src.value)
    if not vals:
        return {"status": "success", "message": "No data to scan.", "outputs": {"outputRange": output_range}}

    results: list[list[Any]] = []
    found = 0
    for row in vals:
        new_row: list[Any] = []
        for cell in row:
            text = "" if cell is None else str(cell)
            matches = regex.findall(text)
            # re.findall returns tuples when there are groups; normalise to the
            # full match like JS .matchAll's m[0]. The simplest way is to use
            # finditer and take group(0).
            full_matches = [m.group(0) for m in regex.finditer(text)]
            if not full_matches:
                new_row.append(None)
            else:
                found += 1
                new_row.append(", ".join(full_matches) if all_matches else full_matches[0])
            _ = matches  # placate linter; we used finditer instead
        results.append(new_row)

    rows = len(results)
    cols = max((len(r) for r in results), default=0)
    out = resolve_range(ctx.workbook_handle, output_range)
    try:
        if hasattr(out, "resize"):
            out.resize(rows, cols).value = results
        else:
            out.value = results
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Write failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Extracted {found} match(es) using '{pattern}' from {rows} rows.",
        "outputs": {"outputRange": out.address},
    }


registry.register("extractPattern", handler, mutates=True)
