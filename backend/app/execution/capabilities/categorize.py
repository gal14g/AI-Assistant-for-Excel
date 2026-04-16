"""categorize — label cells based on ordered if/else rules.

Ported from `frontend/src/engine/capabilities/categorize.ts`. Rules are
evaluated top-to-bottom; first match wins, otherwise `defaultValue` is written.
Operators: contains, equals, startsWith, endsWith, greaterThan, lessThan, regex.
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


def _to_num(val: Any) -> float | None:
    if val is None or val == "":
        return None
    try:
        return float(val)
    except (TypeError, ValueError):
        return None


def _apply_rule(cell: Any, rule: dict[str, Any]) -> bool:
    operator = rule.get("operator")
    raw_val = rule.get("value")
    str_cell = ("" if cell is None else str(cell)).lower()
    str_val = ("" if raw_val is None else str(raw_val)).lower()

    if operator == "contains":
        return str_val in str_cell
    if operator == "equals":
        return str_cell == str_val
    if operator == "startsWith":
        return str_cell.startswith(str_val)
    if operator == "endsWith":
        return str_cell.endswith(str_val)
    if operator == "greaterThan":
        nc, nv = _to_num(cell), _to_num(raw_val)
        return nc is not None and nv is not None and nc > nv
    if operator == "lessThan":
        nc, nv = _to_num(cell), _to_num(raw_val)
        return nc is not None and nv is not None and nc < nv
    if operator == "regex":
        try:
            return re.search(str(raw_val), "" if cell is None else str(cell), re.IGNORECASE) is not None
        except re.error:
            return False
    return False


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    source_range = params.get("sourceRange")
    output_range = params.get("outputRange")
    rules = params.get("rules") or []
    default_value = params.get("defaultValue", "")

    if not source_range or not output_range:
        return {"status": "error", "message": "categorize requires 'sourceRange' and 'outputRange'."}
    if not rules:
        return {"status": "error", "message": "categorize requires at least one rule."}

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would categorize {source_range} with {len(rules)} rules.",
        }

    src = resolve_range(ctx.workbook_handle, source_range)
    vals = _to_2d(src.value)
    if not vals:
        return {"status": "success", "message": "No data to categorize.", "outputs": {"outputRange": output_range}}

    label_counts: dict[str, int] = {}
    out: list[list[Any]] = []
    total_cells = 0
    for row in vals:
        new_row: list[Any] = []
        for cell in row:
            total_cells += 1
            assigned: Any = default_value
            for rule in rules:
                if _apply_rule(cell, rule):
                    label = rule.get("label", "")
                    assigned = label
                    label_counts[label] = label_counts.get(label, 0) + 1
                    break
            new_row.append(assigned)
        out.append(new_row)

    rows = len(out)
    cols = max((len(r) for r in out), default=0)
    dest = resolve_range(ctx.workbook_handle, output_range)
    try:
        if hasattr(dest, "resize"):
            dest.resize(rows, cols).value = out
        else:
            dest.value = out
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Write failed: {exc}", "error": str(exc)}

    summary = ", ".join(f"{k}:{v}" for k, v in label_counts.items()) or "no matches"
    return {
        "status": "success",
        "message": f"Categorized {total_cells} cells — {summary}.",
        "outputs": {"outputRange": dest.address},
    }


registry.register("categorize", handler, mutates=True)
