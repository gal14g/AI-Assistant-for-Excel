"""reorderRows — rearrange rows in place: moveMatching / reverse / clusterByKey."""

from __future__ import annotations

from typing import Any, Optional

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range
from app.execution.utils import normalize_for_compare, parse_number_flexible


def _ensure_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if not isinstance(values[0], list):
        return [values] if len(values) > 1 else [[values[0]]]
    return values


def _test_condition(cell: Any, condition: str, value: Any) -> bool:
    # Normalize strings for equality/contains so Hebrew + RTL marks + NBSP
    # all compare consistently. Use parse_number_flexible for numeric
    # comparisons so text-stored numbers work.
    if condition == "blank":
        return normalize_for_compare(cell) == ""
    if condition == "notBlank":
        return normalize_for_compare(cell) != ""
    if condition == "equals":
        if isinstance(cell, (int, float)) and isinstance(value, (int, float)):
            return cell == value
        return normalize_for_compare(cell) == normalize_for_compare(value)
    if condition == "notEquals":
        if isinstance(cell, (int, float)) and isinstance(value, (int, float)):
            return cell != value
        return normalize_for_compare(cell) != normalize_for_compare(value)
    if condition == "contains":
        return normalize_for_compare(value) in normalize_for_compare(cell)
    if condition == "notContains":
        return normalize_for_compare(value) not in normalize_for_compare(cell)
    if condition == "greaterThan":
        n = parse_number_flexible(cell)
        vv = parse_number_flexible(value)
        return n is not None and vv is not None and n > vv
    if condition == "lessThan":
        n = parse_number_flexible(cell)
        vv = parse_number_flexible(value)
        return n is not None and vv is not None and n < vv
    return False


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    mode = params.get("mode")
    condition_col = params.get("conditionColumn")
    condition = params.get("condition")
    condition_value = params.get("conditionValue")
    destination = params.get("destination", "top")
    has_headers = bool(params.get("hasHeaders", True))

    if not address or not mode:
        return {"status": "error", "message": "reorderRows requires 'range' and 'mode'."}

    if mode in ("moveMatching", "clusterByKey") and condition_col is None:
        return {"status": "error", "message": f"mode={mode} requires 'conditionColumn'."}

    rng = resolve_range(ctx.workbook_handle, address)
    used = rng.current_region if rng.count > 1 else rng
    raw = _ensure_2d(used.value)
    if not raw:
        return {"status": "success", "message": "Range is empty."}

    header_row: Optional[list[Any]] = raw[0] if has_headers else None
    data_rows = [list(r) for r in (raw[1:] if has_headers else raw)]

    moved_count = 0
    if mode == "reverse":
        new_order = list(reversed(data_rows))
    elif mode == "moveMatching":
        matching: list[list[Any]] = []
        rest: list[list[Any]] = []
        for row in data_rows:
            if _test_condition(row[condition_col] if len(row) > condition_col else None, condition or "", condition_value):
                matching.append(row)
            else:
                rest.append(row)
        moved_count = len(matching)
        new_order = (matching + rest) if destination == "top" else (rest + matching)
    elif mode == "clusterByKey":
        buckets: dict[str, list[list[Any]]] = {}
        key_order: list[str] = []
        for row in data_rows:
            raw_key = row[condition_col] if len(row) > condition_col else None
            # Normalize keys so Hebrew with RTL marks, NBSP padding, or NFC
            # variants cluster together.
            key = normalize_for_compare(raw_key)
            if key not in buckets:
                buckets[key] = []
                key_order.append(key)
            buckets[key].append(row)
        new_order = [r for k in key_order for r in buckets[k]]
        moved_count = sum(len(b) for b in buckets.values() if len(b) > 1)
    else:
        return {"status": "error", "message": f"Unknown mode: {mode}"}

    output = ([header_row] if header_row else []) + new_order

    if ctx.dry_run:
        return {"status": "preview", "message": f"Would reorderRows mode={mode} on {used.address}."}

    sheet = used.sheet
    top_left = used[0, 0]
    try:
        target = sheet.range(
            (top_left.row, top_left.column),
            (top_left.row + len(output) - 1, top_left.column + len(output[0]) - 1),
        )
        target.value = output
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Write failed: {exc}", "error": str(exc)}

    if mode == "reverse":
        msg = f"Reversed {len(data_rows)} rows in {used.address}."
    elif mode == "moveMatching":
        msg = f"Moved {moved_count} matching row(s) to {destination} of {used.address}."
    else:
        msg = f"Clustered rows in {used.address} by column {condition_col} ({moved_count} row(s) in multi-row groups)."

    return {
        "status": "success",
        "message": msg,
        "outputs": {"range": used.address, "movedRowCount": moved_count},
    }


registry.register("reorderRows", handler, mutates=True)
