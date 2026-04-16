"""deleteRowsByCondition — delete rows whose column meets a predicate.

Port of `frontend/src/engine/capabilities/deleteRowsByCondition.ts`. Iterates
the used range of the supplied address, collects the 0-based row offsets that
match the predicate, then deletes them bottom-up via the Excel COM
`Rows(n).Delete(Shift:=xlUp)` API so indices don't shift under us.
"""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def _as_2d(raw: Any, shape: tuple[int, int]) -> list[list[Any]]:
    if raw is None:
        return []
    rows_, cols_ = shape
    if not isinstance(raw, list):
        return [[raw]]
    if raw and not isinstance(raw[0], list):
        if rows_ == 1:
            return [list(raw)]
        return [[v] for v in raw]
    return [list(r) for r in raw]


def _meets_condition(cell_val: Any, condition: str, value: Any) -> bool:
    str_val = str(cell_val if cell_val is not None else "").strip()
    is_empty = cell_val is None or cell_val == "" or str_val == ""

    if condition == "blank":
        return is_empty
    if condition == "notBlank":
        return not is_empty
    if condition == "equals":
        if value is None:
            return False
        return str_val.lower() == str(value).lower()
    if condition == "notEquals":
        if value is None:
            return False
        return str_val.lower() != str(value).lower()
    if condition == "contains":
        if value is None:
            return False
        return str(value).lower() in str_val.lower()
    if condition == "greaterThan":
        if value is None:
            return False
        try:
            return float(cell_val) > float(value)
        except (TypeError, ValueError):
            return str_val > str(value)
    if condition == "lessThan":
        if value is None:
            return False
        try:
            return float(cell_val) < float(value)
        except (TypeError, ValueError):
            return str_val < str(value)
    return False


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    column = params.get("column")
    condition = params.get("condition")
    value = params.get("value")
    has_headers = bool(params.get("hasHeaders", True))

    if not address or column is None or not condition:
        return {
            "status": "error",
            "message": "deleteRowsByCondition requires 'range', 'column' and 'condition'.",
        }

    if ctx.dry_run:
        val_str = f' "{value}"' if value is not None else ""
        return {
            "status": "preview",
            "message": f"Would delete rows in {address} where column {column} is {condition}{val_str}.",
        }

    try:
        rng = resolve_range(ctx.workbook_handle, address)
        vals = _as_2d(rng.value, rng.shape)
        if not vals:
            return {"status": "success", "message": "No data found.", "outputs": {}}

        col_idx = int(column) - 1
        start_data_row = 1 if has_headers else 0

        # Absolute sheet-row numbers of matching rows (1-based).
        sheet_start_row = rng.row
        matching_sheet_rows: list[int] = []
        for i in range(start_data_row, len(vals)):
            cell_val = vals[i][col_idx] if 0 <= col_idx < len(vals[i]) else None
            if _meets_condition(cell_val, condition, value):
                matching_sheet_rows.append(sheet_start_row + i)

        if not matching_sheet_rows:
            return {
                "status": "success",
                "message": f'No rows matched condition "{condition}" in column {column}.',
                "outputs": {"range": address, "deletedCount": 0},
            }

        # Bottom-up deletion via COM so earlier indices remain valid.
        sheet = rng.sheet
        for row_num in sorted(matching_sheet_rows, reverse=True):
            sheet.api.Rows(row_num).Delete()

        cond_desc = f'{condition} "{value}"' if value is not None else condition
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"deleteRowsByCondition failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Deleted {len(matching_sheet_rows)} rows where column {column} {cond_desc}.",
        "outputs": {"range": address, "deletedCount": len(matching_sheet_rows)},
    }


registry.register("deleteRowsByCondition", handler, mutates=True)
