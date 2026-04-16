"""splitByGroup — split a data range into separate sheets by unique values
in one column.

Ported from `frontend/src/engine/capabilities/splitByGroup.ts`. For each
distinct value in the group-by column we create a worksheet (name sanitised
to Excel's 31-char / reserved-char limits) and write the matching rows.
"""

from __future__ import annotations

import re
from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


_INVALID_SHEET_CHARS = re.compile(r"[:\\/?*\[\]]")


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


def _sanitize_sheet_name(name: str) -> str:
    cleaned = _INVALID_SHEET_CHARS.sub("_", name)
    return cleaned[:31] or "Group"


def _unique_sheet_name(book: Any, desired: str) -> str:
    """Ensure we don't collide with existing sheets — append _2, _3, ..."""
    existing = {s.name for s in book.sheets}
    if desired not in existing:
        return desired
    base = desired[: 31 - 2]  # leave room for "_N"
    i = 2
    while True:
        candidate = f"{base}_{i}"[:31]
        if candidate not in existing:
            return candidate
        i += 1


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    data_range = params.get("dataRange")
    group_by_column = params.get("groupByColumn")
    keep_headers = bool(params.get("keepHeaders", True))

    if not data_range or group_by_column is None:
        return {
            "status": "error",
            "message": "splitByGroup requires 'dataRange' and 'groupByColumn'.",
        }

    try:
        col_1based = int(group_by_column)
    except (TypeError, ValueError):
        return {"status": "error", "message": "'groupByColumn' must be an integer (1-based)."}

    if col_1based < 1:
        return {"status": "error", "message": "'groupByColumn' must be >= 1."}

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would split {data_range} into sheets by column {col_1based}.",
        }

    rng = resolve_range(ctx.workbook_handle, data_range)
    vals = _to_2d(rng.value)
    if not vals:
        return {"status": "success", "message": "No data to split.", "outputs": {"sheetCount": 0}}

    col_idx = col_1based - 1  # 0-based
    header_row = vals[0]
    data_rows = vals[1:]

    groups: dict[str, list[list[Any]]] = {}
    for row in data_rows:
        key_val = row[col_idx] if col_idx < len(row) else None
        key = ("" if key_val is None else str(key_val)).strip()
        if not key:
            continue
        groups.setdefault(key, []).append(row)

    book = ctx.workbook_handle
    created: list[str] = []

    try:
        for group_name, rows in groups.items():
            clean_name = _sanitize_sheet_name(group_name)
            unique_name = _unique_sheet_name(book, clean_name)
            sheet = book.sheets.add(name=unique_name)

            write_rows: list[list[Any]] = []
            if keep_headers:
                write_rows.append(list(header_row))
            write_rows.extend(rows)

            if write_rows and write_rows[0]:
                rows_count = len(write_rows)
                cols_count = max((len(r) for r in write_rows), default=0)
                target = sheet.range("A1")
                if hasattr(target, "resize"):
                    target.resize(rows_count, cols_count).value = write_rows
                else:
                    target.value = write_rows
            created.append(unique_name)
    except Exception as exc:  # noqa: BLE001
        return {
            "status": "error",
            "message": f"splitByGroup failed after creating {len(created)} sheet(s): {exc}",
            "error": str(exc),
            "outputs": {"sheetCount": len(created), "sheets": created},
        }

    return {
        "status": "success",
        "message": f"Split into {len(created)} sheet(s) by column {col_1based}.",
        "outputs": {"sheetCount": len(created), "sheets": created},
    }


registry.register("splitByGroup", handler, mutates=True)
