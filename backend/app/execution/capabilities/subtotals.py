"""subtotals — insert aggregate rows at each group boundary.

Reads the data range, sorts rows by the group-by column, and builds a new
matrix with a subtotal row appended after every run of matching group keys.
Inserts blank rows below the original range when the output has grown,
then writes the new matrix back in place.
"""

from __future__ import annotations

import re
from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


_COL_RE = re.compile(r"([A-Z]+)(\d+)", re.IGNORECASE)


def _to_float(x: Any) -> float:
    try:
        if x is None or x == "":
            return 0.0
        return float(x)
    except (TypeError, ValueError):
        return 0.0


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    data_range = params.get("dataRange")
    group_by = params.get("groupByColumn")
    subtotal_columns = params.get("subtotalColumns") or []
    aggregation = (params.get("aggregation") or "sum").lower()
    subtotal_label = params.get("subtotalLabel") or "Total"

    if not data_range:
        return {"status": "error", "message": "subtotals requires 'dataRange'."}
    if group_by is None:
        return {"status": "error", "message": "subtotals requires 'groupByColumn'."}
    if not subtotal_columns:
        return {"status": "error", "message": "subtotals requires 'subtotalColumns'."}

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": (
                f"Would add {aggregation} subtotals to {data_range} "
                f"grouped by column {group_by}."
            ),
        }

    try:
        rng = resolve_range(ctx.workbook_handle, data_range)
        raw = rng.value
        if raw is None:
            return {"status": "success", "message": "Not enough rows.", "outputs": {}}
        if not isinstance(raw, list):
            vals: list[list[Any]] = [[raw]]
        elif raw and not isinstance(raw[0], list):
            rows_, _ = rng.shape
            vals = [list(raw)] if rows_ == 1 else [[v] for v in raw]
        else:
            vals = [list(r) for r in raw]

        if len(vals) < 2:
            return {"status": "success", "message": "Not enough rows.", "outputs": {}}

        header_row = list(vals[0])
        data_rows = [list(r) for r in vals[1:]]
        grp_idx = int(group_by) - 1
        sub_cols = [int(c) - 1 for c in subtotal_columns]

        # Sort by group column (stringify for stable ordering).
        data_rows.sort(key=lambda r: str(r[grp_idx] if grp_idx < len(r) else ""))

        out: list[list[Any]] = [header_row]
        current_group = (
            str(data_rows[0][grp_idx])
            if data_rows and grp_idx < len(data_rows[0]) and data_rows[0][grp_idx] is not None
            else ""
        )
        group_rows: list[list[Any]] = []

        def flush(group_key: str) -> None:
            nonlocal group_rows
            out.extend(group_rows)
            sub_row: list[Any] = [None] * len(header_row)
            if grp_idx < len(sub_row):
                sub_row[grp_idx] = f"{group_key} {subtotal_label}"
            for ci in sub_cols:
                nums = [
                    _to_float(r[ci]) for r in group_rows if ci < len(r)
                ]
                if aggregation == "sum":
                    sub_row[ci] = sum(nums)
                elif aggregation == "count":
                    sub_row[ci] = len(nums)
                else:  # average
                    sub_row[ci] = (sum(nums) / len(nums)) if nums else 0
            out.append(sub_row)
            group_rows = []

        for row in data_rows:
            grp = str(row[grp_idx]) if grp_idx < len(row) and row[grp_idx] is not None else ""
            if grp != current_group:
                flush(current_group)
                current_group = grp
            group_rows.append(row)
        flush(current_group)

        # Insert blank rows if new matrix is taller than the original.
        sheet = rng.sheet
        start_row = rng.row
        start_col_idx = rng.column

        # Figure out start column letter for building addresses.
        def _col_letter(col_num: int) -> str:
            letters = ""
            n = col_num
            while n > 0:
                n, rem = divmod(n - 1, 26)
                letters = chr(65 + rem) + letters
            return letters

        start_col = _col_letter(start_col_idx)
        original_rows = len(vals)
        extra_rows = len(out) - original_rows
        if extra_rows > 0:
            try:
                insert_start = start_row + original_rows
                insert_end = insert_start + extra_rows - 1
                sheet.range(f"{insert_start}:{insert_end}").api.Insert(Shift=-4121)  # xlDown
            except Exception as insert_err:  # noqa: BLE001
                return {
                    "status": "error",
                    "message": (
                        f"Failed to insert subtotal rows: {insert_err}. "
                        "Range may contain merged cells."
                    ),
                    "error": str(insert_err),
                }

        rows = len(out)
        cols = max((len(r) for r in out), default=0)
        try:
            sheet.range(f"{start_col}{start_row}").resize(rows, cols).value = out
        except Exception as write_err:  # noqa: BLE001
            return {
                "status": "error",
                "message": (
                    f"Failed to write subtotals: {write_err}. "
                    "Range may contain merged or protected cells."
                ),
                "error": str(write_err),
            }

        subtotal_count = sum(
            1
            for r in out
            if grp_idx < len(r)
            and isinstance(r[grp_idx], str)
            and r[grp_idx].endswith(subtotal_label)
        )
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"subtotals failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Added {subtotal_count} subtotal rows to {data_range}.",
        "outputs": {"range": data_range},
    }


registry.register("subtotals", handler, mutates=True)
