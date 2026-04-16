"""crossTabulate — build a cross-tab (contingency) matrix from raw data.

Takes a 1-based row field, column field and value field from the source
range and produces a matrix: unique row keys × unique column keys, with
per-cell aggregation (count / sum / average). Adds a trailing Total column
and Total row. Output is a static 2D block written to the output range.
"""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def _to_float(x: Any) -> float:
    try:
        if x is None or x == "":
            return 0.0
        return float(x)
    except (TypeError, ValueError):
        return 0.0


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    source = params.get("sourceRange")
    row_field = params.get("rowField")
    column_field = params.get("columnField")
    value_field = params.get("valueField")
    aggregation = (params.get("aggregation") or "count").lower()
    output = params.get("outputRange")

    if not source or not output:
        return {
            "status": "error",
            "message": "crossTabulate requires 'sourceRange' and 'outputRange'.",
        }
    if row_field is None or column_field is None or value_field is None:
        return {
            "status": "error",
            "message": "crossTabulate requires 'rowField', 'columnField', and 'valueField'.",
        }

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": (
                f"Would cross-tabulate {source} (row: col {row_field}, "
                f"col: col {column_field})."
            ),
        }

    try:
        src_rng = resolve_range(ctx.workbook_handle, source)
        raw = src_rng.value
        if raw is None:
            return {"status": "success", "message": "Not enough data.", "outputs": {}}
        if not isinstance(raw, list):
            data: list[list[Any]] = [[raw]]
        elif raw and not isinstance(raw[0], list):
            rows_, _ = src_rng.shape
            data = [list(raw)] if rows_ == 1 else [[v] for v in raw]
        else:
            data = [list(r) for r in raw]

        if len(data) < 2:
            return {"status": "success", "message": "Not enough data.", "outputs": {}}

        row_idx = int(row_field) - 1
        col_idx = int(column_field) - 1
        val_idx = int(value_field) - 1

        # Collect unique row/column keys (preserving first-seen order).
        row_keys: list[str] = []
        col_keys: list[str] = []
        row_set: set[str] = set()
        col_set: set[str] = set()
        for r in range(1, len(data)):
            row = data[r]
            rk = str(row[row_idx]) if row_idx < len(row) and row[row_idx] is not None else ""
            ck = str(row[col_idx]) if col_idx < len(row) and row[col_idx] is not None else ""
            if rk not in row_set:
                row_set.add(rk)
                row_keys.append(rk)
            if ck not in col_set:
                col_set.add(ck)
                col_keys.append(ck)

        # Initialize matrix[rk][ck] = {sum, count}.
        matrix: dict[str, dict[str, dict[str, float]]] = {
            rk: {ck: {"sum": 0.0, "count": 0} for ck in col_keys} for rk in row_keys
        }

        for r in range(1, len(data)):
            row = data[r]
            rk = str(row[row_idx]) if row_idx < len(row) and row[row_idx] is not None else ""
            ck = str(row[col_idx]) if col_idx < len(row) and row[col_idx] is not None else ""
            raw_val = row[val_idx] if val_idx < len(row) else None
            v = _to_float(raw_val)
            if v == 0.0 and aggregation == "count":
                v = 1.0  # count: each row contributes 1
            cell = matrix.get(rk, {}).get(ck)
            if cell is not None:
                cell["sum"] += v
                cell["count"] += 1

        # Build output.
        out_rows: list[list[Any]] = []
        out_rows.append(["", *col_keys, "Total"])

        for rk in row_keys:
            row_out: list[Any] = [rk]
            row_total = 0.0
            for ck in col_keys:
                cell = matrix[rk][ck]
                if aggregation == "count":
                    v = cell["count"]
                elif aggregation == "sum":
                    v = cell["sum"]
                else:  # average
                    v = (cell["sum"] / cell["count"]) if cell["count"] else 0.0
                row_out.append(v)
                row_total += cell["count"] if aggregation == "count" else cell["sum"]
            row_out.append(row_total)
            out_rows.append(row_out)

        totals_row: list[Any] = ["Total"]
        grand_total = 0.0
        for ck in col_keys:
            col_total = 0.0
            for rk in row_keys:
                cell = matrix[rk][ck]
                col_total += cell["count"] if aggregation == "count" else cell["sum"]
            totals_row.append(col_total)
            grand_total += col_total
        totals_row.append(grand_total)
        out_rows.append(totals_row)

        rows = len(out_rows)
        cols = len(out_rows[0])
        out_rng = resolve_range(ctx.workbook_handle, output)
        out_rng.resize(rows, cols).value = out_rows
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"crossTabulate failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": (
            f"Cross-tab: {len(row_keys)} rows × {len(col_keys)} columns "
            f"({aggregation}) written to {output}."
        ),
        "outputs": {"outputRange": output},
    }


registry.register("crossTabulate", handler, mutates=True)
