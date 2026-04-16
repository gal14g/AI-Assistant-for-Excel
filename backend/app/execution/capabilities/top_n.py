"""topN — extract the top N (or bottom N) rows sorted by a value column.

Reads the full data range, sorts rows by the specified 1-based value column
descending (for 'top') or ascending (for 'bottom'), takes the first N rows,
and writes them — with optional header row preserved — to the output range.
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
    data_range = params.get("dataRange")
    value_column = params.get("valueColumn")
    n = params.get("n")
    position = (params.get("position") or "top").lower()
    output_range = params.get("outputRange")
    has_headers = bool(params.get("hasHeaders", True))

    if not data_range or not output_range:
        return {
            "status": "error",
            "message": "topN requires 'dataRange' and 'outputRange'.",
        }
    if value_column is None or n is None:
        return {
            "status": "error",
            "message": "topN requires 'valueColumn' and 'n'.",
        }

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": (
                f"Would extract {position} {n} rows from {data_range} "
                f"by column {value_column} to {output_range}."
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
            rows_, cols_ = rng.shape
            vals = [list(raw)] if rows_ == 1 else [[v] for v in raw]
        else:
            vals = [list(r) for r in raw]

        if len(vals) < 2:
            return {"status": "success", "message": "Not enough rows.", "outputs": {}}

        header_row = vals[0] if has_headers else None
        data_rows = list(vals[1:]) if has_headers else list(vals)
        col_idx = int(value_column) - 1

        # Sort data rows — 'top' = descending, 'bottom' = ascending.
        data_rows.sort(
            key=lambda row: _to_float(row[col_idx] if col_idx < len(row) else 0),
            reverse=(position == "top"),
        )

        taken = data_rows[: int(n)]

        output: list[list[Any]] = []
        if header_row is not None:
            output.append(header_row)
        output.extend(taken)

        if not output:
            return {"status": "success", "message": "No rows to write.", "outputs": {}}

        rows = len(output)
        cols = max((len(r) for r in output), default=0)
        out_rng = resolve_range(ctx.workbook_handle, output_range)
        out_rng.resize(rows, cols).value = output
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"topN failed: {exc}", "error": str(exc)}

    label = "Top" if position == "top" else "Bottom"
    return {
        "status": "success",
        "message": f"{label} {len(taken)} rows by column {value_column} written to {output_range}.",
        "outputs": {"outputRange": output_range},
    }


registry.register("topN", handler, mutates=True)
