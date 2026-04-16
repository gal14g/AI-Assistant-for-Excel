"""unpivot — convert a wide table into a long (tidy) format.

The first ``idColumns`` columns are kept as-is on every output row. Remaining
columns are unpivoted into two new columns (variable, value), so one input row
becomes (cols - idColumns) output rows.
"""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    source = params.get("sourceRange")
    output = params.get("outputRange")
    id_columns = params.get("idColumns")
    variable_col_name = params.get("variableColumnName") or "Attribute"
    value_col_name = params.get("valueColumnName") or "Value"

    if not source or not output:
        return {
            "status": "error",
            "message": "unpivot requires 'sourceRange' and 'outputRange'.",
        }
    if id_columns is None:
        return {"status": "error", "message": "unpivot requires 'idColumns'."}

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": (
                f"Would unpivot {source} ({id_columns} id col(s)) to {output}."
            ),
        }

    try:
        src_rng = resolve_range(ctx.workbook_handle, source)
        raw = src_rng.value
        if raw is None:
            return {"status": "success", "message": "Not enough rows to unpivot.", "outputs": {}}
        if not isinstance(raw, list):
            data: list[list[Any]] = [[raw]]
        elif raw and not isinstance(raw[0], list):
            rows_, _ = src_rng.shape
            data = [list(raw)] if rows_ == 1 else [[v] for v in raw]
        else:
            data = [list(r) for r in raw]

        if len(data) < 2:
            return {"status": "success", "message": "Not enough rows to unpivot.", "outputs": {}}

        id_count = int(id_columns)
        headers = data[0]
        id_headers = [str(h) if h is not None else "" for h in headers[:id_count]]
        value_headers = headers[id_count:]

        out_headers: list[Any] = [*id_headers, variable_col_name, value_col_name]
        out_rows: list[list[Any]] = [out_headers]

        for r in range(1, len(data)):
            row = data[r]
            id_vals = list(row[:id_count])
            # Pad id_vals if row is short.
            if len(id_vals) < id_count:
                id_vals.extend([None] * (id_count - len(id_vals)))
            for c in range(len(value_headers)):
                source_col = id_count + c
                value = row[source_col] if source_col < len(row) else None
                out_rows.append([*id_vals, str(value_headers[c]) if value_headers[c] is not None else "", value])

        rows = len(out_rows)
        cols = len(out_headers)
        out_rng = resolve_range(ctx.workbook_handle, output)
        out_rng.resize(rows, cols).value = out_rows
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"unpivot failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": (
            f"Unpivoted {len(data) - 1} rows into {rows - 1} rows "
            f"({len(value_headers)} value columns → rows)."
        ),
        "outputs": {"outputRange": output},
    }


registry.register("unpivot", handler, mutates=True)
