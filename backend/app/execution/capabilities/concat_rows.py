"""concatRows — write TEXTJOIN formulas per row into an output column."""

from __future__ import annotations

import re
from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    source_range = params.get("sourceRange")
    output_column = params.get("outputColumn")
    separator = params.get("separator", ", ")
    ignore_blanks = bool(params.get("ignoreBlanks", True))
    has_headers = bool(params.get("hasHeaders", True))

    if not source_range or not output_column:
        return {"status": "error", "message": "concatRows requires 'sourceRange' and 'outputColumn'."}

    src = resolve_range(ctx.workbook_handle, source_range)
    used = src.current_region if src.count > 1 else src

    if not used or used.rows.count == 0:
        return {"status": "success", "message": "Source range is empty."}

    addr = used.address
    m = re.match(r"^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$", addr.split("!")[-1] if "!" in addr else addr)
    if not m:
        return {"status": "error", "message": f"Could not parse source address: {addr}"}
    first_col_l, first_row, last_col_l, last_row = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))

    # Normalize output column.
    out_col = output_column
    out_sheet_name = None
    if "!" in out_col:
        parts = out_col.split("!", 1)
        out_sheet_name = parts[0].strip("'")
        out_col = parts[1]
    out_col = re.sub(r"[:$]", "", out_col)
    out_col = re.sub(r"\d+$", "", out_col).upper()
    if not re.match(r"^[A-Z]+$", out_col):
        return {"status": "error", "message": f"Invalid outputColumn: {output_column!r}"}

    src_sheet_name = used.sheet.name
    out_sheet = ctx.workbook_handle.sheets[out_sheet_name] if out_sheet_name else used.sheet
    sheet_prefix = f"'{src_sheet_name}'!" if " " in src_sheet_name else f"{src_sheet_name}!"

    data_start = first_row + 1 if has_headers else first_row
    data_count = max(0, last_row - data_start + 1)
    if data_count == 0:
        return {"status": "success", "message": "No data rows to concatenate."}

    # Build formulas.
    esc_sep = separator.replace('"', '""')
    ignore_tf = "TRUE" if ignore_blanks else "FALSE"
    formulas = [
        [f'=TEXTJOIN("{esc_sep}",{ignore_tf},{sheet_prefix}{first_col_l}{r}:{last_col_l}{r})']
        for r in range(data_start, data_start + data_count)
    ]

    if ctx.dry_run:
        return {"status": "preview", "message": f"Would write {data_count} TEXTJOIN formulas."}

    try:
        if has_headers:
            out_sheet.range(f"{out_col}{first_row}").value = "Joined"
        target = out_sheet.range(f"{out_col}{data_start}:{out_col}{data_start + data_count - 1}")
        target.formula = formulas
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Write failed: {exc}", "error": str(exc)}

    out_addr = f"{out_sheet.name}!{out_col}{data_start}:{out_col}{data_start + data_count - 1}"
    return {
        "status": "success",
        "message": f"Wrote {data_count} TEXTJOIN formulas to {out_addr}.",
        "outputs": {"outputRange": out_addr},
    }


registry.register("concatRows", handler, mutates=True)
