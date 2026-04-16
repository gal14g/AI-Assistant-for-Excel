"""aging — bucket dates into aging categories (0-30, 31-60, 61-90, 90+)."""

from __future__ import annotations

import re
from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    date_column = params.get("dateColumn")
    output_column = params.get("outputColumn")
    buckets = params.get("buckets") or [30, 60, 90]
    reference_date = params.get("referenceDate")
    has_headers = bool(params.get("hasHeaders", True))

    if not date_column or not output_column:
        return {"status": "error", "message": "aging requires 'dateColumn' and 'outputColumn'."}

    date_raw = resolve_range(ctx.workbook_handle, date_column)
    date_used = date_raw.current_region if date_raw.count > 1 else date_raw

    addr = date_used.address
    addr_tail = addr.split("!")[-1]
    m = re.match(r"^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$", addr_tail)
    if not m:
        # Single-cell / single-column fallback.
        m2 = re.match(r"^\$?([A-Z]+)\$?(\d+)$", addr_tail)
        if not m2:
            return {"status": "error", "message": f"Could not parse date column address: {addr}"}
        col_l = m2.group(1)
        top_row = int(m2.group(2))
        bot_row = top_row
    else:
        col_l = m.group(1)
        top_row = int(m.group(2))
        bot_row = int(m.group(4))

    first_data_row = top_row + 1 if has_headers else top_row
    last_data_row = bot_row
    data_count = max(0, last_data_row - first_data_row + 1)
    if data_count == 0:
        return {"status": "success", "message": "No data rows in dateColumn."}

    sheet_name = date_used.sheet.name
    sheet_prefix = f"'{sheet_name}'!" if " " in sheet_name else f"{sheet_name}!"

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

    out_sheet = ctx.workbook_handle.sheets[out_sheet_name] if out_sheet_name else date_used.sheet

    # Reference date expression.
    if reference_date:
        s = str(reference_date).strip()
        rm = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{2,4})$", s)
        if rm:
            y = int(rm.group(3))
            if y < 100:
                y += 2000
            ref_expr = f"DATE({y},{int(rm.group(2))},{int(rm.group(1))})"
        else:
            iso = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})$", s)
            if iso:
                ref_expr = f"DATE({int(iso.group(1))},{int(iso.group(2))},{int(iso.group(3))})"
            else:
                ref_expr = f'DATEVALUE("{s}")'
    else:
        ref_expr = "TODAY()"

    sorted_b = sorted(buckets)
    labels: list[str] = []
    for i, upper in enumerate(sorted_b):
        lower = 0 if i == 0 else sorted_b[i - 1] + 1
        labels.append(f"{lower}-{upper}")
    labels.append(f"{sorted_b[-1]}+")

    first_date_cell = f"{col_l}{first_data_row}"
    age_expr = f"({ref_expr}-{first_date_cell})"
    ifs_parts = [f'{age_expr}<={sorted_b[i]},"{labels[i]}"' for i in range(len(sorted_b))]
    ifs_parts.append(f'TRUE,"{labels[-1]}"')
    formula = f"=IFS({','.join(ifs_parts)})"

    if ctx.dry_run:
        return {"status": "preview", "message": f"Would write aging buckets for {data_count} row(s)."}

    try:
        if has_headers:
            out_sheet.range(f"{out_col}{top_row}").value = "Age Bucket"
        # Write formula to every data row. xlwings handles relative ref
        # adjustment when formulas are supplied as a list (same formula per
        # row, Excel adjusts the source cell row).
        target = out_sheet.range(f"{out_col}{first_data_row}:{out_col}{last_data_row}")
        target.formula = [[formula]] * data_count
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Write failed: {exc}", "error": str(exc)}

    output_addr = f"{sheet_prefix}{out_col}{first_data_row}:{out_col}{last_data_row}"
    return {
        "status": "success",
        "message": f"Wrote aging buckets ({', '.join(labels)}) to {output_addr}.",
        "outputs": {"outputRange": output_addr},
    }


registry.register("aging", handler, mutates=True)
