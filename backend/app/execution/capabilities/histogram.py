"""histogram — FREQUENCY + bin table + optional column chart."""

from __future__ import annotations

import math
import re
from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range
from app.execution.utils import parse_number_flexible


def _col_letters_to_index(s: str) -> int:
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


def _index_to_col_letters(i: int) -> str:
    n = i + 1
    out = ""
    while n > 0:
        rem = (n - 1) % 26
        out = chr(ord("A") + rem) + out
        n = (n - 1) // 26
    return out


def _ensure_2d(v: Any) -> list[list[Any]]:
    if v is None:
        return []
    if not isinstance(v, list):
        return [[v]]
    if not v:
        return []
    if not isinstance(v[0], list):
        return [v] if len(v) > 1 else [[v[0]]]
    return v


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    data_range = params.get("dataRange")
    output_range = params.get("outputRange")
    bins_param = params.get("bins")
    bin_count = params.get("binCount")
    include_chart = bool(params.get("includeChart", True))
    chart_type = params.get("chartType", "columnClustered")
    has_headers = bool(params.get("hasHeaders", True))

    if not data_range or not output_range:
        return {"status": "error", "message": "histogram requires 'dataRange' and 'outputRange'."}

    data_raw = resolve_range(ctx.workbook_handle, data_range)
    data_used = data_raw.current_region if data_raw.count > 1 else data_raw

    raw = _ensure_2d(data_used.value)
    start_row = 1 if has_headers else 0
    nums: list[float] = []
    for r in raw[start_row:]:
        # Tolerate text-stored numbers from CSV imports ($1,234, 50%, etc.)
        n = parse_number_flexible(r[0])
        if n is not None:
            nums.append(n)
    if not nums:
        return {"status": "error", "message": "No numeric values in dataRange."}

    # Bins.
    if bins_param:
        bin_edges = sorted(bins_param)
    else:
        n = len(nums)
        count = bin_count or max(2, math.ceil(math.log2(n) + 1))
        mn, mx = min(nums), max(nums)
        step = (mx - mn) / count if count else 1
        bin_edges = [mn + step * (i + 1) for i in range(count)]

    # Output block header + bin labels + placeholder counts.
    rows = [["Bin", "Count"]]
    for edge in bin_edges:
        rows.append([edge, None])
    rows.append([f">{bin_edges[-1]}", None])

    if ctx.dry_run:
        return {"status": "preview", "message": f"Would build histogram with {len(bin_edges)} bins."}

    out = resolve_range(ctx.workbook_handle, output_range)
    out_top = out[0, 0]
    out_sheet = out.sheet
    out_addr_tail = out.address.split("!")[-1].split(":")[0]
    m = re.match(r"^\$?([A-Z]+)\$?(\d+)$", out_addr_tail)
    if not m:
        return {"status": "error", "message": f"Could not parse output address: {out.address}"}
    out_col_l = m.group(1)
    out_row = int(m.group(2))

    try:
        # Write the block (header + labels + null counts).
        block_target = out_sheet.range(
            (out_row, _col_letters_to_index(out_col_l) + 1),
            (out_row + len(rows) - 1, _col_letters_to_index(out_col_l) + 2),
        )
        block_target.value = rows

        # FREQUENCY array formula into the Count column.
        count_col_l = _index_to_col_letters(_col_letters_to_index(out_col_l) + 1)
        count_first = out_row + 1
        count_last = out_row + len(bin_edges) + 1
        bin_first = out_row + 1
        bin_last = out_row + len(bin_edges)

        # Source data range for FREQUENCY (skip header).
        data_addr = data_used.address
        data_addr_tail = data_addr.split("!")[-1]
        dm = re.match(r"^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$", data_addr_tail)
        sheet_prefix = data_addr.split("!")[0] + "!" if "!" in data_addr else ""
        if dm:
            c1, r1, c2, r2 = dm.group(1), int(dm.group(2)), dm.group(3), int(dm.group(4))
            data_start = r1 + 1 if has_headers else r1
            data_formula_ref = f"{sheet_prefix}{c1}{data_start}:{c2}{r2}"
        else:
            data_formula_ref = data_addr

        bin_ref = f"{out_col_l}{bin_first}:{out_col_l}{bin_last}"
        formula = f"=FREQUENCY({data_formula_ref},{bin_ref})"

        count_target = out_sheet.range(f"{count_col_l}{count_first}:{count_col_l}{count_last}")
        # Array-formula entry: xlwings `formula_array` for array formulas.
        try:
            count_target.api.FormulaArray = formula
        except Exception:
            count_target.formula = [[formula]] * (len(bin_edges) + 1)
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Write failed: {exc}", "error": str(exc)}

    chart_name: str | None = None
    if include_chart:
        try:
            chart_data = out_sheet.range(f"{out_col_l}{out_row}:{count_col_l}{count_last}")
            chart = out_sheet.charts.add()
            chart.chart_type = "column_clustered" if chart_type == "columnClustered" else "bar_clustered"
            chart.set_source_data(chart_data)
            try:
                chart.api[1].ChartTitle.Text = "Histogram"  # Windows COM path
            except Exception:
                pass
            chart_name = chart.name
        except Exception:
            pass

    output_addr = f"{out_sheet.name}!{out_col_l}{out_row}:{count_col_l}{count_last}"
    return {
        "status": "success",
        "message": f"Histogram built with {len(bin_edges)} bins + overflow. Output: {output_addr}.",
        "outputs": {"outputRange": output_addr, **({"chartName": chart_name} if chart_name else {})},
    }


registry.register("histogram", handler, mutates=True)
