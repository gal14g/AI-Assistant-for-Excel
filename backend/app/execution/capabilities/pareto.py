"""pareto — 80/20 analysis: sort by value desc + cumulative % + optional chart."""

from __future__ import annotations

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
    include_chart = bool(params.get("includeChart", True))
    has_headers = bool(params.get("hasHeaders", True))

    if not data_range or not output_range:
        return {"status": "error", "message": "pareto requires 'dataRange' and 'outputRange'."}

    src = resolve_range(ctx.workbook_handle, data_range)
    src_used = src.current_region if src.count > 1 else src

    if src_used.columns.count < 2:
        return {"status": "error", "message": "dataRange must have 2 columns (label | value)."}

    raw = _ensure_2d(src_used.value)
    start = 1 if has_headers else 0
    rows = []
    for r in raw[start:]:
        lbl = r[0]
        val = r[1] if len(r) > 1 else None
        # Tolerate text-stored numbers from CSV imports ($1,234, 50%, etc.)
        num = parse_number_flexible(val)
        if num is not None:
            rows.append({"label": lbl if lbl is not None else "", "value": num})
    if not rows:
        return {"status": "error", "message": "No numeric value rows found."}

    rows.sort(key=lambda r: r["value"], reverse=True)
    total = sum(r["value"] for r in rows)
    running = 0.0
    output: list[list[Any]] = [["Label", "Value", "Cumulative %"]]
    for r in rows:
        running += r["value"]
        output.append([r["label"], r["value"], 0 if total == 0 else running / total])

    if ctx.dry_run:
        return {"status": "preview", "message": f"Would build Pareto with {len(rows)} item(s)."}

    out = resolve_range(ctx.workbook_handle, output_range)
    out_top = out[0, 0]
    out_sheet = out.sheet
    out_addr_tail = out.address.split("!")[-1].split(":")[0]
    m = re.match(r"^\$?([A-Z]+)\$?(\d+)$", out_addr_tail)
    if not m:
        return {"status": "error", "message": f"Could not parse output address: {out.address}"}
    out_col_l = m.group(1)
    out_col_idx = _col_letters_to_index(out_col_l)
    out_row = int(m.group(2))

    try:
        block_target = out_sheet.range(
            (out_row, out_col_idx + 1),
            (out_row + len(output) - 1, out_col_idx + 3),
        )
        block_target.value = output

        # Format cumulative-% column as 0.0%.
        pct_col_l = _index_to_col_letters(out_col_idx + 2)
        out_sheet.range(f"{pct_col_l}{out_row + 1}:{pct_col_l}{out_row + len(output) - 1}").number_format = "0.0%"
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Write failed: {exc}", "error": str(exc)}

    chart_name: str | None = None
    if include_chart:
        try:
            chart_range = out_sheet.range(
                f"{out_col_l}{out_row}:{_index_to_col_letters(out_col_idx + 2)}{out_row + len(output) - 1}"
            )
            chart = out_sheet.charts.add()
            chart.chart_type = "column_clustered"
            chart.set_source_data(chart_range)
            try:
                chart.api[1].ChartTitle.Text = "Pareto Analysis"
            except Exception:
                pass
            chart_name = chart.name
        except Exception:
            pass

    output_addr = f"{out_sheet.name}!{out_col_l}{out_row}:{_index_to_col_letters(out_col_idx + 2)}{out_row + len(output) - 1}"
    return {
        "status": "success",
        "message": f"Pareto written with {len(rows)} item(s). Output: {output_addr}.",
        "outputs": {"outputRange": output_addr, **({"chartName": chart_name} if chart_name else {})},
    }


registry.register("pareto", handler, mutates=True)
