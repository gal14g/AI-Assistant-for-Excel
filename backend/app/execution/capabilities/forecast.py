"""forecast — project a time series forward (FORECAST.LINEAR / FORECAST.ETS)."""

from __future__ import annotations

import re
from datetime import date, timedelta
from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range
from app.execution.utils import parse_date_flexible


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


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    source_range = params.get("sourceRange")
    output_range = params.get("outputRange")
    periods = int(params.get("periods") or 0)
    method = params.get("method", "linear")
    include_chart = bool(params.get("includeChart", True))
    has_headers = bool(params.get("hasHeaders", True))

    if not source_range or not output_range or periods < 1:
        return {"status": "error", "message": "forecast requires sourceRange, outputRange, periods>=1."}

    src = resolve_range(ctx.workbook_handle, source_range)
    src_used = src.current_region if src.count > 1 else src
    if src_used.columns.count < 2:
        return {"status": "error", "message": "sourceRange must have 2 columns (dates | values)."}

    raw = _ensure_2d(src_used.value)
    data_rows = raw[1:] if has_headers else raw[:]
    if len(data_rows) < 2:
        return {"status": "error", "message": "Need at least 2 source rows to infer date step."}

    d1 = parse_date_flexible(data_rows[-2][0])
    d2 = parse_date_flexible(data_rows[-1][0])
    if d1 is None or d2 is None:
        return {"status": "error", "message": "Could not parse the last two source dates."}
    step_days = (d2 - d1).days

    src_addr = src_used.address
    src_tail = src_addr.split("!")[-1]
    m = re.match(r"^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$", src_tail)
    if not m:
        return {"status": "error", "message": f"Could not parse source address: {src_addr}"}
    col_a, row1, col_b, row_n = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
    sheet_prefix = src_addr.split("!")[0] + "!" if "!" in src_addr else ""
    first_data_row = row1 + 1 if has_headers else row1
    known_y = f"{sheet_prefix}{col_b}{first_data_row}:{col_b}{row_n}"
    known_x = f"{sheet_prefix}{col_a}{first_data_row}:{col_a}{row_n}"

    fn_name = "FORECAST.ETS" if method == "ets" else "FORECAST.LINEAR"

    # Compose output.
    out = resolve_range(ctx.workbook_handle, output_range)
    out_top = out[0, 0]
    out_sheet = out.sheet
    out_addr_tail = out.address.split("!")[-1].split(":")[0]
    om = re.match(r"^\$?([A-Z]+)\$?(\d+)$", out_addr_tail)
    if not om:
        return {"status": "error", "message": f"Could not parse output address: {out.address}"}
    out_col_l = om.group(1)
    out_col_idx = _col_letters_to_index(out_col_l)
    out_row = int(om.group(2))

    header: list[Any] = ["Date", "Forecast"]
    body: list[list[Any]] = [header]
    for i in range(periods):
        future_date = d2 + timedelta(days=step_days * (i + 1))
        body.append([future_date.strftime("%d/%m/%Y"), None])

    if ctx.dry_run:
        return {"status": "preview", "message": f"Would project {periods} period(s) via {fn_name}."}

    try:
        block_target = out_sheet.range(
            (out_row, out_col_idx + 1),
            (out_row + periods, out_col_idx + 2),
        )
        block_target.value = body

        # Forecast formulas in the second column.
        fcol_l = _index_to_col_letters(out_col_idx + 1)
        formulas = []
        for i in range(periods):
            date_ref = f"{out_col_l}{out_row + 1 + i}"
            formulas.append([f"={fn_name}({date_ref},{known_y},{known_x})"])
        out_sheet.range(f"{fcol_l}{out_row + 1}:{fcol_l}{out_row + periods}").formula = formulas

        # Format date column.
        out_sheet.range(f"{out_col_l}{out_row + 1}:{out_col_l}{out_row + periods}").number_format = "dd/mm/yyyy"
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Write failed: {exc}", "error": str(exc)}

    chart_name: str | None = None
    if include_chart:
        try:
            chart_range = out_sheet.range(
                f"{out_col_l}{out_row}:{_index_to_col_letters(out_col_idx + 1)}{out_row + periods}"
            )
            chart = out_sheet.charts.add()
            chart.chart_type = "line"
            chart.set_source_data(chart_range)
            try:
                chart.api[1].ChartTitle.Text = f"Forecast ({'ETS' if method == 'ets' else 'Linear'})"
            except Exception:
                pass
            chart_name = chart.name
        except Exception:
            pass

    output_addr = f"{out_sheet.name}!{out_col_l}{out_row}:{_index_to_col_letters(out_col_idx + 1)}{out_row + periods}"
    return {
        "status": "success",
        "message": f"Projected {periods} period(s) via {fn_name}. Output: {output_addr}.",
        "outputs": {"outputRange": output_addr, **({"chartName": chart_name} if chart_name else {})},
    }


registry.register("forecast", handler, mutates=True)
