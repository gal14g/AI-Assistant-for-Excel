"""createChart — embed a chart on a sheet from a source range."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


# Map friendly names to xlChartType enum values.
_CHART_TYPE = {
    "column": 51,          # xlColumnClustered
    "columnClustered": 51,
    "bar": 57,             # xlBarClustered
    "line": 4,             # xlLine
    "pie": 5,              # xlPie
    "scatter": -4169,      # xlXYScatter
    "area": 1,             # xlArea
    "doughnut": -4120,     # xlDoughnut
}


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    source = params.get("sourceRange")
    chart_type = (params.get("chartType") or "column").lower()
    title = params.get("title")
    dest_sheet_name = params.get("sheetName")
    position = params.get("position") or {}
    if not source:
        return {"status": "error", "message": "createChart requires 'sourceRange'."}

    book = ctx.workbook_handle
    src = resolve_range(book, source)
    if ctx.dry_run:
        return {"status": "preview", "message": f"Would create {chart_type} chart from {src.address}."}

    try:
        sheet = book.sheets[dest_sheet_name] if dest_sheet_name else src.sheet
        chart = sheet.charts.add(
            left=position.get("left", 100),
            top=position.get("top", 100),
            width=position.get("width", 400),
            height=position.get("height", 300),
        )
        chart.set_source_data(src)
        xl_type_val = _CHART_TYPE.get(chart_type, 51)
        chart.api[1].ChartType = xl_type_val
        if title:
            chart.api[1].HasTitle = True
            chart.api[1].ChartTitle.Text = title
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Create chart failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Created {chart_type} chart on {sheet.name}.",
        "outputs": {"sheet": sheet.name, "chartType": chart_type},
    }


registry.register("createChart", handler, mutates=True, affects_formatting=True)
