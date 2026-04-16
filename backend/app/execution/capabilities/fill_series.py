"""fillSeries — generate a number / date / weekday / pattern series into a range."""

from __future__ import annotations

from datetime import date, timedelta
from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range
from app.execution.utils import parse_date_flexible


def _add_unit(d: date, step: int, unit: str) -> date:
    if unit == "day":
        return d + timedelta(days=step)
    if unit == "week":
        return d + timedelta(weeks=step)
    if unit == "month":
        mo = d.month - 1 + step
        y = d.year + mo // 12
        mo = mo % 12 + 1
        # Clamp day to end-of-month to avoid Feb-31.
        try:
            return date(y, mo, d.day)
        except ValueError:
            return date(y, mo, 28)
    if unit == "year":
        try:
            return date(d.year + step, d.month, d.day)
        except ValueError:
            return date(d.year + step, d.month, 28)
    return d


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    series_type = params.get("seriesType")
    start = params.get("start")
    step = float(params.get("step", 1))
    pattern = params.get("pattern")
    date_unit = params.get("dateUnit", "day")
    count = params.get("count")
    horizontal = bool(params.get("horizontal", False))

    if not address or not series_type:
        return {"status": "error", "message": "fillSeries requires 'range' and 'seriesType'."}

    rng = resolve_range(ctx.workbook_handle, address)
    total_cells = count if count else (rng.columns.count if horizontal else rng.rows.count)
    total_cells = int(total_cells or 0)
    if total_cells <= 0:
        return {"status": "success", "message": "No cells to fill."}

    series: list[Any] = []
    if series_type == "number":
        s = float(start) if start is not None else 1.0
        for i in range(total_cells):
            series.append(s + i * step)
    elif series_type == "date":
        d = parse_date_flexible(start) or date.today()
        for i in range(total_cells):
            series.append(_add_unit(d, int(i * step), date_unit).strftime("%d/%m/%Y"))
    elif series_type == "weekday":
        d = parse_date_flexible(start) or date.today()
        while d.weekday() >= 5:  # Sat=5, Sun=6
            d += timedelta(days=1)
        for _ in range(total_cells):
            series.append(d.strftime("%d/%m/%Y"))
            d += timedelta(days=int(step))
            while d.weekday() >= 5:
                d += timedelta(days=1)
    elif series_type == "repeatPattern":
        if not pattern:
            return {"status": "error", "message": "seriesType='repeatPattern' requires 'pattern'."}
        for i in range(total_cells):
            series.append(pattern[i % len(pattern)])
    else:
        return {"status": "error", "message": f"Unknown seriesType: {series_type}"}

    if ctx.dry_run:
        return {"status": "preview", "message": f"Would fill {total_cells} {series_type} value(s) to {rng.address}."}

    grid = [series] if horizontal else [[v] for v in series]
    sheet = rng.sheet
    top_left = rng[0, 0]
    try:
        target = sheet.range(
            (top_left.row, top_left.column),
            (top_left.row + len(grid) - 1, top_left.column + len(grid[0]) - 1),
        )
        target.value = grid
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Write failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Wrote {total_cells} {series_type} value(s) to {rng.address}.",
        "outputs": {"range": rng.address, "filledCount": total_cells},
    }


registry.register("fillSeries", handler, mutates=True)
