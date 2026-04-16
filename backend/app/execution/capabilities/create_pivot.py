"""
createPivot — create a PivotTable from a source range.

Cross-platform notes:
- On Windows xlwings gets the full VBA/COM object model — PivotCaches.Create
  plus PivotFields configuration works exactly like Excel's macro recorder.
- On macOS the same calls are translated via AppleScript; the API surface
  matches with only small enum differences. xlwings hides the split so this
  handler is uniform.
"""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


# Excel xlConsolidationFunction values.
_AGG = {
    "sum": -4157,
    "count": -4112,
    "average": -4106,
    "max": -4136,
    "min": -4139,
    "stdev": -4155,
    "product": -4149,
}


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    source = params.get("sourceRange")
    dest = params.get("destinationRange")
    pivot_name = params.get("pivotName") or f"Pivot_{abs(hash(source)) % 10_000_000}"
    rows_fields = params.get("rows") or []
    col_fields = params.get("columns") or []
    values = params.get("values") or []
    if not source:
        return {"status": "error", "message": "createPivot requires 'sourceRange'."}
    if not values:
        return {"status": "error", "message": "createPivot requires at least one value field."}

    book = ctx.workbook_handle
    src = resolve_range(book, source)

    if ctx.dry_run:
        return {"status": "preview", "message": f"Would create pivot {pivot_name!r} on {src.address}."}

    try:
        # Destination sheet — either a specified range or a new sheet named after the pivot.
        if dest:
            dest_rng = resolve_range(book, dest)
            dest_sheet = dest_rng.sheet
            dest_cell = dest_rng
        else:
            existing = next((s for s in book.sheets if s.name == pivot_name[:31]), None)
            dest_sheet = existing or book.sheets.add(pivot_name[:31])
            dest_cell = dest_sheet.range("A1")

        # PivotCaches().Create(SourceType=xlDatabase=1, SourceData=range).CreatePivotTable(TableDestination, TableName)
        cache = book.api.PivotCaches().Create(SourceType=1, SourceData=src.api)
        pivot = cache.CreatePivotTable(TableDestination=dest_cell.api, TableName=pivot_name)

        # Row fields — xlRowField = 1
        for field_name in rows_fields:
            try:
                pivot.PivotFields(field_name).Orientation = 1
            except Exception:
                pass

        # Column fields — xlColumnField = 2
        for field_name in col_fields:
            try:
                pivot.PivotFields(field_name).Orientation = 2
            except Exception:
                pass

        # Value fields — xlDataField = 4 with aggregation function.
        for v in values:
            field = v.get("field")
            agg = _AGG.get((v.get("summarizeBy") or "sum").lower(), _AGG["sum"])
            display = v.get("displayName") or f"{v.get('summarizeBy', 'Sum')} of {field}"
            try:
                pivot.AddDataField(pivot.PivotFields(field), display, agg)
            except Exception:
                # Fallback path — direct orientation assignment.
                pivot.PivotFields(field).Orientation = 4
                pivot.PivotFields(field).Function = agg
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Create pivot failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Created pivot {pivot_name!r} on {dest_sheet.name}.",
        "outputs": {"pivotName": pivot_name, "destinationSheet": dest_sheet.name},
    }


registry.register("createPivot", handler, mutates=True, affects_formatting=True)
