"""addReportHeader — insert a formatted report title row above existing data, merged across the data width."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import parse_address, resolve_range


# Alignment constants (match Excel xlHAlign / xlVAlign enums).
_H_CENTER = -4108  # xlCenter
_V_CENTER = -4108  # xlCenter


def _hex_to_rgb_int(hex_color: str) -> int:
    """Convert "#RRGGBB" → Excel long int (BGR packed)."""
    h = hex_color.lstrip("#")
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    return r + g * 256 + b * 65536


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    title = params.get("title")
    sheet_name = params.get("sheetName")
    address = params.get("range")
    font_size = params.get("fontSize", 16)
    fill_color = params.get("fillColor", "#4472C4")
    font_color = params.get("fontColor", "#FFFFFF")
    bold = params.get("bold", True)

    if not title:
        return {"status": "error", "message": "addReportHeader requires 'title'."}

    book = ctx.workbook_handle

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would add report header {title!r}.",
        }

    try:
        # Resolve target sheet.
        if address:
            target = resolve_range(book, address)
            sheet = target.sheet
        elif sheet_name:
            sheet = book.sheets[sheet_name]
        else:
            sheet = book.sheets.active

        # Determine column span.
        if address:
            target_rng = resolve_range(book, address)
            # xlwings shape is (rows, cols) for multi-cell; single-cell has .count==1.
            if target_rng.count == 1:
                column_count = 1
            else:
                try:
                    column_count = target_rng.shape[1]
                except Exception:
                    column_count = target_rng.api.Columns.Count
        else:
            used = sheet.used_range
            try:
                column_count = used.shape[1] if used.count > 1 else 5
            except Exception:
                column_count = 5
            if column_count < 1:
                column_count = 5

        # Insert a new top row — shift existing rows down. xlShiftDown = -4121.
        sheet.api.Rows("1:1").Insert(Shift=-4121)

        # Write the title into A1 and merge across columns.
        sheet.range("A1").value = title
        header_rng = sheet.range((1, 1), (1, column_count))
        try:
            # Use MergeCells=False to merge into a single cell (Excel default behavior).
            header_rng.api.Merge()
        except Exception:
            pass

        # Apply formatting to the merged header.
        try:
            header_rng.api.Font.Size = font_size
            header_rng.api.Font.Bold = bool(bold)
            header_rng.api.Font.Color = _hex_to_rgb_int(font_color)
            header_rng.api.Interior.Color = _hex_to_rgb_int(fill_color)
            header_rng.api.HorizontalAlignment = _H_CENTER
            header_rng.api.VerticalAlignment = _V_CENTER
            header_rng.api.RowHeight = font_size * 2.5
        except Exception:
            # Formatting the merged range sometimes requires addressing the top cell.
            sheet.range("A1").api.RowHeight = font_size * 2.5

    except Exception as exc:  # noqa: BLE001
        return {
            "status": "error",
            "message": f"Add report header failed: {exc}",
            "error": str(exc),
        }

    return {
        "status": "success",
        "message": f"Added report header {title!r} to {sheet.name}.",
        "outputs": {"sheet": sheet.name, "range": address or "A1"},
    }


registry.register("addReportHeader", handler, mutates=True, affects_formatting=True)
