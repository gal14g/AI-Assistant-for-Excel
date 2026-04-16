"""sheetPosition — move a sheet to a specific tab-order position (0-based)."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    position = params.get("position")
    sheet_name = params.get("sheetName")
    if position is None:
        return {"status": "error", "message": "sheetPosition requires 'position'."}

    book = ctx.workbook_handle
    sheet = book.sheets[sheet_name] if sheet_name else book.sheets.active

    if ctx.dry_run:
        return {"status": "preview", "message": f"Would move {sheet.name!r} to position {position}."}

    try:
        # xlwings doesn't expose position directly; use COM Move.
        # Excel's Move(Before, After) — position 0 means before the first sheet.
        sheets_api = book.sheets[0].api.Parent.Sheets  # Sheets collection
        count = sheets_api.Count
        target = max(0, min(int(position), count - 1))
        if target == 0:
            sheet.api.Move(Before=sheets_api.Item(1))
        else:
            # Move AFTER the sheet currently at `target` in 1-based COM terms,
            # accounting for the sheet we're removing from its current spot.
            # Simplest: if moving later, move after target's sheet; if earlier,
            # move before target's sheet.
            current = book.sheets.index(sheet)
            if target > current:
                sheet.api.Move(After=sheets_api.Item(target + 1))
            else:
                sheet.api.Move(Before=sheets_api.Item(target + 1))
    except Exception as exc:  # noqa: BLE001
        return {
            "status": "error",
            "message": f"Move failed: {exc}",
            "error": str(exc),
        }

    return {
        "status": "success",
        "message": f"Moved sheet {sheet.name!r} to position {position}.",
        "outputs": {"sheetName": sheet.name},
    }


registry.register("sheetPosition", handler, mutates=False)
