"""freezePanes — freeze the top rows / left columns of a sheet."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    rows = int(params.get("rowCount", 0))
    cols = int(params.get("columnCount", 0))
    sheet_name = params.get("sheetName")

    book = ctx.workbook_handle
    sheet = book.sheets[sheet_name] if sheet_name else book.sheets.active

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would freeze {rows} row(s) and {cols} column(s) on {sheet.name}.",
        }

    # Freeze panes requires activating the sheet and then setting SplitRow/SplitColumn
    # on the active window. xlwings gives us the underlying Window via book.app.api.
    try:
        sheet.activate()
        app_api = book.app.api
        window = app_api.ActiveWindow
        # Clear any existing freeze first.
        window.FreezePanes = False
        window.SplitRow = rows
        window.SplitColumn = cols
        window.FreezePanes = True
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Freeze panes failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Froze {rows} row(s) / {cols} column(s) on {sheet.name}.",
        "outputs": {"sheet": sheet.name, "rowCount": rows, "columnCount": cols},
    }


registry.register("freezePanes", handler, mutates=False, affects_formatting=True)
