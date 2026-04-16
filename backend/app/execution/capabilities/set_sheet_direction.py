"""setSheetDirection — request right-to-left or left-to-right sheet display.

In MCP mode this works: xlwings can set `sheet.api.DisplayRightToLeft = True`
via COM on Windows (and via AppleScript on Mac). This is the primary-path
implementation that the add-in handler returns a warning for.
"""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    direction = (params.get("direction") or "").strip().lower()
    sheet_name = params.get("sheetName")

    if direction not in ("rtl", "ltr"):
        return {
            "status": "error",
            "message": f"Unknown direction: {params.get('direction')!r}. Expected 'rtl' or 'ltr'.",
        }

    book = ctx.workbook_handle
    sheet = book.sheets[sheet_name] if sheet_name else book.sheets.active

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would set {sheet.name!r} display to {direction.upper()}.",
        }

    try:
        # COM property on Windows Excel. On Mac AppleScript the same property
        # path works through xlwings' AppleScript shim.
        sheet.api.DisplayRightToLeft = direction == "rtl"
    except Exception as exc:  # noqa: BLE001
        return {
            "status": "error",
            "message": (
                f"Could not set sheet direction on {sheet.name!r}: {exc}. "
                "On Excel versions / platforms without the DisplayRightToLeft "
                "property, toggle manually via View > Sheet Right-to-Left."
            ),
            "error": str(exc),
        }

    return {
        "status": "success",
        "message": f"Set {sheet.name!r} direction to {direction.upper()}.",
    }


registry.register("setSheetDirection", handler, mutates=False, affects_formatting=True)
