"""tabColor — set or clear the color of a worksheet's tab.

xlwings: `sheet.api.Tab.Color = <BGR-int>` (Windows COM) — convert hex to BGR.
On Mac, AppleScript equivalent. Fall through gracefully if the COM property
isn't available.
"""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry


def _hex_to_bgr_int(hex_color: str) -> int:
    """Convert #RRGGBB (Excel-style) to the BGR integer COM expects."""
    h = hex_color.lstrip("#")
    if len(h) != 6:
        raise ValueError(f"Invalid hex color: {hex_color!r}")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return (b << 16) | (g << 8) | r


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    color = params.get("color", "")
    sheet_name = params.get("sheetName")
    book = ctx.workbook_handle
    sheet = book.sheets[sheet_name] if sheet_name else book.sheets.active

    if ctx.dry_run:
        return {"status": "preview", "message": f"Would set tab color of {sheet.name!r} to {color}."}

    try:
        if (color or "").strip().lower() in ("", "none"):
            # Clear tab color.
            sheet.api.Tab.ColorIndex = -4142  # xlColorIndexNone
        else:
            sheet.api.Tab.Color = _hex_to_bgr_int(color)
    except Exception as exc:  # noqa: BLE001
        return {
            "status": "error",
            "message": f"Setting tab color on {sheet.name!r} failed: {exc}",
            "error": str(exc),
        }

    return {
        "status": "success",
        "message": f"Set tab color on {sheet.name!r} to {color or '(cleared)'}.",
        "outputs": {"sheetName": sheet.name},
    }


registry.register("tabColor", handler, mutates=False, affects_formatting=True)
