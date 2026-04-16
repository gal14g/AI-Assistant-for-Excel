"""
insertTextBox — insert a text box via the COM Shapes.AddTextbox API.

Styling mirrors the TS handler: font (size/family/color), fill color, and
horizontal alignment. Orientation=1 = horizontal text flow.
"""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry


# msoTextOrientationHorizontal = 1
_ORIENT_HORIZONTAL = 1

# XlHAlign-style alignment values for TextFrame2.TextRange.ParagraphFormat.Alignment.
# msoAlignLeft=1, msoAlignCenter=2, msoAlignRight=3
_ALIGN_MAP = {
    "left": 1,
    "center": 2,
    "centre": 2,
    "right": 3,
}


def _hex_to_rgb_int(color: str) -> int:
    h = color.lstrip("#")
    if len(h) != 6:
        raise ValueError(f"Invalid color {color!r}; expected #RRGGBB.")
    return int(h[0:2], 16) + int(h[2:4], 16) * 256 + int(h[4:6], 16) * 65536


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    text = params.get("text")
    left = params.get("left")
    top = params.get("top")
    width = params.get("width")
    height = params.get("height")
    sheet_name = params.get("sheetName")
    font_size = params.get("fontSize")
    font_family = params.get("fontFamily")
    font_color = params.get("fontColor")
    fill_color = params.get("fillColor")
    h_align = params.get("horizontalAlignment")

    if text is None or left is None or top is None or width is None or height is None:
        return {
            "status": "error",
            "message": "insertTextBox requires 'text', 'left', 'top', 'width', 'height'.",
        }

    preview = text if len(text) <= 30 else text[:30] + "..."

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would insert text box {preview!r} at ({left}, {top}).",
        }

    book = ctx.workbook_handle
    try:
        if sheet_name:
            try:
                sheet = book.sheets[sheet_name]
            except Exception:
                return {
                    "status": "error",
                    "message": f"Sheet {sheet_name!r} not found. Please check the sheet name.",
                }
        else:
            sheet = book.sheets.active

        shape = sheet.api.Shapes.AddTextbox(
            _ORIENT_HORIZONTAL,
            float(left),
            float(top),
            float(width),
            float(height),
        )

        # Set text first, then styling.
        try:
            shape.TextFrame2.TextRange.Text = str(text)
        except Exception:
            shape.TextFrame.Characters.Text = str(text)

        # Font styling — prefer TextFrame2 (modern), fall back to Characters.Font.
        try:
            font = shape.TextFrame2.TextRange.Font
            if font_size is not None:
                font.Size = float(font_size)
            if font_family is not None:
                font.Name = str(font_family)
            if font_color is not None:
                font.Fill.ForeColor.RGB = _hex_to_rgb_int(font_color)
        except Exception:  # noqa: BLE001 — fall back for older Excel
            try:
                font = shape.TextFrame.Characters.Font
                if font_size is not None:
                    font.Size = float(font_size)
                if font_family is not None:
                    font.Name = str(font_family)
                if font_color is not None:
                    font.Color = _hex_to_rgb_int(font_color)
            except Exception:  # noqa: BLE001
                pass

        if fill_color:
            try:
                shape.Fill.ForeColor.RGB = _hex_to_rgb_int(fill_color)
                shape.Fill.Solid()
            except Exception:  # noqa: BLE001
                pass

        if h_align:
            align_val = _ALIGN_MAP.get(h_align.lower() if isinstance(h_align, str) else "")
            if align_val is not None:
                try:
                    shape.TextFrame2.TextRange.ParagraphFormat.Alignment = align_val
                except Exception:
                    # xlHAlign on TextFrame: xlHAlignLeft=-4131, center=-4108, right=-4152
                    legacy = {1: -4131, 2: -4108, 3: -4152}.get(align_val)
                    if legacy is not None:
                        try:
                            shape.TextFrame.HorizontalAlignment = legacy
                        except Exception:  # noqa: BLE001
                            pass
    except Exception as exc:  # noqa: BLE001
        return {
            "status": "error",
            "message": f"Insert text box failed: {exc}",
            "error": str(exc),
        }

    on_sheet = f" on {sheet_name!r}" if sheet_name else ""
    return {
        "status": "success",
        "message": f"Inserted text box {preview!r} at ({left}, {top}){on_sheet}.",
        "outputs": {"sheet": sheet.name},
    }


registry.register("insertTextBox", handler, mutates=True)
