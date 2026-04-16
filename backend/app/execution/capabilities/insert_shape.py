"""
insertShape — insert a geometric shape via the COM Shapes.AddShape API.

Shape type names are mapped to MsoAutoShapeType integer values. The TS
handler uses Excel.GeometricShapeType which has a different enum layout;
we translate its string-keyed shape names into the matching MSO values.
"""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry


# MsoAutoShapeType values — see Microsoft docs.
_SHAPE_TYPES = {
    "rectangle": 1,      # msoShapeRectangle
    "oval": 9,           # msoShapeOval
    "ellipse": 9,        # msoShapeOval (TS exposes this alias)
    "diamond": 4,        # msoShapeDiamond
    "rightTriangle": 6,  # msoShapeRightTriangle
    "rightArrow": 33,    # msoShapeRightArrow
    "leftArrow": 34,     # msoShapeLeftArrow
    "upArrow": 35,       # msoShapeUpArrow
    "downArrow": 36,     # msoShapeDownArrow
    "star5": 92,         # msoShape5pointStar
    "heart": 21,         # msoShapeHeart
    "arrow": 33,         # alias → rightArrow
    "line": 9,           # fallback — proper lines need AddLine
}


def _hex_to_rgb_int(color: str) -> int:
    h = color.lstrip("#")
    if len(h) != 6:
        raise ValueError(f"Invalid color {color!r}; expected #RRGGBB.")
    return int(h[0:2], 16) + int(h[2:4], 16) * 256 + int(h[4:6], 16) * 65536


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    shape_type = params.get("shapeType")
    left = params.get("left")
    top = params.get("top")
    width = params.get("width")
    height = params.get("height")
    sheet_name = params.get("sheetName")
    fill_color = params.get("fillColor")
    line_color = params.get("lineColor")
    line_weight = params.get("lineWeight")
    text_content = params.get("textContent")

    if shape_type is None or left is None or top is None or width is None or height is None:
        return {
            "status": "error",
            "message": "insertShape requires 'shapeType', 'left', 'top', 'width', 'height'.",
        }

    mso_type = _SHAPE_TYPES.get(shape_type)
    if mso_type is None:
        return {
            "status": "error",
            "message": (
                f"Unknown shape type {shape_type!r}. Supported: "
                + ", ".join(sorted(_SHAPE_TYPES.keys()))
            ),
        }

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would insert {shape_type} shape at ({left}, {top}).",
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

        shape = sheet.api.Shapes.AddShape(
            mso_type, float(left), float(top), float(width), float(height)
        )

        if fill_color:
            try:
                shape.Fill.ForeColor.RGB = _hex_to_rgb_int(fill_color)
                shape.Fill.Solid()
            except Exception:  # noqa: BLE001
                pass
        if line_color:
            try:
                shape.Line.ForeColor.RGB = _hex_to_rgb_int(line_color)
            except Exception:  # noqa: BLE001
                pass
        if line_weight is not None:
            try:
                shape.Line.Weight = float(line_weight)
            except Exception:  # noqa: BLE001
                pass
        if text_content:
            try:
                shape.TextFrame2.TextRange.Text = str(text_content)
            except Exception:
                try:
                    shape.TextFrame.Characters.Text = str(text_content)
                except Exception:  # noqa: BLE001
                    pass
    except Exception as exc:  # noqa: BLE001
        return {
            "status": "error",
            "message": f"Insert shape failed: {exc}",
            "error": str(exc),
        }

    on_sheet = f" on {sheet_name!r}" if sheet_name else ""
    return {
        "status": "success",
        "message": f"Inserted {shape_type} shape at ({left}, {top}){on_sheet}.",
        "outputs": {"sheet": sheet.name},
    }


registry.register("insertShape", handler, mutates=True)
