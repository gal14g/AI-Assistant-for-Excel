"""formatCells — apply font / fill / alignment / border formatting to a range."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


# Excel HorizontalAlignment enum values.
_H_ALIGN = {"left": -4131, "center": -4108, "right": -4152, "general": 1}
_V_ALIGN = {"top": -4160, "center": -4108, "bottom": -4107}


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    if not address:
        return {"status": "error", "message": "formatCells requires 'range'."}

    rng = resolve_range(ctx.workbook_handle, address)
    if ctx.dry_run:
        return {"status": "preview", "message": f"Would format {rng.address}."}

    applied: list[str] = []
    try:
        font = params.get("font") or {}
        if font:
            if "name" in font:
                rng.api.Font.Name = font["name"]
                applied.append(f"font={font['name']}")
            if "size" in font:
                rng.api.Font.Size = font["size"]
                applied.append(f"size={font['size']}")
            if "bold" in font:
                rng.api.Font.Bold = bool(font["bold"])
                applied.append("bold" if font["bold"] else "not-bold")
            if "italic" in font:
                rng.api.Font.Italic = bool(font["italic"])
            if "color" in font:
                rng.font.color = font["color"]
                applied.append(f"color={font['color']}")

        fill = params.get("fill") or {}
        if fill.get("color"):
            rng.color = fill["color"]
            applied.append(f"fill={fill['color']}")

        align = params.get("alignment") or {}
        h = align.get("horizontal")
        if h in _H_ALIGN:
            rng.api.HorizontalAlignment = _H_ALIGN[h]
            applied.append(f"h-align={h}")
        v = align.get("vertical")
        if v in _V_ALIGN:
            rng.api.VerticalAlignment = _V_ALIGN[v]
            applied.append(f"v-align={v}")
        if "wrap" in align:
            rng.api.WrapText = bool(align["wrap"])

        borders = params.get("borders") or {}
        if borders:
            # 9 = xlEdgeLeft, 7 = xlEdgeTop, 10 = xlEdgeBottom, 11 = xlEdgeRight, 12 = xlInsideVertical, 14 = xlInsideHorizontal
            # We use Borders() with index 11 (xlInsideVertical) + 12 (xlInsideHorizontal) to apply all-cell borders.
            style = 1  # xlContinuous
            weight = 2  # xlThin
            for b_idx in (7, 10, 9, 11, 12, 14):
                try:
                    rng.api.Borders(b_idx).LineStyle = style
                    rng.api.Borders(b_idx).Weight = weight
                except Exception:
                    continue
            applied.append("borders")
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Format failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Formatted {rng.address} ({', '.join(applied) or 'no-op'}).",
        "outputs": {"range": rng.address},
    }


registry.register("formatCells", handler, mutates=False, affects_formatting=True)
