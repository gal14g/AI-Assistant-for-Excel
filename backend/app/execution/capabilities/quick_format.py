"""quickFormat — compound formatting: freeze header row, add auto-filters, style header, autofit columns, optional zebra-stripe."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


_H_CENTER = -4108  # xlCenter


def _hex_to_rgb_int(hex_color: str) -> int:
    """Convert "#RRGGBB" → Excel long int (BGR packed)."""
    h = hex_color.lstrip("#")
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    return r + g * 256 + b * 65536


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    if not address:
        return {"status": "error", "message": "quickFormat requires 'range'."}

    freeze_header = params.get("freezeHeader", True)
    add_filters = params.get("addFilters", True)
    auto_fit = params.get("autoFit", True)
    zebra_stripe = params.get("zebraStripe", False)
    header_color = params.get("headerColor", "#4472C4")
    header_font_color = params.get("headerFontColor", "#FFFFFF")

    book = ctx.workbook_handle
    rng = resolve_range(book, address)
    sheet = rng.sheet

    if ctx.dry_run:
        return {"status": "preview", "message": f"Would apply quick format to {rng.address}."}

    applied: list[str] = []

    try:
        # Row count — for zebra-stripe iteration.
        if rng.count == 1:
            row_count = 1
        else:
            try:
                row_count = rng.shape[0]
            except Exception:
                row_count = rng.api.Rows.Count

        # 1) Freeze header row — activate sheet, set SplitRow=1, FreezePanes=True.
        if freeze_header:
            try:
                sheet.activate()
                window = book.app.api.ActiveWindow
                window.FreezePanes = False
                window.SplitRow = 1
                window.SplitColumn = 0
                window.FreezePanes = True
                applied.append("freeze: Y")
            except Exception:
                applied.append("freeze: N")
        else:
            applied.append("freeze: N")

        # 2) Add auto-filter — AutoFilter is invoked on the range itself.
        if add_filters:
            try:
                # Clear any existing AutoFilter on the sheet first — only one is allowed.
                if sheet.api.AutoFilterMode:
                    sheet.api.AutoFilterMode = False
                rng.api.AutoFilter()
                applied.append("filters: Y")
            except Exception:
                applied.append("filters: N")
        else:
            applied.append("filters: N")

        # 3) Format header row (first row of the range).
        try:
            header_api = rng.api.Rows(1)
            header_api.Font.Bold = True
            header_api.Font.Color = _hex_to_rgb_int(header_font_color)
            header_api.Interior.Color = _hex_to_rgb_int(header_color)
            header_api.HorizontalAlignment = _H_CENTER
        except Exception:
            pass

        # 4) Auto-fit columns.
        if auto_fit:
            try:
                rng.api.Columns.AutoFit()
                applied.append("autofit: Y")
            except Exception:
                applied.append("autofit: N")
        else:
            applied.append("autofit: N")

        # 5) Zebra-stripe data rows.
        if zebra_stripe:
            even_rgb = _hex_to_rgb_int("#F2F2F2")
            odd_rgb = _hex_to_rgb_int("#FFFFFF")
            for r in range(1, row_count):
                color = even_rgb if (r - 1) % 2 == 0 else odd_rgb
                try:
                    rng.api.Rows(r + 1).Interior.Color = color
                except Exception:
                    continue
            applied.append("zebra: Y")
        else:
            applied.append("zebra: N")

    except Exception as exc:  # noqa: BLE001
        return {
            "status": "error",
            "message": f"Quick format failed: {exc}",
            "error": str(exc),
        }

    return {
        "status": "success",
        "message": f"Applied quick format ({', '.join(applied)}).",
        "outputs": {"range": rng.address},
    }


registry.register("quickFormat", handler, mutates=True, affects_formatting=True)
