"""highlightDuplicates — conditional-format duplicate values in a range.

Uses the COM FormatConditions collection to add a duplicate-values rule
directly. On platforms without the DuplicateValues enum, falls back to a
COUNTIF-based custom formula.
"""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range
from app.execution.capabilities.tab_color import _hex_to_bgr_int  # reuse


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    fill_color = params.get("fillColor") or "#FFCCCC"
    font_color = params.get("fontColor") or "#C50F1F"
    if not address:
        return {"status": "error", "message": "highlightDuplicates requires 'range'."}

    rng = resolve_range(ctx.workbook_handle, address)

    if ctx.dry_run:
        return {"status": "preview", "message": f"Would highlight duplicates in {rng.address}."}

    try:
        # Try the DuplicateValues CF type first (COM xlDuplicate = 8,
        # AddTop10 enum int). Excel COM: FormatConditions.AddUniqueValues(),
        # then set DupeUnique = 0 (xlDuplicate) / 1 (xlUnique).
        cf = rng.api.FormatConditions.AddUniqueValues()
        cf.DupeUnique = 0  # xlDuplicate
        cf.Interior.Color = _hex_to_bgr_int(fill_color)
        cf.Font.Color = _hex_to_bgr_int(font_color)
    except Exception:
        # Fallback to custom formula.
        try:
            addr = rng.address
            # Use the top-left of the range as the anchor cell in COUNTIF.
            first_cell = rng[0, 0].address.replace("$", "")
            cf = rng.api.FormatConditions.Add(
                Type=2,  # xlExpression
                Formula1=f"=COUNTIF({addr},{first_cell})>1",
            )
            cf.Interior.Color = _hex_to_bgr_int(fill_color)
            cf.Font.Color = _hex_to_bgr_int(font_color)
        except Exception as exc:  # noqa: BLE001
            return {
                "status": "error",
                "message": f"Could not add duplicate-highlight rule: {exc}",
                "error": str(exc),
            }

    return {
        "status": "success",
        "message": f"Added duplicate-highlight rule on {rng.address}.",
        "outputs": {"range": rng.address},
    }


registry.register("highlightDuplicates", handler, mutates=False, affects_formatting=True)
