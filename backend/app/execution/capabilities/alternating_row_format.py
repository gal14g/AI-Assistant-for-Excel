"""alternatingRowFormat — apply zebra-stripe (alternating fill colors) across a range's rows, with optional bold header styling."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


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
        return {"status": "error", "message": "alternatingRowFormat requires 'range'."}

    even_color = params.get("evenColor", "#F2F2F2")
    odd_color = params.get("oddColor", "#FFFFFF")
    has_headers = params.get("hasHeaders", True)

    rng = resolve_range(ctx.workbook_handle, address)

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would apply alternating row colors to {rng.address}.",
        }

    try:
        # Row count — handles single-row ranges.
        if rng.count == 1:
            row_count = 1
        else:
            try:
                row_count = rng.shape[0]
            except Exception:
                row_count = rng.api.Rows.Count

        start_row = 1 if has_headers else 0
        even_rgb = _hex_to_rgb_int(even_color)
        odd_rgb = _hex_to_rgb_int(odd_color)
        header_rgb = _hex_to_rgb_int("#D9E2F3")

        # Format the header row if applicable.
        if has_headers and row_count > 0:
            header_api = rng.api.Rows(1)
            try:
                header_api.Font.Bold = True
                header_api.Interior.Color = header_rgb
            except Exception:
                pass

        # Apply alternating colors to data rows. rng.api.Rows is 1-based.
        for r in range(start_row, row_count):
            data_row_index = r - start_row
            color = even_rgb if data_row_index % 2 == 0 else odd_rgb
            try:
                rng.api.Rows(r + 1).Interior.Color = color
            except Exception:
                continue

        data_rows = row_count - start_row

    except Exception as exc:  # noqa: BLE001
        return {
            "status": "error",
            "message": f"Alternating row format failed: {exc}",
            "error": str(exc),
        }

    return {
        "status": "success",
        "message": f"Applied alternating row colors to {data_rows} rows.",
        "outputs": {"range": rng.address, "rowsFormatted": data_rows},
    }


registry.register("alternatingRowFormat", handler, mutates=True, affects_formatting=True)
