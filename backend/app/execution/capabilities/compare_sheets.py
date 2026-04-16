"""compareSheets — diff two ranges, highlight diffs or write a diff report.

Port of `frontend/src/engine/capabilities/compareSheets.ts`. Two modes:

    highlightDiffs=True  → color differing cells in rangeA in place.
    highlightDiffs=False → write a Row | Col | ValueA | ValueB report to
                           outputRange, or to a new "Diff_Report" sheet if
                           outputRange is omitted.
"""

from __future__ import annotations

import re
from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


_HEX_RE = re.compile(r"^#?([0-9a-fA-F]{6})$")


def _hex_to_rgb(color: str) -> tuple[int, int, int]:
    m = _HEX_RE.match(color.strip())
    if not m:
        return (255, 217, 102)  # fallback to #FFD966
    h = m.group(1)
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def _as_2d(raw: Any, shape: tuple[int, int]) -> list[list[Any]]:
    if raw is None:
        return []
    rows_, cols_ = shape
    if not isinstance(raw, list):
        return [[raw]]
    if raw and not isinstance(raw[0], list):
        if rows_ == 1:
            return [list(raw)]
        return [[v] for v in raw]
    return [list(r) for r in raw]


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    range_a = params.get("rangeA")
    range_b = params.get("rangeB")
    highlight_diffs = bool(params.get("highlightDiffs", False))
    highlight_color = params.get("highlightColor") or "#FFD966"
    output_range = params.get("outputRange")

    if not range_a or not range_b:
        return {"status": "error", "message": "compareSheets requires 'rangeA' and 'rangeB'."}

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would compare {range_a} vs {range_b}.",
        }

    try:
        rng_a = resolve_range(ctx.workbook_handle, range_a)
        rng_b = resolve_range(ctx.workbook_handle, range_b)

        vals_a = _as_2d(rng_a.value, rng_a.shape)
        vals_b = _as_2d(rng_b.value, rng_b.shape)

        rows = max(len(vals_a), len(vals_b))
        cols = max(
            len(vals_a[0]) if vals_a else 0,
            len(vals_b[0]) if vals_b else 0,
        )

        diffs: list[tuple[int, int, str, str]] = []
        for r in range(rows):
            for c in range(cols):
                a = ""
                if r < len(vals_a) and c < len(vals_a[r]):
                    a_val = vals_a[r][c]
                    a = str(a_val if a_val is not None else "")
                b = ""
                if r < len(vals_b) and c < len(vals_b[r]):
                    b_val = vals_b[r][c]
                    b = str(b_val if b_val is not None else "")
                if a != b:
                    diffs.append((r + 1, c + 1, a, b))

        if highlight_diffs:
            rgb = _hex_to_rgb(highlight_color)
            highlighted = 0
            for d_row, d_col, _a, _b in diffs:
                try:
                    cell = rng_a[d_row - 1, d_col - 1]
                    cell.color = rgb
                    highlighted += 1
                except Exception:  # noqa: BLE001 — merged / protected cells
                    continue

            skipped = len(diffs) - highlighted
            tail = f" ({skipped} skipped — merged/protected)" if skipped > 0 else ""
            return {
                "status": "success",
                "message": f"Highlighted {highlighted} difference(s) in {range_a}{tail}.",
                "outputs": {"outputRange": range_a},
            }

        # Diff report
        start_row = rng_a.row
        # Compute start column letter from rngA's address
        a_addr = rng_a.address
        cell_part = a_addr.split("!")[-1] if "!" in a_addr else a_addr
        m = re.match(r"\$?([A-Z]+)", cell_part.replace("$", ""))
        start_col_letter = m.group(1) if m else "A"

        report_rows: list[list[Any]] = [["Row", "Column", range_a, range_b]]
        for d_row, d_col, a_str, b_str in diffs:
            # Match TS behavior: Column column is `startCol + (col-1)` — a
            # string/number concat on the TS side. We just emit the numeric
            # column offset prefixed by the start column letter for clarity.
            col_repr = f"{start_col_letter}{d_col - 1}" if start_col_letter else d_col - 1
            report_rows.append([start_row + d_row - 1, col_repr, a_str, b_str])

        if not diffs:
            report_rows.append(["No differences found", "", "", ""])

        if output_range:
            out_rng = resolve_range(ctx.workbook_handle, output_range)
            out_addr_resolved = output_range
        else:
            # Create (or reuse) Diff_Report sheet.
            book = ctx.workbook_handle
            try:
                sheet = book.sheets["Diff_Report"]
            except Exception:  # noqa: BLE001
                sheet = book.sheets.add("Diff_Report")
            out_rng = sheet.range("A1")
            out_addr_resolved = "Diff_Report!A1"

        out_rng.resize(len(report_rows), 4).value = report_rows
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"compareSheets failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Found {len(diffs)} difference(s) between {range_a} and {range_b}.",
        "outputs": {"outputRange": out_addr_resolved},
    }


registry.register("compareSheets", handler, mutates=True, affects_formatting=True)
