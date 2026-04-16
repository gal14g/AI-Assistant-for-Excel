"""consolidateAllSheets — merge data from all worksheets into one combined sheet.

Port of `frontend/src/engine/capabilities/consolidateAllSheets.ts`. Iterates
every worksheet in the workbook, reads used-range data, and writes a combined
matrix into a destination sheet (created if missing).
"""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry


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
    output_sheet_name = params.get("outputSheetName") or "Combined"
    has_headers = bool(params.get("hasHeaders", True))
    exclude_sheets = params.get("excludeSheets") or []

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would consolidate all sheets into {output_sheet_name!r}.",
        }

    try:
        book = ctx.workbook_handle
        exclude_set = {output_sheet_name, *exclude_sheets}

        combined: list[list[Any]] = []
        sheet_count = 0

        for sheet in book.sheets:
            if sheet.name in exclude_set:
                continue

            used = sheet.used_range
            # An empty sheet's used_range can still report shape (1, 1) with a
            # None value. Skip when there's no meaningful data.
            try:
                values = _as_2d(used.value, used.shape)
            except Exception:  # noqa: BLE001
                continue

            if not values:
                continue
            # Skip sheets whose entire used range is empty/None.
            if all(all(v is None or v == "" for v in row) for row in values):
                continue

            sheet_count += 1
            if not combined:
                # First non-empty sheet: include everything (headers + data).
                combined.extend(values)
            else:
                start_row = 1 if has_headers else 0
                for r in range(start_row, len(values)):
                    combined.append(values[r])

        if not combined:
            return {
                "status": "success",
                "message": "No data found in any qualifying sheets.",
                "outputs": {},
            }

        # Create or reuse the output sheet.
        try:
            out_sheet = book.sheets[output_sheet_name]
        except Exception:  # noqa: BLE001
            out_sheet = book.sheets.add(output_sheet_name)

        width = max(len(r) for r in combined)
        for row in combined:
            while len(row) < width:
                row.append(None)

        out_sheet.range("A1").resize(len(combined), width).value = combined
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"consolidateAllSheets failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": (
            f"Combined {sheet_count} sheets ({len(combined)} total rows) "
            f"into {output_sheet_name!r}."
        ),
        "outputs": {"outputRange": f"{output_sheet_name}!A1"},
    }


registry.register("consolidateAllSheets", handler, mutates=True)
