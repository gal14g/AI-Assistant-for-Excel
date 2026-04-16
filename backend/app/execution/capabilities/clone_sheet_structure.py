"""cloneSheetStructure — duplicate a sheet's headers + formatting, drop the data.

Port of `frontend/src/engine/capabilities/cloneSheetStructure.ts`. Uses Excel's
native COM `Worksheet.Copy(After:=...)` so column widths, cell formatting, and
the header row are preserved — then clears the used range contents below the
header row.
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
    source_sheet_name = params.get("sourceSheet")
    new_sheet_name = params.get("newSheetName")

    if not source_sheet_name or not new_sheet_name:
        return {
            "status": "error",
            "message": "cloneSheetStructure requires 'sourceSheet' and 'newSheetName'.",
        }

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would clone structure of {source_sheet_name!r} → {new_sheet_name!r}.",
        }

    try:
        book = ctx.workbook_handle
        try:
            source = book.sheets[source_sheet_name]
        except Exception as exc:  # noqa: BLE001
            return {"status": "error", "message": f"Source sheet {source_sheet_name!r} not found.", "error": str(exc)}

        # Copy the sheet after the last sheet, preserving formatting.
        last_sheet = book.sheets[len(book.sheets) - 1]
        source.api.Copy(After=last_sheet.api)

        # xlwings doesn't return the new sheet from Copy — find the active one,
        # which Excel sets to the freshly-created copy.
        copy = book.sheets.active
        copy.name = new_sheet_name

        # Preserve row-1 (header) values, clear contents below.
        used = copy.used_range
        try:
            values = _as_2d(used.value, used.shape)
            row_count, col_count = used.shape
        except Exception:  # noqa: BLE001
            values = []
            row_count = col_count = 0

        if values and row_count > 1:
            headers = values[0] if values and values[0] else []
            used.clear_contents()
            if headers:
                width = len(headers)
                copy.range("A1").resize(1, width).value = [list(headers)]
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"cloneSheetStructure failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": (
            f"Cloned structure of {source_sheet_name!r} → {new_sheet_name!r} "
            f"(headers + formatting, no data)."
        ),
        "outputs": {"sheet": new_sheet_name},
    }


registry.register("cloneSheetStructure", handler, mutates=True)
