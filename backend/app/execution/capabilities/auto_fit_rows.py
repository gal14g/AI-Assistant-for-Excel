"""autoFitRows — auto-fit row heights on a range (or the sheet's used range)."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    sheet_name = params.get("sheetName")

    if address:
        rng = resolve_range(ctx.workbook_handle, address)
    else:
        book = ctx.workbook_handle
        sheet = book.sheets[sheet_name] if sheet_name else book.sheets.active
        rng = sheet.used_range

    if ctx.dry_run:
        return {"status": "preview", "message": f"Would autofit rows on {rng.address}."}

    try:
        rng.rows.autofit()
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"AutoFit rows failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Autofit rows on {rng.address}.",
        "outputs": {"range": rng.address},
    }


registry.register("autoFitRows", handler, mutates=False, affects_formatting=True)
