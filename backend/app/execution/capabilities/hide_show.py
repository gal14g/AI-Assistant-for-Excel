"""hideShow — hide or unhide rows / columns / sheets."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    target_kind = (params.get("targetKind") or "rows").lower()  # rows|columns|sheet
    op = (params.get("operation") or "hide").lower()
    address = params.get("range")
    sheet_name = params.get("sheetName")

    if ctx.dry_run:
        return {"status": "preview", "message": f"Would {op} {target_kind}."}

    book = ctx.workbook_handle
    try:
        if target_kind == "sheet":
            sheet = book.sheets[sheet_name] if sheet_name else book.sheets.active
            sheet.visible = op == "show"
        else:
            if not address:
                return {"status": "error", "message": "hideShow on rows/columns requires 'range'."}
            rng = resolve_range(book, address)
            collection = rng.rows if target_kind.startswith("row") else rng.columns
            collection.api.Hidden = op == "hide"
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"{op} {target_kind} failed: {exc}", "error": str(exc)}

    return {"status": "success", "message": f"{op.capitalize()}d {target_kind}."}


registry.register("hideShow", handler, mutates=False, affects_formatting=True)
