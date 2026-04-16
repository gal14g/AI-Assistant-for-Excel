"""createTable — convert a range to an Excel table (ListObject)."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    name = params.get("tableName") or params.get("name")
    has_headers = bool(params.get("hasHeaders", True))
    style = params.get("tableStyle") or "TableStyleMedium2"
    if not address:
        return {"status": "error", "message": "createTable requires 'range'."}

    rng = resolve_range(ctx.workbook_handle, address)
    if ctx.dry_run:
        return {"status": "preview", "message": f"Would create table on {rng.address}."}

    try:
        # Sheet.api.ListObjects.Add(SourceType=xlSrcRange, Source=range, XlListObjectHasHeaders=xlYes)
        list_object = rng.sheet.api.ListObjects.Add(
            1,  # xlSrcRange
            rng.api,
            None,
            1 if has_headers else 2,  # xlYes | xlNo
        )
        if name:
            list_object.Name = name
        list_object.TableStyle = style
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Create table failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Created table on {rng.address}.",
        "outputs": {"tableName": list_object.Name, "range": rng.address},
    }


registry.register("createTable", handler, mutates=True, affects_formatting=True)
