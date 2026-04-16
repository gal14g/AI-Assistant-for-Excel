"""namedRange — create / update / delete a workbook or sheet-scoped named range."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    op = (params.get("operation") or "create").lower()
    name = params.get("name")
    address = params.get("range")
    sheet_scope = params.get("sheetName")
    if not name:
        return {"status": "error", "message": "namedRange requires 'name'."}

    book = ctx.workbook_handle
    if ctx.dry_run:
        return {"status": "preview", "message": f"Would {op} named range {name!r}."}

    try:
        if op == "delete":
            target = book.names[name]
            target.delete()
            return {"status": "success", "message": f"Deleted named range {name!r}.", "outputs": {"name": name}}

        if not address:
            return {"status": "error", "message": "create/update named range requires 'range'."}
        rng = resolve_range(book, address)
        refers_to = f"={rng.sheet.name}!{rng.address}"
        if sheet_scope:
            # Sheet-scoped: prefix name with "Sheet1!Name" — matches Excel's internal format.
            full_name = f"{sheet_scope}!{name}"
        else:
            full_name = name
        # Remove existing with same name before re-creating, for idempotency.
        try:
            book.names[full_name].delete()
        except Exception:
            pass
        book.names.add(full_name, refers_to)
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"{op} failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"{op.capitalize()}d named range {name!r}.",
        "outputs": {"name": name},
    }


registry.register("namedRange", handler, mutates=True)
