"""insertDeleteRows — insert or delete whole rows / columns."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    op = (params.get("operation") or "").lower()
    address = params.get("range") or params.get("target")
    axis = (params.get("axis") or "rows").lower()
    if op not in ("insert", "delete") or not address:
        return {"status": "error", "message": "insertDeleteRows requires 'operation' (insert|delete) and 'range'."}

    rng = resolve_range(ctx.workbook_handle, address)
    if ctx.dry_run:
        return {"status": "preview", "message": f"Would {op} {axis} at {rng.address}."}

    whole = rng.rows if axis.startswith("row") else rng.columns
    try:
        if op == "insert":
            whole.api.Insert()
        else:
            whole.api.Delete()
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"{op} failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"{op.capitalize()}ed {axis} at {rng.address}.",
        "outputs": {"range": rng.address},
    }


registry.register("insertDeleteRows", handler, mutates=True)
