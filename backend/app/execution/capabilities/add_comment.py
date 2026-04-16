"""addComment — attach a comment to a cell."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("cell") or params.get("range")
    text = params.get("comment") or params.get("text")
    if not address or not text:
        return {"status": "error", "message": "addComment requires 'cell' and 'comment'."}

    rng = resolve_range(ctx.workbook_handle, address)
    if ctx.dry_run:
        return {"status": "preview", "message": f"Would add comment to {rng.address}."}

    try:
        # Delete any existing comment first for idempotency.
        if rng.api.Comment is not None:
            rng.api.Comment.Delete()
        rng.api.AddComment(str(text))
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Add comment failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Added comment to {rng.address}.",
        "outputs": {"cell": rng.address},
    }


registry.register("addComment", handler, mutates=True)
