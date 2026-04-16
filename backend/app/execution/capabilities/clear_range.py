"""clearRange — wipe values / formats / both from a range."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    mode = (params.get("clearType") or "contents").lower()
    if not address:
        return {"status": "error", "message": "clearRange requires 'range'."}

    rng = resolve_range(ctx.workbook_handle, address)

    if ctx.dry_run:
        return {"status": "preview", "message": f"Would clear {mode} of {rng.address}."}

    # xlwings `.clear()` wipes values + formats, `.clear_contents()` wipes
    # only values. Match the frontend's clearType options.
    if mode in ("contents", "values"):
        rng.clear_contents()
    elif mode in ("formats", "format"):
        # xlwings has no clear_formats shortcut — reset number format + interior.
        rng.number_format = "General"
        try:
            rng.color = None  # cell fill
        except Exception:
            pass
    elif mode == "all":
        rng.clear()
    else:
        return {"status": "error", "message": f"Unknown clearType {mode!r}."}

    return {
        "status": "success",
        "message": f"Cleared {mode} of {rng.address}.",
        "outputs": {"range": rng.address},
    }


registry.register("clearRange", handler, mutates=True)
