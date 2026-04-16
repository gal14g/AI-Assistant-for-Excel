"""setNumberFormat — apply a number format string to a range."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    fmt = params.get("format") or params.get("numberFormat")
    if not address or not fmt:
        return {"status": "error", "message": "setNumberFormat requires 'range' and 'format'."}

    rng = resolve_range(ctx.workbook_handle, address)
    if ctx.dry_run:
        return {"status": "preview", "message": f"Would set format {fmt!r} on {rng.address}."}

    rng.number_format = fmt
    return {
        "status": "success",
        "message": f"Set number format {fmt!r} on {rng.address}.",
        "outputs": {"range": rng.address, "format": fmt},
    }


registry.register("setNumberFormat", handler, mutates=False, affects_formatting=True)
