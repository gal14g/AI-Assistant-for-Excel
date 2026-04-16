"""addHyperlink — add a hyperlink to a cell."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("cell") or params.get("range")
    url = params.get("url") or params.get("address")
    display = params.get("displayText") or url
    if not address or not url:
        return {"status": "error", "message": "addHyperlink requires 'cell' and 'url'."}

    rng = resolve_range(ctx.workbook_handle, address)
    if ctx.dry_run:
        return {"status": "preview", "message": f"Would add hyperlink to {rng.address}."}

    try:
        rng.sheet.api.Hyperlinks.Add(Anchor=rng.api, Address=url, TextToDisplay=display)
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Add hyperlink failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Added hyperlink to {rng.address}: {url}.",
        "outputs": {"cell": rng.address, "url": url},
    }


registry.register("addHyperlink", handler, mutates=True)
