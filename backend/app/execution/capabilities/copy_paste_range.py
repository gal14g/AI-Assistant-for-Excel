"""copyPasteRange — copy a source range to a destination range."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    src = params.get("sourceRange") or params.get("source")
    dst = params.get("destinationRange") or params.get("destination")
    paste_type = (params.get("pasteType") or "all").lower()
    if not src or not dst:
        return {"status": "error", "message": "copyPasteRange requires 'sourceRange' and 'destinationRange'."}

    source = resolve_range(ctx.workbook_handle, src)
    destination = resolve_range(ctx.workbook_handle, dst)

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would copy {source.address} → {destination.address} ({paste_type}).",
        }

    # xlwings Range.copy(destination) defaults to "all" (values + formats +
    # formulas). For values-only / formats-only we fall back to property
    # assignment.
    if paste_type in ("all", "allwithoutborders"):
        source.copy(destination)
    elif paste_type == "values":
        destination.value = source.value
    elif paste_type in ("formats", "format"):
        destination.number_format = source.number_format
        try:
            destination.color = source.color
        except Exception:
            pass
    elif paste_type == "formulas":
        destination.formula = source.formula
    else:
        return {"status": "error", "message": f"Unknown pasteType {paste_type!r}."}

    return {
        "status": "success",
        "message": f"Copied {source.address} → {destination.address} ({paste_type}).",
        "outputs": {"sourceRange": source.address, "destinationRange": destination.address},
    }


registry.register("copyPasteRange", handler, mutates=True)
