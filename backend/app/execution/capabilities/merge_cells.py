"""mergeCells — merge or unmerge a range."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    # Accept both `operation` (historic) and `mergeType` (frontend/schema).
    # "unmerge" mergeType must flip operation to "unmerge".
    op = (params.get("operation") or "").lower()
    merge_type = (params.get("mergeType") or "").lower()
    if merge_type == "unmerge":
        op = "unmerge"
    elif not op:
        op = "merge"
    across = bool(params.get("across", False)) or merge_type == "mergeacross"
    if not address:
        return {"status": "error", "message": "mergeCells requires 'range'."}

    rng = resolve_range(ctx.workbook_handle, address)
    if ctx.dry_run:
        return {"status": "preview", "message": f"Would {op} {rng.address}."}

    try:
        if op == "unmerge":
            rng.api.UnMerge()
        else:
            rng.api.Merge(Across=across)
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"{op} failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"{op.capitalize()}d {rng.address}.",
        "outputs": {"range": rng.address},
    }


registry.register("mergeCells", handler, mutates=True)
