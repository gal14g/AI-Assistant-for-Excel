"""findReplace — find-and-replace values in a range."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    find_text = params.get("findText")
    replace_text = params.get("replaceText", "")
    match_case = bool(params.get("matchCase", False))
    whole_cell = bool(params.get("wholeCell", False))
    if not address or find_text is None:
        return {"status": "error", "message": "findReplace requires 'range' and 'findText'."}

    rng = resolve_range(ctx.workbook_handle, address)
    if ctx.dry_run:
        return {"status": "preview", "message": f"Would replace {find_text!r} → {replace_text!r} in {rng.address}."}

    # xlwings exposes Range.api.Replace (Excel COM) with full option surface.
    try:
        rng.api.Replace(
            What=find_text,
            Replacement=replace_text,
            LookAt=1 if whole_cell else 2,  # xlWhole | xlPart
            MatchCase=match_case,
        )
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Replace failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Replaced {find_text!r} → {replace_text!r} in {rng.address}.",
        "outputs": {"range": rng.address},
    }


registry.register("findReplace", handler, mutates=True)
