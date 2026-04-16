"""
writeFormula — write a formula to a cell, optionally fill down.

Dynamic-array functions (FILTER, UNIQUE, XLOOKUP, SORT, SEQUENCE) are
supported natively by modern desktop Excel running under xlwings, so we
don't need the Item 3 dynamic-array rewriter here. If the user's desktop
Excel happens to be 2016/2019 the formula will evaluate to #NAME? — the
post-write error check catches that and reports it back through MCP.
"""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


_ERRS = ("#SPILL!", "#REF!", "#VALUE!", "#NAME?", "#NULL!", "#N/A", "#DIV/0!", "#CALC!")


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    cell = params.get("cell")
    formula = params.get("formula")
    fill_down = params.get("fillDown")
    if not cell or not formula:
        return {"status": "error", "message": "writeFormula requires 'cell' and 'formula'."}

    if ctx.dry_run:
        suffix = f" and fill down {fill_down} rows" if fill_down else ""
        return {
            "status": "preview",
            "message": f"Would write formula {formula} to {cell}{suffix}.",
        }

    rng = resolve_range(ctx.workbook_handle, cell)
    rng.formula = formula

    if fill_down and fill_down > 1:
        # xlwings doesn't expose autoFill directly. Mimic Excel's AutoFill
        # semantics by copying the formula down — relative references adjust
        # automatically because we use Range.copy/Range.paste (not literal
        # string duplication).
        source = rng
        target = rng.resize(fill_down, 1)
        source.copy(target)

    # Post-write error check — read back the first cell's evaluated value.
    first_val = str(rng.value) if rng.count == 1 else str((rng.value[0] if isinstance(rng.value, list) else rng.value))
    if any(err in first_val for err in _ERRS):
        return {
            "status": "error",
            "message": f"Formula wrote to {cell} but produced {first_val}.",
            "error": f"Formula error: {first_val}",
        }

    filled = f" (filled down {fill_down} rows)" if fill_down else ""
    return {
        "status": "success",
        "message": f"Wrote formula to {rng.address}{filled}.",
        "outputs": {"cell": rng.address},
    }


registry.register("writeFormula", handler, mutates=True)
