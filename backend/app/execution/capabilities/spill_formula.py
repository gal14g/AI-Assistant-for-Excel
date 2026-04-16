"""
spillFormula — write a dynamic-array formula that Excel spills automatically.

Modern desktop Excel (365 / 2021+) handles SPILL natively. We prefer
`Range.Formula2` (the dynamic-array-aware setter) and fall back to the
legacy `formula` property if Formula2 isn't exposed by this Excel build.
Post-write we read the cell back to surface formula errors early.
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
    sheet_name = params.get("sheetName")

    if not cell or not formula:
        return {
            "status": "error",
            "message": "spillFormula requires 'cell' and 'formula'.",
        }

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would write dynamic array formula to {cell}: {formula}",
        }

    # Qualify the cell with the sheet if specified and not already qualified.
    cell_address = (
        f"{sheet_name}!{cell}"
        if sheet_name and "!" not in cell and not cell.startswith("[[")
        else cell
    )

    try:
        rng = resolve_range(ctx.workbook_handle, cell_address)

        # Prefer Formula2 (dynamic-array aware). xlwings doesn't expose it
        # directly, so go through the COM .api on Windows. On macOS the
        # AppleScript bridge may not support Formula2 — fall back to .formula.
        wrote_via_formula2 = False
        try:
            rng.api.Formula2 = formula
            wrote_via_formula2 = True
        except Exception:  # noqa: BLE001
            try:
                rng.formula2 = formula  # some xlwings versions expose this
                wrote_via_formula2 = True
            except Exception:  # noqa: BLE001
                rng.formula = formula

        # Read back the anchor cell to detect formula errors. Single-cell reads
        # return a scalar, not a 2D list.
        first_val = str(rng.value)
        if any(err in first_val for err in _ERRS):
            return {
                "status": "error",
                "message": (
                    f"Formula wrote to {cell} but produced {first_val}. "
                    "The formula may need to be corrected."
                ),
                "error": f"Formula error: {first_val}",
            }

        # Best-effort detection of the spill range size via the COM
        # SpillingToRange property. Harmless if unsupported.
        spill_info = "spill size unknown"
        try:
            spill = rng.api.SpillingToRange
            if spill is not None:
                total = int(spill.Rows.Count) * int(spill.Columns.Count)
                spill_info = f"spilled to {total} cells ({spill.Address})"
        except Exception:  # noqa: BLE001
            pass
    except Exception as exc:  # noqa: BLE001
        return {
            "status": "error",
            "message": f"Spill formula failed: {exc}",
            "error": str(exc),
        }

    _unused = wrote_via_formula2  # noqa: F841 — kept for debugging clarity
    return {
        "status": "success",
        "message": f"Wrote dynamic array formula to {cell}, {spill_info}.",
        "outputs": {"cell": cell_address},
    }


registry.register("spillFormula", handler, mutates=True)
