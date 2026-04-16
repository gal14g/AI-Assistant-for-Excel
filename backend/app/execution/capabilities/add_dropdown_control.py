"""addDropdownControl — create a dropdown (list validation) on a cell using literal values or a range reference."""

from __future__ import annotations

import re
from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


_V_LIST = 3       # xlValidateList
_ALERT_STOP = 1   # xlValidAlertStop

_SIMPLE_CELL_RE = re.compile(r"^[A-Z]{1,3}\d*$", re.IGNORECASE)
# Strip a leading [Book.xlsx] workbook qualifier from a range reference.
_WB_PREFIX_RE = re.compile(r"^\[[^\]]+\]")


def _is_range_reference(source: str) -> bool:
    trimmed = source.strip()
    if "!" in trimmed or ":" in trimmed:
        return True
    if _SIMPLE_CELL_RE.match(trimmed):
        return True
    return False


def _strip_workbook_qualifier(ref: str) -> str:
    """Remove `[Book.xlsx]` prefix from a range reference (Excel's Validation
    source can't reference an external workbook name in the formula)."""
    return _WB_PREFIX_RE.sub("", ref.strip())


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    cell = params.get("cell")
    list_source = params.get("listSource")
    prompt_message = params.get("promptMessage")
    sheet_name = params.get("sheetName")

    if not cell or not list_source:
        return {"status": "error", "message": "addDropdownControl requires 'cell' and 'listSource'."}

    # Qualify the cell address with sheetName when provided and not already qualified.
    cell_address = cell
    if sheet_name and "!" not in cell:
        cell_address = f"{sheet_name}!{cell}"

    rng = resolve_range(ctx.workbook_handle, cell_address)

    if ctx.dry_run:
        return {"status": "preview", "message": f"Would add dropdown control to {rng.address}."}

    # Build source and an informational count label.
    if _is_range_reference(list_source):
        source = "=" + _strip_workbook_qualifier(list_source)
        option_count_label = f"range {list_source}"
    else:
        source = list_source
        option_count_label = f"{len(list_source.split(','))} options"

    try:
        v = rng.api.Validation
        try:
            v.Delete()
        except Exception:
            pass
        v.Add(Type=_V_LIST, AlertStyle=_ALERT_STOP, Formula1=source)
        v.InCellDropdown = True
        if prompt_message:
            v.ShowInput = True
            v.InputTitle = "Select a value"
            v.InputMessage = prompt_message
    except Exception as exc:  # noqa: BLE001
        return {
            "status": "error",
            "message": f"Add dropdown failed: {exc}",
            "error": str(exc),
        }

    return {
        "status": "success",
        "message": f"Added dropdown control to {rng.address} with {option_count_label}.",
        "outputs": {"range": rng.address},
    }


registry.register("addDropdownControl", handler, mutates=True, affects_formatting=False)
