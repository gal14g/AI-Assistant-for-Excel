"""tieredFormula — generate tier-based IFS formulas for tax/grade/commission tiers."""

from __future__ import annotations

import re
from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def _relative_cell_ref(address: str) -> str:
    """Strip any absolute-row / absolute-col markers from a cell address."""
    tail = address.split("!")[-1].split(":")[0]
    return tail.replace("$", "")


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    source_range = params.get("sourceRange")
    output_range = params.get("outputRange")
    tiers = params.get("tiers") or []
    mode = params.get("mode", "lookup")
    default_value = params.get("defaultValue", 0)
    has_headers = bool(params.get("hasHeaders", True))

    if not source_range or not output_range or not tiers:
        return {"status": "error", "message": "tieredFormula requires sourceRange, outputRange, tiers."}

    sorted_tiers = sorted(tiers, key=lambda t: float(t["threshold"]))

    src = resolve_range(ctx.workbook_handle, source_range)
    out = resolve_range(ctx.workbook_handle, output_range)

    if src.columns.count != 1 or out.columns.count != 1:
        return {"status": "error", "message": "sourceRange and outputRange must each be a single column."}

    src_top = src[0, 0]
    first_data_row = src_top.row + (1 if has_headers else 0)
    src_col_letter = re.match(r"^([A-Z]+)", src_top.address.split("!")[-1].replace("$", "")).group(1)
    src_cell_ref = f"{src_col_letter}{first_data_row}"

    src_rows_data = max(0, src.rows.count - (1 if has_headers else 0))
    out_rows_data = max(0, out.rows.count - (1 if has_headers else 0))
    n_rows = min(src_rows_data, out_rows_data)
    if n_rows == 0:
        return {"status": "success", "message": "No data rows to fill."}

    if mode == "lookup":
        parts = [f"{src_cell_ref}>={t['threshold']},{t['value']}" for t in reversed(sorted_tiers)]
        formula = f"=IFS({','.join(parts)},TRUE,{default_value})"
    elif mode == "tax":
        segments = []
        for i, t in enumerate(sorted_tiers):
            next_t = sorted_tiers[i + 1]["threshold"] if i + 1 < len(sorted_tiers) else None
            upper = src_cell_ref if next_t is None else f"MIN({src_cell_ref},{next_t})"
            segments.append(f"MAX(0,{upper}-{t['threshold']})*{t['value']}")
        formula = "=" + "+".join(segments)
    else:
        return {"status": "error", "message": f"Unknown mode: {mode}"}

    if ctx.dry_run:
        return {"status": "preview", "message": f"Would write {mode} tier formula for {n_rows} row(s)."}

    out_top = out[0, 0]
    out_first_row = out_top.row + (1 if has_headers else 0)
    out_col_letter = re.match(r"^([A-Z]+)", out_top.address.split("!")[-1].replace("$", "")).group(1)

    sheet = out.sheet
    try:
        anchor = sheet.range(f"{out_col_letter}{out_first_row}")
        anchor.formula = formula
        if n_rows > 1:
            target = sheet.range(f"{out_col_letter}{out_first_row}:{out_col_letter}{out_first_row + n_rows - 1}")
            # xlwings: set .formula on the target range with the same formula
            # relative refs auto-adjust.
            target.formula = [[formula]] * n_rows
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Write failed: {exc}", "error": str(exc)}

    out_addr = f"{sheet.name}!{out_col_letter}{out_first_row}:{out_col_letter}{out_first_row + n_rows - 1}"
    return {
        "status": "success",
        "message": f"Wrote {mode}-mode tier formulas ({len(tiers)} tier(s)) to {out_addr}.",
        "outputs": {"outputRange": out_addr},
    }


registry.register("tieredFormula", handler, mutates=True)
