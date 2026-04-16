"""rankColumn — write RANK formulas for each value in a source column.

For each data row, writes =RANK(<srcCol><row>,$<srcCol>$<first>:$<srcCol>$<last>,<order>)
where order = 0 for descending (default) and 1 for ascending. Preserves the
header row when hasHeaders is true and labels the output column "Rank".
"""

from __future__ import annotations

import re
from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


_COL_RE = re.compile(r"[A-Z]+", re.IGNORECASE)


def _column_letter(ref: str) -> str:
    stripped = ref.replace("$", "")
    if "!" in stripped:
        stripped = stripped.split("!", 1)[1]
    m = _COL_RE.search(stripped)
    return m.group(0).upper() if m else "A"


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    source = params.get("sourceRange")
    output = params.get("outputRange")
    order = (params.get("order") or "descending").lower()
    has_headers = bool(params.get("hasHeaders", True))

    if not source or not output:
        return {
            "status": "error",
            "message": "rankColumn requires 'sourceRange' and 'outputRange'.",
        }

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would write {order} rank formulas from {source} to {output}.",
        }

    try:
        src_rng = resolve_range(ctx.workbook_handle, source)
        sheet = src_rng.sheet

        try:
            used = src_rng.current_region
            start_row = used.row
            last_row = start_row + used.rows.count - 1
        except Exception:
            used = sheet.used_range
            last_row = used.row + used.rows.count - 1

        first_data_row = 2 if has_headers else 1
        if last_row < first_data_row:
            return {"status": "success", "message": "No data rows found.", "outputs": {}}

        src_col = _column_letter(source)
        out_col = _column_letter(output)
        rank_order = 1 if order == "ascending" else 0

        if has_headers:
            sheet.range(f"{out_col}1").value = "Rank"

        formulas = [
            [
                f"=RANK({src_col}{r},${src_col}${first_data_row}:${src_col}${last_row},{rank_order})"
            ]
            for r in range(first_data_row, last_row + 1)
        ]
        out_rng = sheet.range(f"{out_col}{first_data_row}:{out_col}{last_row}")
        out_rng.formula = formulas
        row_count = last_row - first_data_row + 1
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"rankColumn failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Created {row_count} rank formulas in {output}.",
        "outputs": {"outputRange": output},
    }


registry.register("rankColumn", handler, mutates=True)
