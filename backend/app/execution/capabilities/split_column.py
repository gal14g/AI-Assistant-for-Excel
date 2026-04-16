"""splitColumn — split a text column into multiple columns by a delimiter.

Ported from `frontend/src/engine/capabilities/splitColumn.ts`. Uses the source
range's start row so the output aligns row-for-row, and writes a single 2D
batch so merged/protected cells fail loudly rather than silently partial.
"""

from __future__ import annotations

import re
from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import parse_address, resolve_range


def _col_letters_to_index(col: str) -> int:
    n = 0
    for c in col.upper():
        n = n * 26 + (ord(c) - 64)
    return n - 1


def _index_to_col_letters(idx: int) -> str:
    s = ""
    idx += 1
    while idx > 0:
        r = (idx - 1) % 26
        s = chr(65 + r) + s
        idx = (idx - 1) // 26
    return s


def _to_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if not isinstance(values[0], list):
        return [values]
    return values


def _start_row_from_address(addr: str) -> int:
    """Pull the first row number out of `Sheet!$A$2:$A$50` or similar."""
    cell_part = addr.split("!")[-1] if "!" in addr else addr
    cell_part = cell_part.replace("$", "")
    m = re.match(r"[A-Za-z]+(\d+)", cell_part)
    return int(m.group(1)) if m else 1


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    source_range = params.get("sourceRange")
    delimiter = params.get("delimiter")
    output_start_column = params.get("outputStartColumn")
    output_headers = params.get("outputHeaders")
    parts = int(params.get("parts") or 2)

    if not source_range or delimiter is None or not output_start_column:
        return {
            "status": "error",
            "message": "splitColumn requires 'sourceRange', 'delimiter', and 'outputStartColumn'.",
        }

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": (
                f"Would split {source_range} by '{delimiter}' into {parts} columns "
                f"starting at {output_start_column}."
            ),
        }

    src = resolve_range(ctx.workbook_handle, source_range)
    vals = _to_2d(src.value)
    if not vals:
        return {"status": "success", "message": "No data to split.", "outputs": {"outputRange": None}}

    start_row = _start_row_from_address(src.address)

    out_grid: list[list[Any]] = []
    for row in vals:
        first_cell = row[0] if row else ""
        cell = "" if first_cell is None else str(first_cell)
        split_parts = cell.split(delimiter)[:parts]
        while len(split_parts) < parts:
            split_parts.append("")
        out_grid.append([p.strip() for p in split_parts])

    # Overwrite the first row with headers if provided (matches TS behavior).
    if output_headers and out_grid:
        for c in range(min(len(output_headers), parts)):
            out_grid[0][c] = output_headers[c]

    start_col_idx = _col_letters_to_index(output_start_column)
    end_col_letters = _index_to_col_letters(start_col_idx + parts - 1)
    start_col_letters = _index_to_col_letters(start_col_idx)

    # Keep output on the same sheet as the source range.
    parsed = parse_address(source_range)
    sheet_prefix = ""
    if parsed.sheet:
        needs_quote = any(ch in parsed.sheet for ch in " !'")
        sheet_prefix = f"'{parsed.sheet}'!" if needs_quote else f"{parsed.sheet}!"

    out_addr = (
        f"{sheet_prefix}{start_col_letters}{start_row}:"
        f"{end_col_letters}{start_row + len(out_grid) - 1}"
    )

    try:
        out_rng = resolve_range(ctx.workbook_handle, out_addr)
        out_rng.value = out_grid
    except Exception as exc:  # noqa: BLE001
        return {
            "status": "error",
            "message": (
                f"Failed to write split results: {exc}. Range may contain merged or protected cells."
            ),
            "error": str(exc),
        }

    # Best-effort autofit of the new columns (non-fatal).
    try:
        out_rng.columns.autofit()
    except Exception:  # noqa: BLE001
        pass

    return {
        "status": "success",
        "message": (
            f"Split {len(vals)} rows from {src.address} into {parts} columns "
            f"starting at {output_start_column}."
        ),
        "outputs": {"outputRange": out_rng.address},
    }


registry.register("splitColumn", handler, mutates=True)
