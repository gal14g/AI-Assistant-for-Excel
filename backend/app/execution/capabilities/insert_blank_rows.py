"""insertBlankRows — insert blank rows at explicit row numbers or every Nth row."""

from __future__ import annotations

import re
from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry


def _parse_first_row(address: str) -> int:
    """Extract the starting row number from a range address like 'A2:D100'."""
    tail = address.split("!")[-1].split(":")[0]
    m = re.match(r"^\$?[A-Z]+\$?(\d+)$", tail)
    return int(m.group(1)) if m else 1


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    sheet_name = params.get("sheetName")
    positions = params.get("positions")
    every = params.get("every")
    range_addr = params.get("range")
    count = int(params.get("count") or 1)

    book = ctx.workbook_handle
    sheet = book.sheets[sheet_name] if sheet_name else book.sheets.active

    # Compute the 1-based row numbers to insert before.
    targets: list[int] = []
    if positions:
        targets = [int(p) for p in positions]
    elif every and range_addr:
        rng = sheet.range(range_addr)
        first_row = _parse_first_row(rng.address)
        total_rows = rng.rows.count
        for i in range(int(every), total_rows, int(every)):
            targets.append(first_row + i)
    else:
        return {
            "status": "error",
            "message": "Provide either 'positions' OR 'every' + 'range'.",
        }

    if not targets:
        return {"status": "success", "message": "No positions to insert at."}

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would insert {count * len(targets)} blank row(s) at {len(targets)} position(s).",
        }

    # Descending so earlier inserts don't shift later positions.
    targets.sort(reverse=True)

    try:
        for row in targets:
            # Select rows [row, row + count) and insert.
            row_range = sheet.api.Rows(f"{row}:{row + count - 1}")
            row_range.Insert()
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Insert failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Inserted {count * len(targets)} blank row(s) at {len(targets)} position(s).",
        "outputs": {"rowsInserted": count * len(targets)},
    }


registry.register("insertBlankRows", handler, mutates=True)
