"""
Pre-mutation snapshots for xlwings-based plan execution.

Before any mutating step runs, we capture the `values`, `formulas`, and
`number_format` of every range the step will touch. The snapshots are
pushed onto the `ExecutorContext.snapshot_stack`, keyed by the ranges they
cover. `undo_last()` pops the top of that stack and restores.

Design notes:
- 20-entry in-memory stack per workbook (matches the frontend's rolling
  window — older snapshots are dropped).
- We snapshot at the *Range* granularity, not the Worksheet, so a plan that
  only rewrites A1:C10 doesn't need to hoist the entire sheet into memory.
- xlwings round-trips values as native Python types (datetime, float, str),
  which lets us persist / compare snapshots without format conversion.
- For `dry_run=True` we skip snapshotting entirely — the run doesn't mutate.

Wire-format:
    {
      "book":   "Budget.xlsx",
      "ranges": [
          {"sheet": "Sheet1", "address": "A1:C10",
           "values": [...], "formulas": [...], "number_format": [...]}
      ],
      "timestamp": "2026-04-16T12:34:56"
    }
"""

from __future__ import annotations

from datetime import datetime
from typing import TYPE_CHECKING, Any, Iterable

if TYPE_CHECKING:
    import xlwings as xw


# Matches the frontend's `MAX_SNAPSHOT_ENTRIES` — see
# `frontend/src/engine/snapshot.ts`.
MAX_SNAPSHOT_STACK = 20


def capture_snapshot(
    book: "xw.Book",
    addresses: Iterable[str],
    *,
    default_sheet: str | None = None,
) -> dict[str, Any]:
    """
    Capture values / formulas / number_format for each address in `addresses`.
    Returns a serialisable dict that can be pushed onto the snapshot stack
    and later replayed by `restore_snapshot`.

    `addresses` may contain bare cells, ranges, sheet-qualified ranges, or
    cross-workbook references — whatever `range_utils.resolve_range` accepts.
    Cross-workbook snapshots land in the returned dict under their source
    book's `book` field so multi-book plans can restore each book independently.
    """
    from app.execution.range_utils import resolve_range, parse_address

    entries: list[dict[str, Any]] = []
    for addr in addresses:
        # We snapshot against the requested book when the address is workbook-
        # qualified; else against `book`. This matters for multi-book plans.
        parsed = parse_address(addr)
        target_book = book
        if parsed.workbook:
            # resolve_range will error if the book isn't open — let that propagate.
            rng = resolve_range(book, addr, default_sheet=default_sheet)
            target_book = rng.sheet.book
        else:
            rng = resolve_range(book, addr, default_sheet=default_sheet)

        # xlwings returns scalar for single-cell ranges; wrap into 2D for uniform
        # restore semantics.
        values = rng.value
        formulas = rng.formula
        number_format = rng.number_format

        if rng.count == 1:
            values = [[values]]
            formulas = [[formulas]]
            number_format = [[number_format]]
        else:
            if not isinstance(values, list) or (values and not isinstance(values[0], list)):
                # Single-row/column range comes back as a flat list — reshape.
                values = _to_2d(values, rng.shape)
                formulas = _to_2d(formulas, rng.shape)
                number_format = _to_2d(number_format, rng.shape)

        entries.append(
            {
                "book": target_book.name,
                "sheet": rng.sheet.name,
                "address": rng.address,
                "values": values,
                "formulas": formulas,
                "number_format": number_format,
            }
        )

    return {
        "timestamp": datetime.utcnow().isoformat(),
        "ranges": entries,
    }


def restore_snapshot(
    app_books: "xw.main.Books",
    snapshot: dict[str, Any],
) -> int:
    """
    Replay a snapshot back into the workbook(s) it came from. Returns the
    number of ranges successfully restored.

    If a range's origin book is no longer open, that entry is skipped and
    the restore count only reflects successful writes.
    """
    restored = 0
    for entry in snapshot.get("ranges", []):
        book_name = entry["book"]
        sheet_name = entry["sheet"]
        address = entry["address"]
        try:
            book = app_books[book_name]
        except Exception:
            continue  # book closed since snapshot — best-effort skip
        try:
            sheet = book.sheets[sheet_name]
            rng = sheet.range(address)
        except Exception:
            continue
        # Restore order matters: number_format before values so formats
        # survive a subsequent value write.
        try:
            rng.number_format = entry["number_format"]
        except Exception:
            pass
        # Use formula if non-empty; else fall back to values. This preserves
        # live formulas (e.g. =SUMIFS(...)) that the step had displaced.
        formulas = entry.get("formulas") or []
        if _any_nonempty(formulas):
            rng.formula = formulas
        else:
            rng.value = entry["values"]
        restored += 1
    return restored


def _to_2d(flat: Any, shape: tuple[int, int]) -> list[list[Any]]:
    """Coerce xlwings' sometimes-flat returns into a 2D matrix of known shape."""
    if flat is None:
        return [[None]]
    if not isinstance(flat, list):
        return [[flat]]
    if flat and isinstance(flat[0], list):
        return flat  # already 2D
    rows, cols = shape
    if rows == 1:
        return [list(flat)]
    if cols == 1:
        return [[v] for v in flat]
    # Fall back to single-row if shape disagrees — callers will see raw values.
    return [list(flat)]


def _any_nonempty(matrix: list[list[Any]]) -> bool:
    for row in matrix or []:
        for cell in row or []:
            if cell not in (None, "", 0, False):
                return True
    return False
