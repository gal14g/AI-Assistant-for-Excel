"""
Range / address resolution for the Python (xlwings) executor.

Mirrors the null-safe logic in `frontend/src/engine/capabilities/rangeUtils.ts`
so a plan authored for the add-in resolves identically when executed via MCP.

Supported address forms (same set the TS side accepts):

    A1                      single cell on the active sheet of the default book
    A1:C10                  range on the active sheet
    Sheet1!A1:C10           sheet-qualified
    'My Sheet'!A1:C10       quoted sheet name (spaces / non-ASCII)
    [Other.xlsx]Sheet!A1    workbook-qualified (cross-workbook!)
    [[Sheet!A1:C10]]        RangeToken notation — the taskpane's canonical form
    [[ [Book.xlsx]Sheet!A1:C10 ]]   workbook-qualified RangeToken

Cross-workbook resolution:
    Unlike Office.js (which is single-workbook), xlwings gives us app-level
    access to every open book. When an address starts with `[Book.xlsx]` we
    look up that book in the running Excel app and read/write against it.

Failure modes return `ValueError` with an actionable message — the MCP
tooling surfaces these as `isError` responses to the chat client.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import TYPE_CHECKING, Optional

if TYPE_CHECKING:  # xlwings is a heavy import; keep it lazy for type hints only
    import xlwings as xw


# ── Token notation ──────────────────────────────────────────────────────────
# The taskpane UI wraps user-click references in [[ ... ]] so the planner
# knows they are cell references rather than literal strings. Strip the
# wrapper before downstream processing.
_TOKEN_RE = re.compile(r"^\s*\[\[\s*(.+?)\s*\]\]\s*$")


def normalize_token(address: str) -> str:
    """Strip `[[ ... ]]` wrappers from a RangeToken if present."""
    m = _TOKEN_RE.match(address)
    return m.group(1) if m else address


# ── Workbook qualifier ──────────────────────────────────────────────────────
# `[Other.xlsx]Sheet!A1:B5` — the leading bracketed segment is the workbook.
# Sheet names may be single-quoted. Cell portion always starts with a column
# letter after the exclamation.
_WB_SHEET_RANGE_RE = re.compile(
    r"""^
        (?:\[(?P<book>[^\]]+)\])?          # optional [Book.xlsx]
        (?:
            '(?P<qsheet>[^']+)'            # quoted sheet name
            |
            (?P<sheet>[^!'\[]+?)           # bare sheet name
        )?
        (?:!(?P<cell>.+))?                 # !A1:B5 (optional when we have a sheet with no !)
        $""",
    re.VERBOSE,
)


@dataclass
class ResolvedAddress:
    """Fully-parsed address broken into its components."""

    workbook: Optional[str]
    """Book name (not path) if the address was `[Book.xlsx]...`; else None."""

    sheet: Optional[str]
    """Sheet name if explicit; else None (→ caller uses active sheet)."""

    cell: Optional[str]
    """Cell / range portion, e.g. `A1:C10`. May be None for whole-sheet ops."""

    raw: str
    """Original address string (for error messages)."""


def parse_address(address: str) -> ResolvedAddress:
    """
    Decompose an address string into (workbook, sheet, cell). Permissive:
    accepts every form listed in the module docstring. Doesn't *resolve*
    against a running Excel instance — that's `resolve_range()`'s job.

    Raises `ValueError` for clearly-malformed input.
    """
    if not address or not isinstance(address, str):
        raise ValueError(f"Address must be a non-empty string (got {address!r}).")

    raw = address
    inner = normalize_token(address).strip()

    # Special-case: "A1" / "A1:C10" / "A:A" / "1:1" with no sheet qualifier.
    # The regex-based split below handles these too, but this shortcut keeps
    # error messages clean for the common case.
    if "!" not in inner and "[" not in inner:
        return ResolvedAddress(workbook=None, sheet=None, cell=inner, raw=raw)

    m = _WB_SHEET_RANGE_RE.match(inner)
    if not m:
        raise ValueError(
            f"Could not parse address {raw!r}. Expected forms: "
            "A1, Sheet!A1, 'Sheet Name'!A1, [Book.xlsx]Sheet!A1."
        )

    return ResolvedAddress(
        workbook=(m.group("book") or None),
        sheet=(m.group("qsheet") or m.group("sheet") or None),
        cell=(m.group("cell") or None),
        raw=raw,
    )


# ── Live resolution against xlwings ─────────────────────────────────────────


def resolve_range(
    default_book: "xw.Book",
    address: str,
    *,
    default_sheet: Optional[str] = None,
) -> "xw.Range":
    """
    Turn an address string into a live `xlwings.Range`.

    `default_book` is used when the address doesn't include a workbook
    qualifier. `default_sheet` is used similarly when the address lacks a
    sheet name — callers typically pass `default_book.sheets.active.name`.

    Cross-workbook references walk `default_book.app.books` to find the
    target book by name (case-insensitive). If the book isn't open, raises
    `ValueError` — we deliberately don't silently fail over to opening the
    file from disk, since that would surprise users.
    """
    import xlwings as xw  # noqa: F401  — lazy import

    parsed = parse_address(address)

    # 1. Pick the workbook.
    if parsed.workbook:
        book = _find_open_book(default_book, parsed.workbook)
        if book is None:
            raise ValueError(
                f"Workbook {parsed.workbook!r} referenced in {parsed.raw!r} "
                "is not open. Open it in Excel and retry."
            )
    else:
        book = default_book

    # 2. Pick the sheet.
    sheet_name = parsed.sheet or default_sheet
    if sheet_name:
        try:
            sheet = book.sheets[sheet_name]
        except Exception as exc:
            raise ValueError(
                f"Sheet {sheet_name!r} not found in workbook {book.name!r} "
                f"(address: {parsed.raw!r})."
            ) from exc
    else:
        sheet = book.sheets.active

    # 3. Pick the cell/range. When `cell` is None we return the used range.
    if parsed.cell:
        try:
            return sheet.range(parsed.cell)
        except Exception as exc:
            raise ValueError(
                f"Range {parsed.cell!r} could not be resolved on sheet "
                f"{sheet.name!r} (address: {parsed.raw!r})."
            ) from exc
    return sheet.used_range


def _find_open_book(anchor: "xw.Book", name: str) -> Optional["xw.Book"]:
    """Case-insensitive match against every book open in the same Excel app."""
    target = name.lower()
    for book in anchor.app.books:
        if book.name.lower() == target:
            return book
        # Also accept a match against the file stem (without .xlsx) so the
        # planner can write "Budget" without the extension.
        stem = book.name.rsplit(".", 1)[0].lower()
        if stem == target.rsplit(".", 1)[0].lower():
            return book
    return None
