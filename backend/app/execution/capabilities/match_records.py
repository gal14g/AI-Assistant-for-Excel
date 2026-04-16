"""matchRecords — lookup/match records between two ranges.

Port of `frontend/src/engine/capabilities/matchRecords.ts`. Two execution paths:

    preferFormula=True (default) → write XLOOKUP formulas (desktop Excel 2019+
                                   on Windows always supports XLOOKUP, so no
                                   probing/fallback is needed here — unlike
                                   the web-add-in where Excel 2016 is in play).
    preferFormula=False          → computed JS-side value match, written as
                                   static cell values.

Multi-column composite keys, writeValue, and matchType="contains" all route
through `_composite_key_match`, which mirrors the TS deterministic matcher.
"""

from __future__ import annotations

import re
from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


_COL_RE = re.compile(r"[A-Z]+", re.IGNORECASE)


# ── Address helpers ─────────────────────────────────────────────────────────


def _strip_workbook_qualifier(addr: str) -> str:
    """`[Book.xlsx]Sheet!A:A` → `Sheet!A:A`."""
    if addr.startswith("[[") and addr.endswith("]]"):
        addr = addr[2:-2].strip()
    if addr.startswith("["):
        end = addr.find("]")
        if end != -1:
            return addr[end + 1 :]
    return addr


def _split_sheet(addr: str) -> tuple[str, str]:
    """('Sheet!A1' | 'A1') → (sheet_prefix_with_bang, ref)."""
    if "!" in addr:
        sheet, ref = addr.split("!", 1)
        return f"{sheet}!", ref
    return "", addr


def _col_letter_to_index(col: str) -> int:
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _offset_column(col: str, offset: int) -> str:
    num = _col_letter_to_index(col) + offset
    if num <= 0:
        return "A"
    result = ""
    while num > 0:
        rem = (num - 1) % 26
        result = chr(65 + rem) + result
        num = (num - 1) // 26
    return result or "A"


def _count_cols_in_addr(addr: str) -> int:
    _, ref = _split_sheet(addr)
    cols = _COL_RE.findall(ref)
    if not cols:
        return 1
    if len(cols) < 2:
        return 1
    return _col_letter_to_index(cols[1]) - _col_letter_to_index(cols[0]) + 1


def _get_rel_cell_ref(range_address: str, row_offset: int) -> str:
    """
    "Sheet1!A:A", offset 0 → "Sheet1!A1"
    "Sheet1!A2:A100", offset 2 → "Sheet1!A4"
    """
    prefix, ref = _split_sheet(range_address)
    stripped = ref.replace("$", "")
    col_m = _COL_RE.search(stripped)
    col = col_m.group(0).upper() if col_m else "A"
    row_m = re.search(r"\d+", stripped)
    start_row = int(row_m.group(0)) if row_m else 1
    return f"{prefix}{col}{start_row + row_offset}"


def _get_column_ref(range_address: str, col_offset: int) -> str:
    """ "Sheet1!A:A", offset 1 → "Sheet1!B:B" """
    prefix, ref = _split_sheet(range_address)
    start_col_m = _COL_RE.search(ref)
    start_col = start_col_m.group(0).upper() if start_col_m else "A"
    col = _offset_column(start_col, col_offset)
    return f"{prefix}{col}:{col}"


def _quote_sheet_in_ref(ref: str) -> str:
    """Wrap non-ASCII sheet names in single quotes for formula strings."""
    if "!" not in ref:
        return ref
    sheet, cell = ref.split("!", 1)
    if not sheet or sheet.startswith("'"):
        return ref
    # Quote if any non-ASCII or whitespace characters present.
    if any(not c.isascii() or c.isspace() for c in sheet):
        return f"'{sheet}'!{cell}"
    return ref


def _build_output_range(output_addr: str, row_count: int, col_count: int) -> str:
    prefix, ref = _split_sheet(output_addr)
    stripped = ref.replace("$", "")
    col_m = _COL_RE.search(stripped)
    start_col = col_m.group(0).upper() if col_m else "A"
    row_m = re.search(r"\d+", stripped)
    start_row = int(row_m.group(0)) if row_m else 1
    end_col = _offset_column(start_col, col_count - 1)
    end_row = start_row + row_count - 1
    return f"{prefix}{start_col}{start_row}:{end_col}{end_row}"


# ── Normalisation helpers ───────────────────────────────────────────────────


def _as_2d(raw: Any, shape: tuple[int, int]) -> list[list[Any]]:
    if raw is None:
        return []
    rows_, cols_ = shape
    if not isinstance(raw, list):
        return [[raw]]
    if raw and not isinstance(raw[0], list):
        if rows_ == 1:
            return [list(raw)]
        return [[v] for v in raw]
    return [list(r) for r in raw]


def _normalize(v: Any) -> str:
    return str(v if v is not None else "").strip().lower()


def _is_empty_row(row: list[Any]) -> bool:
    return all(v is None or v == "" for v in row)


def _fill_down(vals: list[list[Any]]) -> tuple[list[list[Any]], list[bool]]:
    """Forward-fill merged-cell blanks. Rows with no prior non-empty stay empty."""
    filled: list[list[Any]] = []
    was_empty: list[bool] = []
    last: list[Any] | None = None
    for row in vals:
        if _is_empty_row(row):
            was_empty.append(True)
            filled.append(list(last) if last is not None else row)
        else:
            last = row
            was_empty.append(False)
            filled.append(row)
    return filled, was_empty


# ── Formula builders ────────────────────────────────────────────────────────


def _build_xlookup_formulas(
    lookup_addr: str,
    source_addr: str,
    return_columns: list[int],
    match_mode: str,
    row_count: int,
) -> list[list[str]]:
    formulas: list[list[str]] = []
    for row in range(row_count):
        row_formulas: list[str] = []
        for col_idx in return_columns:
            lookup_cell = _quote_sheet_in_ref(_get_rel_cell_ref(lookup_addr, row))
            lookup_arr = _quote_sheet_in_ref(_get_column_ref(source_addr, 0))
            return_arr = _quote_sheet_in_ref(_get_column_ref(source_addr, col_idx - 1))
            row_formulas.append(
                f'=IFERROR(XLOOKUP({lookup_cell},{lookup_arr},{return_arr},"",{match_mode}),"")'
            )
        formulas.append(row_formulas)
    return formulas


# ── Formula-based path ──────────────────────────────────────────────────────


def _formula_match(
    ctx: ExecutorContext,
    params: dict[str, Any],
    return_columns: list[int],
) -> dict[str, Any]:
    lookup_addr = _strip_workbook_qualifier(params["lookupRange"])
    source_addr = _strip_workbook_qualifier(params["sourceRange"])
    output_addr = _strip_workbook_qualifier(params["outputRange"])
    match_mode = "0" if (params.get("matchType") or "exact") == "exact" else "1"

    # Determine actual data row count via current_region / used_range — avoids
    # the 1M-row blow-up that naïve "A:A" handling triggers in xlwings too.
    lookup_rng = resolve_range(ctx.workbook_handle, params["lookupRange"])
    try:
        used = lookup_rng.current_region
        row_count = used.rows.count
    except Exception:  # noqa: BLE001
        shape = lookup_rng.shape
        row_count = shape[0] if shape else 0

    if row_count == 0:
        return {"status": "success", "message": "No rows to process.", "outputs": {}}

    precise_output = _build_output_range(output_addr, row_count, len(return_columns))
    output_rng = resolve_range(ctx.workbook_handle, precise_output)
    output_rng.formula = _build_xlookup_formulas(
        lookup_addr, source_addr, return_columns, match_mode, row_count
    )

    return {
        "status": "success",
        "message": f"Created {row_count} XLOOKUP formulas in {output_addr}.",
        "outputs": {"outputRange": params["outputRange"]},
    }


# ── Value-based path ────────────────────────────────────────────────────────


def _computed_match(
    ctx: ExecutorContext,
    params: dict[str, Any],
    return_columns: list[int],
) -> dict[str, Any]:
    lookup_rng = resolve_range(ctx.workbook_handle, params["lookupRange"])
    source_rng = resolve_range(ctx.workbook_handle, params["sourceRange"])

    lookup_values = _as_2d(lookup_rng.value, lookup_rng.shape)
    source_values = _as_2d(source_rng.value, source_rng.shape)

    # First-seen wins — matches the TS `if (!index.has(key))` guard.
    index: dict[str, list[Any]] = {}
    for row in source_values:
        if not row:
            continue
        key = _normalize(row[0])
        if key not in index:
            index[key] = row

    results: list[list[Any]] = []
    for lookup_row in lookup_values:
        key = _normalize(lookup_row[0]) if lookup_row else ""
        source_row = index.get(key)
        out_row: list[Any] = []
        for col_idx in return_columns:
            if source_row and 0 <= col_idx - 1 < len(source_row):
                val = source_row[col_idx - 1]
                out_row.append(val if val is not None else None)
            else:
                out_row.append(None)
        results.append(out_row)

    # Resolve the precise output range using the lookup range's actual start row.
    start_row = lookup_rng.row
    output_addr = _strip_workbook_qualifier(params["outputRange"])
    prefix, ref = _split_sheet(output_addr)
    col_m = _COL_RE.search(ref.replace("$", ""))
    out_col = col_m.group(0).upper() if col_m else "A"
    end_col = _offset_column(out_col, len(return_columns) - 1)
    precise_addr = f"{prefix}{out_col}{start_row}:{end_col}{start_row + len(results) - 1}"

    output_rng = resolve_range(ctx.workbook_handle, precise_addr)
    output_rng.value = results

    matched = sum(1 for r in results if any(v is not None for v in r))
    return {
        "status": "success",
        "message": f"Matched {matched}/{len(lookup_values)} records, wrote to {precise_addr}.",
        "outputs": {"outputRange": params["outputRange"]},
    }


# ── Composite key path ──────────────────────────────────────────────────────


def _composite_key_match(
    ctx: ExecutorContext,
    params: dict[str, Any],
    return_columns: list[int],
) -> dict[str, Any]:
    write_value = params.get("writeValue") or "match"
    match_type = params.get("matchType")
    is_contains = match_type in ("contains", "approximate")

    lookup_rng = resolve_range(ctx.workbook_handle, params["lookupRange"])
    source_rng = resolve_range(ctx.workbook_handle, params["sourceRange"])

    source_vals = _as_2d(source_rng.value, source_rng.shape)
    lookup_vals = _as_2d(lookup_rng.value, lookup_rng.shape)

    lookup_start_row = lookup_rng.row

    filled_source, _ = _fill_down(source_vals)
    filled_lookup, lookup_was_empty = _fill_down(lookup_vals)

    def to_key(row: list[Any]) -> str:
        return "\x00".join(_normalize(v) for v in row)

    source_keys: set[str] = set()
    source_list: list[str] = []
    for row in filled_source:
        if not _is_empty_row(row):
            k = to_key(row)
            source_keys.add(k)
            if is_contains:
                source_list.append(k)

    if is_contains:
        def matches(key: str) -> bool:
            return any(key in s or s in key for s in source_list)
    else:
        def matches(key: str) -> bool:
            return key in source_keys

    # Resolve the output sheet + column from the output address.
    output_addr = _strip_workbook_qualifier(params["outputRange"])
    prefix, ref = _split_sheet(output_addr)
    # strip quotes from sheet name for workbook lookup
    sheet_name = prefix[:-1].strip("'") if prefix else None
    col_m = _COL_RE.search(ref.replace("$", ""))
    out_col = col_m.group(0).upper() if col_m else "G"

    book = ctx.workbook_handle
    out_ws = book.sheets[sheet_name] if sheet_name else book.sheets.active

    match_count = 0
    skipped_count = 0
    for i, row in enumerate(filled_lookup):
        if _is_empty_row(row) and lookup_was_empty[i]:
            continue
        if matches(to_key(row)):
            sheet_row = lookup_start_row + i
            try:
                out_ws.range(f"{out_col}{sheet_row}").value = write_value
                match_count += 1
            except Exception:  # noqa: BLE001
                skipped_count += 1

    tail = f" ({skipped_count} skipped — merged/protected)" if skipped_count > 0 else ""
    return {
        "status": "success",
        "message": (
            f"Composite match: {match_count}/{len(filled_lookup)} rows matched — "
            f"wrote {write_value!r} to {sheet_name or 'active sheet'} column {out_col}{tail}."
        ),
        "outputs": {"outputRange": params["outputRange"]},
    }


# ── Entry point ─────────────────────────────────────────────────────────────


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    lookup_range = params.get("lookupRange")
    source_range = params.get("sourceRange")
    output_range = params.get("outputRange")
    if not lookup_range or not source_range or not output_range:
        return {
            "status": "error",
            "message": "matchRecords requires 'lookupRange', 'sourceRange' and 'outputRange'.",
        }

    prefer_formula = bool(params.get("preferFormula", True))
    return_columns = list(params.get("returnColumns") or []) or [1]

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would match records from {lookup_range} against {source_range} → {output_range}.",
        }

    try:
        lookup_cols = _count_cols_in_addr(_strip_workbook_qualifier(lookup_range))
        source_cols = _count_cols_in_addr(_strip_workbook_qualifier(source_range))
        route_composite = (
            lookup_cols > 1
            or source_cols > 1
            or params.get("writeValue") is not None
            or params.get("matchType") == "contains"
        )

        if route_composite:
            return _composite_key_match(ctx, params, return_columns)
        if prefer_formula:
            return _formula_match(ctx, params, return_columns)
        return _computed_match(ctx, params, return_columns)
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"matchRecords failed: {exc}", "error": str(exc)}


registry.register("matchRecords", handler, mutates=True)
