"""extractMatchedToNewRow — within-row-match extraction.

When row[keyColumnIndexA] == row[keyColumnIndexB], lift the
`extractColumnIndexes` values into a new row below, duplicating the key
value in column A's position.
"""

from __future__ import annotations

from typing import Any, Optional

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range
from app.execution.utils import normalize_string, normalize_for_compare


def _ensure_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if not isinstance(values[0], list):
        return [values] if len(values) > 1 else [[values[0]]]
    return values


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    source_range = params.get("sourceRange")
    key_a = params.get("keyColumnIndexA")
    key_b = params.get("keyColumnIndexB")
    extract_idxs = params.get("extractColumnIndexes") or []
    has_headers = bool(params.get("hasHeaders", True))
    case_sensitive = bool(params.get("caseSensitive", False))

    if not source_range or key_a is None or key_b is None or not extract_idxs:
        return {
            "status": "error",
            "message": "extractMatchedToNewRow requires sourceRange + keyColumnIndexA/B + extractColumnIndexes.",
        }
    key_a = int(key_a)
    key_b = int(key_b)
    extract_idxs = [int(i) for i in extract_idxs]

    src = resolve_range(ctx.workbook_handle, source_range)
    used = src.current_region if src.count > 1 else src
    raw = _ensure_2d(used.value)
    if not raw:
        return {"status": "success", "message": "Source range is empty."}

    orig_cols = len(raw[0])
    for idx in [key_a, key_b, *extract_idxs]:
        if idx < 0 or idx >= orig_cols:
            return {"status": "error", "message": f"Column index {idx} out of range (src has {orig_cols} cols)."}

    header_row: Optional[list[Any]] = raw[0] if has_headers else None
    data_rows = raw[1:] if has_headers else raw[:]

    def eq(a: Any, b: Any) -> bool:
        if a in (None, "") or b in (None, ""):
            return False
        # Normalize string comparisons (trim, NFC, strip bidi/zero-width,
        # collapse whitespace) so Hebrew with RTL marks or NBSP still matches.
        if isinstance(a, str) or isinstance(b, str):
            if case_sensitive:
                return normalize_string(a) == normalize_string(b)
            return normalize_for_compare(a) == normalize_for_compare(b)
        return a == b

    new_data: list[list[Any]] = []
    matched_count = 0
    extract_set = set(extract_idxs)

    for row in data_rows:
        row = list(row) + [None] * (orig_cols - len(row))
        a_val, b_val = row[key_a], row[key_b]
        if not eq(a_val, b_val):
            new_data.append(row)
            continue
        matched_count += 1
        trimmed = [None if c in extract_set else row[c] for c in range(orig_cols)]
        new_data.append(trimmed)

        new_row: list[Any] = []
        for c in range(orig_cols):
            if c == key_a:
                new_row.append(a_val)
            elif c in extract_set:
                new_row.append(row[c])
            else:
                new_row.append(None)
        new_data.append(new_row)

    if matched_count == 0:
        return {
            "status": "success",
            "message": f"No matches — columns {key_a}/{key_b} never agreed on any row.",
            "outputs": {"outputRange": used.address, "matchedRowCount": 0},
        }

    output: list[list[Any]] = []
    if header_row:
        output.append(list(header_row) + [None] * (orig_cols - len(header_row)))
    output.extend(new_data)

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would extract {matched_count} matched row(s) in {used.address}.",
        }

    sheet = used.sheet
    top_left = used[0, 0]
    try:
        target = sheet.range(
            (top_left.row, top_left.column),
            (top_left.row + len(output) - 1, top_left.column + orig_cols - 1),
        )
        target.value = output
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Write failed: {exc}", "error": str(exc)}

    output_addr = sheet.range(
        (top_left.row, top_left.column),
        (top_left.row + len(output) - 1, top_left.column + orig_cols - 1),
    ).address
    return {
        "status": "success",
        "message": f"Extracted {matched_count} matched row(s) into new rows. Output: {output_addr}.",
        "outputs": {"outputRange": output_addr, "matchedRowCount": matched_count},
    }


registry.register("extractMatchedToNewRow", handler, mutates=True)
