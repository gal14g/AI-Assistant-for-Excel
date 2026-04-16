"""lateralSpreadDuplicates — duplicate-sidecar layout.

Port of frontend/src/engine/capabilities/lateralSpreadDuplicates.ts. Reads
the source range values, groups by key column, builds a widened output grid
with each first-occurrence row's duplicates laid out horizontally on the
left or right, and writes back in place.
"""

from __future__ import annotations

from typing import Any, Optional

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range
from app.execution.utils import normalize_for_compare


def _ensure_2d(values: Any) -> list[list[Any]]:
    """Normalize xlwings `.value` return types to a 2D list."""
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if not isinstance(values[0], list):
        # Single row or single column — wrap.
        return [values] if len(values) > 1 else [[values[0]]]
    return values


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    source_range = params.get("sourceRange")
    key_col = params.get("keyColumnIndex")
    has_headers = bool(params.get("hasHeaders", True))
    direction = params.get("direction", "left")
    remove_orig = bool(params.get("removeOriginalDuplicates", True))

    if not source_range or key_col is None:
        return {
            "status": "error",
            "message": "lateralSpreadDuplicates requires 'sourceRange' and 'keyColumnIndex'.",
        }
    key_col = int(key_col)

    src = resolve_range(ctx.workbook_handle, source_range)
    used = src.current_region if src.count > 1 else src
    # xlwings used_range covers the whole sheet; current_region is closer to
    # Office.js getUsedRange(false) for a given anchor.

    raw = _ensure_2d(used.value)
    if not raw:
        return {"status": "success", "message": "Source range is empty."}

    orig_cols = len(raw[0])
    if key_col < 0 or key_col >= orig_cols:
        return {
            "status": "error",
            "message": f"keyColumnIndex {key_col} out of range (source has {orig_cols} columns).",
        }

    header_row: Optional[list[Any]] = raw[0] if has_headers else None
    data_rows = raw[1:] if has_headers else raw[:]

    # Pass 1: first-occurrence + duplicates map.
    first_idx: dict[str, int] = {}
    dup_idxs: dict[str, list[int]] = {}
    for i, row in enumerate(data_rows):
        raw_key = row[key_col] if len(row) > key_col else None
        if raw_key is None or raw_key == "":
            continue
        # Normalize for bucketing so Hebrew with RTL marks, NBSP padding,
        # and case variations group together even if visually distinct.
        key = normalize_for_compare(raw_key)
        if not key:
            continue
        if key not in first_idx:
            first_idx[key] = i
        else:
            dup_idxs.setdefault(key, []).append(i)

    max_dups = max((len(v) for v in dup_idxs.values()), default=0)
    if max_dups == 0:
        return {
            "status": "success",
            "message": f"No duplicates found in column {key_col} of {used.address}.",
            "outputs": {"outputRange": used.address, "duplicateGroupCount": 0, "duplicateRowCount": 0},
        }

    sidecar_cols = max_dups * orig_cols
    new_width = sidecar_cols + orig_cols

    # Pass 2: surviving rows (non-duplicate or first-occurrence).
    duplicate_set = {idx for arr in dup_idxs.values() for idx in arr}
    survivors: list[int] = []
    for i in range(len(data_rows)):
        if remove_orig and i in duplicate_set:
            continue
        survivors.append(i)

    def empty_row() -> list[Any]:
        return [None] * new_width

    def build_anchor_row(src_row: list[Any], dup_rows: list[list[Any]]) -> list[Any]:
        out = empty_row()
        for b, dup in enumerate(dup_rows):
            if direction == "left":
                block_start = (max_dups - 1 - b) * orig_cols
            else:
                block_start = orig_cols + b * orig_cols
            for c in range(orig_cols):
                out[block_start + c] = dup[c] if c < len(dup) else None
        anchor_start = sidecar_cols if direction == "left" else 0
        for c in range(orig_cols):
            out[anchor_start + c] = src_row[c] if c < len(src_row) else None
        return out

    output_grid: list[list[Any]] = []
    if has_headers and header_row:
        widened = empty_row()
        def place_block(col: int) -> None:
            for c in range(orig_cols):
                widened[col + c] = header_row[c] if c < len(header_row) else None
        if direction == "left":
            for b in range(max_dups):
                place_block(b * orig_cols)
            place_block(sidecar_cols)
        else:
            place_block(0)
            for b in range(max_dups):
                place_block(orig_cols + b * orig_cols)
        output_grid.append(widened)

    for i in survivors:
        anchor = data_rows[i]
        key_raw = anchor[key_col] if len(anchor) > key_col else None
        key = normalize_for_compare(key_raw)
        if first_idx.get(key) == i and dup_idxs.get(key):
            dup_rows_data = [data_rows[idx] for idx in dup_idxs[key]]
            output_grid.append(build_anchor_row(anchor, dup_rows_data))
        else:
            output_grid.append(build_anchor_row(anchor, []))

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would lateral-spread duplicates of column {key_col} in {used.address}.",
        }

    sheet = used.sheet
    top_left = used[0, 0]  # first cell of the used range

    try:
        # xlwings: assign a 2D list to a single cell to write a full block.
        target = sheet.range(
            (top_left.row, top_left.column),
            (top_left.row + len(output_grid) - 1, top_left.column + new_width - 1),
        )
        target.value = output_grid

        # If we shrunk the row count, blank the tail of the original source
        # region so stale data doesn't remain.
        orig_rows = used.rows.count
        if len(output_grid) < orig_rows:
            tail = sheet.range(
                (top_left.row + len(output_grid), top_left.column),
                (top_left.row + orig_rows - 1, top_left.column + orig_cols - 1),
            )
            tail.value = [[None] * orig_cols for _ in range(orig_rows - len(output_grid))]
    except Exception as exc:  # noqa: BLE001
        return {
            "status": "error",
            "message": f"Failed to write lateral-spread layout: {exc}",
            "error": str(exc),
        }

    group_count = sum(1 for arr in dup_idxs.values() if arr)
    dup_row_count = len(duplicate_set)
    output_addr = sheet.range(
        (top_left.row, top_left.column),
        (top_left.row + len(output_grid) - 1, top_left.column + new_width - 1),
    ).address
    return {
        "status": "success",
        "message": (
            f"Lateral-spread {dup_row_count} duplicate row(s) across {group_count} key(s) "
            f"into {max_dups} sidecar block(s) on the {direction} of the anchor. Output: {output_addr}."
        ),
        "outputs": {
            "outputRange": output_addr,
            "duplicateGroupCount": group_count,
            "duplicateRowCount": dup_row_count,
        },
    }


registry.register("lateralSpreadDuplicates", handler, mutates=True)
