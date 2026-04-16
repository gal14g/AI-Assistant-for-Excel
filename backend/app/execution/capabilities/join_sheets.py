"""joinSheets — SQL-style join between two ranges on key columns.

Port of `frontend/src/engine/capabilities/joinSheets.ts`. Supports inner /
left / right / full outer joins. Both ranges are expected to include a header
row; the right range's key column is omitted from the output so the join key
appears only once (mirroring the TS semantics).
"""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


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


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    left_range = params.get("leftRange")
    right_range = params.get("rightRange")
    left_key_column = params.get("leftKeyColumn")
    right_key_column = params.get("rightKeyColumn")
    join_type = (params.get("joinType") or "inner").lower()
    output_range = params.get("outputRange")

    if not left_range or not right_range or left_key_column is None or right_key_column is None or not output_range:
        return {
            "status": "error",
            "message": (
                "joinSheets requires 'leftRange', 'rightRange', 'leftKeyColumn', "
                "'rightKeyColumn' and 'outputRange'."
            ),
        }

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": (
                f"Would {join_type} join {left_range} with {right_range} → {output_range}."
            ),
        }

    try:
        left_rng = resolve_range(ctx.workbook_handle, left_range)
        right_rng = resolve_range(ctx.workbook_handle, right_range)

        left_data = _as_2d(left_rng.value, left_rng.shape)
        right_data = _as_2d(right_rng.value, right_rng.shape)

        if len(left_data) < 1 or len(right_data) < 1:
            return {"status": "success", "message": "One or both ranges are empty.", "outputs": {}}

        l_key_idx = int(left_key_column) - 1
        r_key_idx = int(right_key_column) - 1

        left_headers = list(left_data[0])
        right_headers = list(right_data[0])
        right_headers_filtered = [h for i, h in enumerate(right_headers) if i != r_key_idx]
        combined_headers = [*left_headers, *right_headers_filtered]

        # Build right-side index — each key → list of row indices so multiple
        # matches produce multiple output rows (matches TS behavior).
        right_index: dict[str, list[int]] = {}
        for r in range(1, len(right_data)):
            row = right_data[r]
            key_raw = row[r_key_idx] if 0 <= r_key_idx < len(row) else None
            key = str(key_raw if key_raw is not None else "").lower()
            right_index.setdefault(key, []).append(r)

        null_right = [None] * len(right_headers_filtered)
        null_left = [None] * len(left_headers)

        result_rows: list[list[Any]] = [combined_headers]
        matched_right_keys: set[str] = set()

        for lr in range(1, len(left_data)):
            left_row = left_data[lr]
            key_raw = left_row[l_key_idx] if 0 <= l_key_idx < len(left_row) else None
            key = str(key_raw if key_raw is not None else "").lower()
            right_matches = right_index.get(key)

            if right_matches:
                matched_right_keys.add(key)
                for rr in right_matches:
                    right_row = right_data[rr]
                    right_vals = [v for i, v in enumerate(right_row) if i != r_key_idx]
                    result_rows.append([*left_row, *right_vals])
            elif join_type in ("left", "full"):
                result_rows.append([*left_row, *null_right])
            # inner: skip unmatched left rows

        if join_type in ("right", "full"):
            for rr in range(1, len(right_data)):
                right_row = right_data[rr]
                key_raw = right_row[r_key_idx] if 0 <= r_key_idx < len(right_row) else None
                key = str(key_raw if key_raw is not None else "").lower()
                if key not in matched_right_keys:
                    right_vals = [v for i, v in enumerate(right_row) if i != r_key_idx]
                    result_rows.append([*null_left, *right_vals])

        # Normalize row widths so xlwings accepts a rectangular 2D array.
        width = len(combined_headers)
        for r in result_rows:
            while len(r) < width:
                r.append(None)

        out_rng = resolve_range(ctx.workbook_handle, output_range)
        out_rng.resize(len(result_rows), width).value = result_rows
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"joinSheets failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": (
            f"Joined {len(left_data) - 1} left rows with {len(right_data) - 1} right rows → "
            f"{len(result_rows) - 1} result rows ({join_type})."
        ),
        "outputs": {"outputRange": output_range},
    }


registry.register("joinSheets", handler, mutates=True)
