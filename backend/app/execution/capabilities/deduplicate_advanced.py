"""deduplicateAdvanced — remove duplicate rows with a keep-which-row strategy.

Port of `frontend/src/engine/capabilities/deduplicateAdvanced.ts`. Supports the
same four strategies — first, last, mostComplete, newest — and mirrors the
Excel-serial date handling of the TS source so results align cell-for-cell.
"""

from __future__ import annotations

import math
from datetime import datetime
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


def _to_epoch_ms(val: Any) -> float:
    """Mirror the TS "Excel serial number → epoch ms" coercion."""
    if val is None or val == "":
        return -math.inf
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        # Excel serial (days since 1899-12-30) → epoch ms, matching TS:
        # (val - 25569) * 86400000
        return (float(val) - 25569) * 86_400_000
    if isinstance(val, datetime):
        return val.timestamp() * 1000
    try:
        # Fallback to ISO / natural parsing.
        return datetime.fromisoformat(str(val)).timestamp() * 1000
    except Exception:  # noqa: BLE001
        return -math.inf


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    key_columns = params.get("keyColumns") or []
    keep_strategy = (params.get("keepStrategy") or "first").lower()
    date_column = params.get("dateColumn")
    has_headers = bool(params.get("hasHeaders", True))

    if not address or not key_columns:
        return {
            "status": "error",
            "message": "deduplicateAdvanced requires 'range' and 'keyColumns'.",
        }

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": (
                f"Would deduplicate {address} on columns "
                f"[{', '.join(str(c) for c in key_columns)}] keeping {keep_strategy!r}."
            ),
        }

    try:
        rng = resolve_range(ctx.workbook_handle, address)
        all_vals = _as_2d(rng.value, rng.shape)
        if not all_vals:
            return {"status": "success", "message": "No data found.", "outputs": {}}

        header_row = all_vals[0] if has_headers else None
        data_rows = all_vals[1:] if has_headers else list(all_vals)

        # Composite key with null-byte separator (matches TS "\x00" join).
        groups: dict[str, list[int]] = {}
        for i, row in enumerate(data_rows):
            key_parts = []
            for col in key_columns:
                idx = int(col) - 1
                v = row[idx] if 0 <= idx < len(row) else None
                key_parts.append(str(v if v is not None else "").lower())
            key = "\x00".join(key_parts)
            groups.setdefault(key, []).append(i)

        keep_indices: set[int] = set()
        for indices in groups.values():
            if keep_strategy == "first":
                pick = indices[0]
            elif keep_strategy == "last":
                pick = indices[-1]
            elif keep_strategy == "mostcomplete":
                best_idx = indices[0]
                fewest_blanks = math.inf
                for idx in indices:
                    blanks = sum(1 for v in data_rows[idx] if v is None or v == "")
                    if blanks < fewest_blanks:
                        fewest_blanks = blanks
                        best_idx = idx
                pick = best_idx
            elif keep_strategy == "newest":
                dc = int(date_column if date_column is not None else 1) - 1
                best_idx = indices[0]
                latest = -math.inf
                for idx in indices:
                    val = data_rows[idx][dc] if 0 <= dc < len(data_rows[idx]) else None
                    t = _to_epoch_ms(val)
                    if t > latest:
                        latest = t
                        best_idx = idx
                pick = best_idx
            else:
                pick = indices[0]
            keep_indices.add(pick)

        kept_rows = [row for i, row in enumerate(data_rows) if i in keep_indices]
        output: list[list[Any]] = []
        if header_row is not None:
            output.append(list(header_row))
        output.extend([list(r) for r in kept_rows])

        # Pad with nulls so stale rows get overwritten (matches TS behavior).
        total_rows = len(all_vals)
        cols = len(all_vals[0])
        while len(output) < total_rows:
            output.append([None] * cols)

        # Normalize row widths so xlwings accepts a rectangular 2D array.
        for r in output:
            while len(r) < cols:
                r.append(None)

        rng.resize(len(output), cols).value = output
        removed = len(data_rows) - len(kept_rows)
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"deduplicateAdvanced failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": (
            f"Removed {removed} duplicate rows, kept {len(kept_rows)} unique rows "
            f"(strategy: {keep_strategy})."
        ),
        "outputs": {"range": address},
    }


registry.register("deduplicateAdvanced", handler, mutates=True)
