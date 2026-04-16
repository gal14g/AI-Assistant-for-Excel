"""fuzzyMatch — fuzzy string matching between two columns using similarity scoring.

Port of `frontend/src/engine/capabilities/fuzzyMatch.ts`. The TS side uses a
Levenshtein-based ratio (1 - distance/maxLen). We mirror that exact semantic
here with a pure-Python implementation so scores match cell-for-cell.
"""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def _levenshtein(a: str, b: str) -> int:
    m, n = len(a), len(b)
    if m == 0:
        return n
    if n == 0:
        return m
    # Single-row DP to keep memory O(min(m, n)).
    prev = list(range(n + 1))
    for i in range(1, m + 1):
        curr = [i] + [0] * n
        for j in range(1, n + 1):
            if a[i - 1] == b[j - 1]:
                curr[j] = prev[j - 1]
            else:
                curr[j] = 1 + min(prev[j], curr[j - 1], prev[j - 1])
        prev = curr
    return prev[n]


def _similarity(a: str, b: str) -> float:
    if a == b:
        return 1.0
    max_len = max(len(a), len(b))
    if max_len == 0:
        return 1.0
    return 1.0 - _levenshtein(a, b) / max_len


def _as_2d(raw: Any, shape: tuple[int, int]) -> list[list[Any]]:
    """Normalize xlwings read to 2D, matching the frontend's Excel.Range.values."""
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
    lookup_range = params.get("lookupRange")
    source_range = params.get("sourceRange")
    output_range = params.get("outputRange")
    threshold = float(params.get("threshold", 0.7) or 0.7)
    write_value = params.get("writeValue")
    # Preserve TS default semantic (returnBestMatch default = False, but both
    # branches actually produce the same output — writing the source row's
    # first column — so this flag has no observable effect there or here).
    return_best_match = bool(params.get("returnBestMatch", False))

    if not lookup_range or not source_range or not output_range:
        return {
            "status": "error",
            "message": "fuzzyMatch requires 'lookupRange', 'sourceRange' and 'outputRange'.",
        }

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": (
                f"Would fuzzy match {lookup_range} against {source_range} "
                f"(threshold {round(threshold * 100)}%), output to {output_range}."
            ),
        }

    try:
        lookup_rng = resolve_range(ctx.workbook_handle, lookup_range)
        source_rng = resolve_range(ctx.workbook_handle, source_range)

        lookup_vals = _as_2d(lookup_rng.value, lookup_rng.shape)
        source_vals = _as_2d(source_rng.value, source_rng.shape)

        if not lookup_vals or not source_vals:
            return {"status": "success", "message": "No data to match.", "outputs": {}}

        # Normalize source first-column strings once.
        source_strings: list[str] = [
            str(row[0] if row else "").strip().lower() if row else ""
            for row in source_vals
        ]

        results: list[list[Any]] = []
        match_count = 0

        for row in lookup_vals:
            lookup_str = str(row[0] if row else "").strip().lower() if row else ""
            if not lookup_str:
                results.append([None])
                continue

            best_score = 0.0
            best_idx = -1
            for j, src in enumerate(source_strings):
                if not src:
                    continue
                score = _similarity(lookup_str, src)
                if score > best_score:
                    best_score = score
                    best_idx = j

            if best_score >= threshold and best_idx >= 0:
                if write_value is not None:
                    out_val: Any = write_value
                else:
                    # Both returnBestMatch branches produce the same value
                    # in the TS reference — the first column of the best row.
                    _ = return_best_match  # intentionally unused; kept for param parity
                    src_row = source_vals[best_idx]
                    out_val = str(src_row[0]) if src_row and src_row[0] is not None else ""
                results.append([out_val])
                match_count += 1
            else:
                results.append([None])

        out_rng = resolve_range(ctx.workbook_handle, output_range)
        out_rng.resize(len(results), 1).value = results
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"fuzzyMatch failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": (
            f"Fuzzy matched {match_count}/{len(lookup_vals)} rows "
            f"(threshold {round(threshold * 100)}%)."
        ),
        "outputs": {"outputRange": output_range},
    }


registry.register("fuzzyMatch", handler, mutates=True)
