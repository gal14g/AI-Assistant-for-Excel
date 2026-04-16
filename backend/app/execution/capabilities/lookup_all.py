"""lookupAll — return ALL matching rows between two ranges joined by a delimiter.

Port of `frontend/src/engine/capabilities/lookupAll.ts`. For each lookup value,
finds every matching source row (by first column) and collects values from the
specified 1-based return column. Results are joined with a delimiter and
written to the output range.
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
    lookup_range = params.get("lookupRange")
    source_range = params.get("sourceRange")
    return_column = params.get("returnColumn")
    output_range = params.get("outputRange")
    delimiter = params.get("delimiter", ", ")
    match_type = (params.get("matchType") or "exact").lower()

    if not lookup_range or not source_range or return_column is None or not output_range:
        return {
            "status": "error",
            "message": (
                "lookupAll requires 'lookupRange', 'sourceRange', 'returnColumn', "
                "and 'outputRange'."
            ),
        }

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": (
                f"Would look up all matches from {lookup_range} in {source_range}, "
                f"return column {return_column} to {output_range}."
            ),
        }

    try:
        lookup_rng = resolve_range(ctx.workbook_handle, lookup_range)
        source_rng = resolve_range(ctx.workbook_handle, source_range)

        lookup_vals = _as_2d(lookup_rng.value, lookup_rng.shape)
        source_vals = _as_2d(source_rng.value, source_rng.shape)

        if not lookup_vals or not source_vals:
            return {"status": "success", "message": "No data to look up.", "outputs": {}}

        ret_col_idx = int(return_column) - 1

        # Build an index: source key (first column) → list of return-column values.
        index: dict[str, list[str]] = {}
        for row in source_vals:
            if not row:
                continue
            key = str(row[0] if row[0] is not None else "").strip().lower()
            if not key:
                continue
            ret_val_raw = row[ret_col_idx] if 0 <= ret_col_idx < len(row) else None
            ret_val = str(ret_val_raw if ret_val_raw is not None else "")
            index.setdefault(key, []).append(ret_val)

        results: list[list[Any]] = []
        matched_lookups = 0
        total_matches = 0

        for row in lookup_vals:
            lookup_str = str(row[0] if row and row[0] is not None else "").strip().lower()
            if not lookup_str:
                results.append([None])
                continue

            matched: list[str] = []
            if match_type == "exact":
                matched = list(index.get(lookup_str, []))
            else:
                # "contains": either direction — lookup contains key OR key contains lookup.
                for key, vals in index.items():
                    if lookup_str in key or key in lookup_str:
                        matched.extend(vals)

            if matched:
                results.append([delimiter.join(matched)])
                matched_lookups += 1
                total_matches += len(matched)
            else:
                results.append([None])

        out_rng = resolve_range(ctx.workbook_handle, output_range)
        out_rng.resize(len(results), 1).value = results
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"lookupAll failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": (
            f"Found matches for {matched_lookups}/{len(lookup_vals)} lookup values "
            f"({total_matches} total matches)."
        ),
        "outputs": {"outputRange": output_range},
    }


registry.register("lookupAll", handler, mutates=True)
