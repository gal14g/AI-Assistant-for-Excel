"""consolidateRanges — merge multiple ranges into one consolidated range.

Port of `frontend/src/engine/capabilities/consolidateRanges.ts`. Supports
vertical stacking (default) and horizontal (side-by-side) joining, with
optional source-label column and deduplication.
"""

from __future__ import annotations

import json
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
    source_ranges = params.get("sourceRanges") or []
    output_range = params.get("outputRange")
    direction = (params.get("direction") or "vertical").lower()
    add_source_label = bool(params.get("addSourceLabel", False))
    deduplicate = bool(params.get("deduplicate", False))

    if not source_ranges or not output_range:
        return {
            "status": "error",
            "message": "consolidateRanges requires 'sourceRanges' and 'outputRange'.",
        }

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would consolidate {len(source_ranges)} ranges to {output_range}.",
        }

    try:
        tables: list[list[list[Any]]] = []
        for addr in source_ranges:
            rng = resolve_range(ctx.workbook_handle, addr)
            tables.append(_as_2d(rng.value, rng.shape))

        if direction == "vertical":
            combined: list[list[Any]] = []
            for i, table in enumerate(tables):
                if not table:
                    continue
                # Include header only from the first source (matches TS).
                start = 0 if i == 0 else 1
                for r in range(start, len(table)):
                    row = [source_ranges[i], *table[r]] if add_source_label else list(table[r])
                    combined.append(row)

            if deduplicate:
                seen: set[str] = set()
                deduped: list[list[Any]] = []
                for row in combined:
                    key = json.dumps(row, default=str)
                    if key in seen:
                        continue
                    seen.add(key)
                    deduped.append(row)
                combined = deduped

            if not combined:
                return {"status": "success", "message": "No data to consolidate.", "outputs": {}}

            width = max(len(r) for r in combined)
            for row in combined:
                while len(row) < width:
                    row.append(None)

            out_rng = resolve_range(ctx.workbook_handle, output_range)
            out_rng.resize(len(combined), width).value = combined

            dedup_suffix = " (deduped)" if deduplicate else ""
            return {
                "status": "success",
                "message": (
                    f"Consolidated {len(tables)} ranges → {len(combined)} rows "
                    f"in {output_range}{dedup_suffix}."
                ),
                "outputs": {"outputRange": output_range},
            }

        # horizontal: join side-by-side
        max_rows = max((len(t) for t in tables), default=0)
        if max_rows == 0:
            return {"status": "success", "message": "No data to consolidate.", "outputs": {}}

        combined_h: list[list[Any]] = [[] for _ in range(max_rows)]
        for table in tables:
            # Compute canonical width so missing rows still take space.
            width = max((len(r) for r in table), default=0) if table else 0
            for r in range(max_rows):
                if r < len(table):
                    row = list(table[r])
                    while len(row) < width:
                        row.append(None)
                    combined_h[r].extend(row)
                else:
                    combined_h[r].extend([None] * width)

        final_width = max((len(r) for r in combined_h), default=0)
        for row in combined_h:
            while len(row) < final_width:
                row.append(None)

        out_rng = resolve_range(ctx.workbook_handle, output_range)
        out_rng.resize(len(combined_h), final_width).value = combined_h

        return {
            "status": "success",
            "message": (
                f"Horizontally joined {len(tables)} ranges → "
                f"{len(combined_h)} rows × {final_width} columns."
            ),
            "outputs": {"outputRange": output_range},
        }
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"consolidateRanges failed: {exc}", "error": str(exc)}


registry.register("consolidateRanges", handler, mutates=True)
