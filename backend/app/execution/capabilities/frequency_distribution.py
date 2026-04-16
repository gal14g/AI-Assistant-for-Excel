"""frequencyDistribution — count occurrences of each unique value.

Flattens the source range, counts each unique value (case-insensitive for
strings), then writes a ``Value / Count [/ Percent]`` table to the output
range. Supports sorting by value or frequency, ascending or descending.
"""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def _natural_key(s: str) -> list[Any]:
    """Numeric-aware sort key mirroring JS localeCompare({numeric: true})."""
    import re

    parts = re.split(r"(\d+)", s)
    key: list[Any] = []
    for part in parts:
        if part.isdigit():
            key.append((0, int(part)))
        else:
            key.append((1, part.lower()))
    return key


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    source = params.get("sourceRange")
    output = params.get("outputRange")
    sort_by = (params.get("sortBy") or "frequency").lower()
    ascending = bool(params.get("ascending", False))
    include_percent = bool(params.get("includePercent", True))

    if not source or not output:
        return {
            "status": "error",
            "message": "frequencyDistribution requires 'sourceRange' and 'outputRange'.",
        }

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would compute frequency distribution of {source} → {output}.",
        }

    try:
        src_rng = resolve_range(ctx.workbook_handle, source)
        raw = src_rng.value

        # Normalize to 2D.
        if raw is None:
            vals: list[list[Any]] = []
        elif not isinstance(raw, list):
            vals = [[raw]]
        elif raw and not isinstance(raw[0], list):
            rows_, cols_ = src_rng.shape
            vals = [list(raw)] if rows_ == 1 else [[v] for v in raw]
        else:
            vals = [list(r) for r in raw]

        if not vals:
            return {"status": "success", "message": "No data found.", "outputs": {}}

        # Flatten non-null, non-empty cells.
        flat: list[Any] = []
        for row in vals:
            for cell in row:
                if cell is None or cell == "":
                    continue
                flat.append(cell)

        # Count — case-insensitive key for strings.
        counts: dict[str, dict[str, Any]] = {}
        for val in flat:
            key = val.lower() if isinstance(val, str) else str(val)
            if key not in counts:
                counts[key] = {"display": val, "count": 0}
            counts[key]["count"] += 1

        entries = list(counts.values())

        if sort_by == "frequency":
            entries.sort(key=lambda e: e["count"], reverse=not ascending)
        else:  # sort by value — numeric-aware
            entries.sort(
                key=lambda e: _natural_key(str(e["display"])),
                reverse=not ascending,
            )

        total_count = len(flat)
        header: list[Any] = (
            ["Value", "Count", "Percent"] if include_percent else ["Value", "Count"]
        )
        data_rows: list[list[Any]] = []
        for e in entries:
            row: list[Any] = [e["display"], e["count"]]
            if include_percent:
                pct = round((e["count"] / total_count) * 10000) / 100 if total_count else 0
                row.append(pct)
            data_rows.append(row)

        output_rows = [header] + data_rows
        rows = len(output_rows)
        cols = len(header)

        out_rng = resolve_range(ctx.workbook_handle, output)
        out_rng.resize(rows, cols).value = output_rows
    except Exception as exc:  # noqa: BLE001
        return {
            "status": "error",
            "message": f"frequencyDistribution failed: {exc}",
            "error": str(exc),
        }

    return {
        "status": "success",
        "message": f"Found {len(entries)} unique values across {total_count} total entries.",
        "outputs": {"outputRange": output},
    }


registry.register("frequencyDistribution", handler, mutates=True)
