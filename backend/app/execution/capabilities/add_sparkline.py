"""
addSparkline — attach sparkline mini-charts to a location range.

Uses desktop Excel's VBA/COM SparklineGroups API (always present on Excel
2010+), so unlike the TS handler we don't need an ExcelApi-version fallback.
Type: 1=line, 2=column, 3=winloss.
"""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


# Excel XlSparkType enum values.
_SPARK_TYPE = {
    "line": 1,
    "column": 2,
    "winloss": 3,
    "winLoss": 3,
}


def _hex_to_rgb_int(color: str) -> int:
    """#RRGGBB → Excel BGR long (R + G*256 + B*65536)."""
    h = color.lstrip("#")
    if len(h) != 6:
        raise ValueError(f"Invalid color {color!r}; expected #RRGGBB.")
    return int(h[0:2], 16) + int(h[2:4], 16) * 256 + int(h[4:6], 16) * 65536


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    data_range = params.get("dataRange")
    location_range = params.get("locationRange")
    spark_type = (params.get("sparklineType") or "line")
    color = params.get("color")

    if not data_range or not location_range:
        return {
            "status": "error",
            "message": "addSparkline requires 'dataRange' and 'locationRange'.",
        }

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would add {spark_type} sparklines to {location_range} from {data_range}.",
        }

    try:
        data_rng = resolve_range(ctx.workbook_handle, data_range)
        loc_rng = resolve_range(ctx.workbook_handle, location_range)

        type_val = _SPARK_TYPE.get(spark_type, 1)

        # The SparklineGroups collection lives on the *location* sheet's COM
        # Range. API: location.SparklineGroups.Add(Type, SourceData) where
        # SourceData is the external address string (sheet-qualified).
        source_address = data_rng.sheet.name + "!" + data_rng.address.lstrip("=")
        # data_rng.address already includes the sheet prefix in many xlwings
        # versions — strip any duplicate and rebuild as "SheetName!A1:B2".
        cell_only = data_rng.address.split("!")[-1]
        source_address = f"{data_rng.sheet.name}!{cell_only}"

        group = loc_rng.api.SparklineGroups.Add(type_val, source_address)

        if color:
            try:
                group.SeriesColor.Color = _hex_to_rgb_int(color)
            except Exception:  # noqa: BLE001 — best-effort styling
                pass
    except Exception as exc:  # noqa: BLE001
        return {
            "status": "error",
            "message": f"Add sparkline failed: {exc}",
            "error": str(exc),
        }

    return {
        "status": "success",
        "message": f"Added {spark_type} sparklines to {loc_rng.address} from {data_rng.address}.",
        "outputs": {"range": loc_rng.address},
    }


registry.register("addSparkline", handler, mutates=False, affects_formatting=True)
