"""coerceDataType — convert values in a range to number/text/date.

Ported from `frontend/src/engine/capabilities/coerceDataType.ts`. Text-to-number
strips currency symbols and commas before parsing. Text-to-date uses
`dateutil.parser` when available (it's in requirements) and falls back to a
few common strptime formats.
"""

from __future__ import annotations

import re
from datetime import datetime
from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range

try:
    from dateutil import parser as _dateutil_parser  # type: ignore
except Exception:  # noqa: BLE001
    _dateutil_parser = None  # type: ignore


_NUM_STRIP = re.compile(r"[$€£¥,\s]")
_EXCEL_EPOCH_DAYS = 25569  # days from 1900-01-01 to 1970-01-01 (Excel serial)


def _to_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if not isinstance(values[0], list):
        return [values]
    return values


def _parse_date(val: Any) -> datetime | None:
    if val is None or val == "":
        return None
    if isinstance(val, datetime):
        return val
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        if 1 < val < 2958466:
            # Excel serial → unix seconds. 86400 s/day.
            return datetime.utcfromtimestamp((val - _EXCEL_EPOCH_DAYS) * 86400)
        return None
    s = str(val).strip()
    if _dateutil_parser is not None:
        try:
            return _dateutil_parser.parse(s, dayfirst=False, yearfirst=False)
        except (ValueError, OverflowError, TypeError):
            pass
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%dT%H:%M:%S"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def _format_date(d: datetime, fmt: str) -> str:
    """Translate Excel-style tokens (yyyy, mm, dd) into the formatted string."""
    yyyy = f"{d.year:04d}"
    mm = f"{d.month:02d}"
    dd = f"{d.day:02d}"
    # Replace yyyy (case-insensitive), then mm, then dd — same order as TS.
    out = re.sub(r"yyyy", yyyy, fmt, flags=re.IGNORECASE)
    out = out.replace("mm", mm)
    out = out.replace("dd", dd)
    return out


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    target_type = params.get("targetType")
    date_format = params.get("dateFormat") or "yyyy-mm-dd"

    if not address or target_type not in ("number", "text", "date"):
        return {
            "status": "error",
            "message": "coerceDataType requires 'range' and 'targetType' ∈ {number, text, date}.",
        }

    rng = resolve_range(ctx.workbook_handle, address)

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would convert values in {rng.address} to {target_type}.",
        }

    vals = _to_2d(rng.value)
    if not vals:
        return {"status": "success", "message": "No data found.", "outputs": {"range": rng.address}}

    out: list[list[Any]] = [list(row) for row in vals]
    total = sum(len(r) for r in out)
    converted = 0

    for r in range(len(out)):
        for c in range(len(out[r])):
            val = out[r][c]
            if val is None or val == "":
                continue
            if target_type == "number":
                cleaned = _NUM_STRIP.sub("", str(val))
                try:
                    num = float(cleaned)
                except ValueError:
                    continue
                # Preserve int-valued floats as ints (matches xlwings round-trip).
                out[r][c] = int(num) if num.is_integer() else num
                converted += 1
            elif target_type == "text":
                out[r][c] = str(val)
                converted += 1
            elif target_type == "date":
                parsed = _parse_date(val)
                if parsed is not None:
                    out[r][c] = _format_date(parsed, date_format)
                    converted += 1

    try:
        rng.value = out
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Write failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Converted {converted}/{total} cells to {target_type} in {rng.address}.",
        "outputs": {"range": rng.address},
    }


registry.register("coerceDataType", handler, mutates=True)
