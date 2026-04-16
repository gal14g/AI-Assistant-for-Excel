"""normalizeDates — standardise date formats in a column.

Ported from `frontend/src/engine/capabilities/normalizeDates.ts`. Handles:
- Excel serial numbers (floats in a sensible range)
- yyyy-mm-dd / yyyy/mm/dd
- dd/mm/yyyy / mm/dd/yyyy with a heuristic disambiguation
- d-MMM-yy / d-MMM-yyyy
- datetime values coming directly from xlwings
- fallback to dateutil.parser when installed
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


_MONTH_ABBR = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}

_EXCEL_EPOCH_DAYS = 25569

_ISO_RE = re.compile(r"^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$")
_DMY_RE = re.compile(r"^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})$")
_MMM_RE = re.compile(r"^(\d{1,2})[/-]([A-Za-z]{3})[/-](\d{2,4})$")


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


def _try_parse_date(val: Any) -> datetime | None:
    if val is None or val == "":
        return None
    if isinstance(val, datetime):
        return val

    if isinstance(val, (int, float)) and not isinstance(val, bool):
        if 1 < val < 2958466:
            return datetime.utcfromtimestamp((val - _EXCEL_EPOCH_DAYS) * 86400)
        return None

    s = str(val).strip()

    m = _ISO_RE.match(s)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None

    m = _DMY_RE.match(s)
    if m:
        a, b, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            if a > 12:
                return datetime(y, b, a)
            if b > 12:
                return datetime(y, a, b)
            # Ambiguous → assume dd/mm/yyyy (TS uses same default)
            return datetime(y, b, a)
        except ValueError:
            return None

    m = _MMM_RE.match(s)
    if m:
        mon = _MONTH_ABBR.get(m.group(2).lower())
        if mon is not None:
            year = int(m.group(3))
            if year < 100:
                year += 2000 if year < 50 else 1900
            try:
                return datetime(year, mon, int(m.group(1)))
            except ValueError:
                return None

    if _dateutil_parser is not None:
        try:
            return _dateutil_parser.parse(s)
        except (ValueError, OverflowError, TypeError):
            return None

    return None


def _format_date(d: datetime, fmt: str) -> str:
    yyyy = f"{d.year:04d}"
    mm = f"{d.month:02d}"
    dd = f"{d.day:02d}"
    out = re.sub(r"yyyy", yyyy, fmt, flags=re.IGNORECASE)
    out = out.replace("mm", mm)
    out = out.replace("dd", dd)
    return out


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    output_format = params.get("outputFormat")

    if not address or not output_format:
        return {"status": "error", "message": "normalizeDates requires 'range' and 'outputFormat'."}

    rng = resolve_range(ctx.workbook_handle, address)

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would normalize dates in {rng.address} to '{output_format}'.",
        }

    vals = _to_2d(rng.value)
    if not vals:
        return {"status": "success", "message": "No data found.", "outputs": {"range": rng.address}}

    out: list[list[Any]] = [list(row) for row in vals]
    total = sum(len(r) for r in out)
    normalized = 0

    for r in range(len(out)):
        for c in range(len(out[r])):
            parsed = _try_parse_date(out[r][c])
            if parsed is not None:
                out[r][c] = _format_date(parsed, output_format)
                normalized += 1

    try:
        rng.value = out
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Write failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"Normalized {normalized}/{total} dates to '{output_format}' in {rng.address}.",
        "outputs": {"range": rng.address},
    }


registry.register("normalizeDates", handler, mutates=True)
