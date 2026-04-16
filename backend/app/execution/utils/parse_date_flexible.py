"""parse_date_flexible — tolerant date parser.

Python mirror of frontend/src/engine/utils/parseDateFlexible.ts. Same
behavior and same test coverage — see backend/tests/test_parse_utilities.py.

Accepts every shape realistic user data produces:
  - ISO 8601: 2026-03-15, 2026/03/15, 2026-03-15T14:30:00
  - dd/mm/yyyy (default)
  - mm/dd/yyyy (when locale_hint='mdy')
  - dd-mm-yyyy
  - Excel serial date (int/float 1..2_958_465)
  - Python datetime / date
  - Month-name forms: "15 Mar 2026", "March 15, 2026"
  - 2-digit years with cutoff (< 50 → 20xx, >= 50 → 19xx)

Returns a `date` on success, `None` on failure.
"""

from __future__ import annotations

import re
from datetime import date, datetime, timedelta
from typing import Any, Literal, Optional

DateLocaleHint = Literal["dmy", "mdy", "auto"]

_MONTH_NAMES: dict[str, int] = {
    "jan": 1, "january": 1,
    "feb": 2, "february": 2,
    "mar": 3, "march": 3,
    "apr": 4, "april": 4,
    "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7,
    "aug": 8, "august": 8,
    "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12,
}

# Excel's epoch quirk: serial 0 = 1899-12-30 in most implementations.
_EXCEL_EPOCH = date(1899, 12, 30)
_EXCEL_MAX_SERIAL = 2_958_465  # 9999-12-31


def _normalize_year(y: int) -> int:
    if y >= 100:
        return y
    return 2000 + y if y < 50 else 1900 + y


def _make_date_if_valid(y: int, m: int, d: int) -> Optional[date]:
    if m < 1 or m > 12 or d < 1 or d > 31:
        return None
    try:
        return date(y, m, d)
    except ValueError:
        return None


def _excel_serial_to_date(serial: float) -> Optional[date]:
    if serial < 1 or serial > _EXCEL_MAX_SERIAL:
        return None
    try:
        return _EXCEL_EPOCH + timedelta(days=int(serial))
    except (OverflowError, ValueError):
        return None


_ISO_RE = re.compile(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[T ].*)?$")
_NUMERIC_DATE_RE = re.compile(r"^(\d{1,4})[-/.](\d{1,2})[-/.](\d{1,4})$")
_DAY_FIRST_MONTHNAME_RE = re.compile(r"^(\d{1,2})[-/ ]+([A-Za-z]+)[-/ ,]+(\d{2,4})$")
_MONTH_FIRST_MONTHNAME_RE = re.compile(r"^([A-Za-z]+)[-/ ]+(\d{1,2})[-/ ,]+(\d{2,4})$")


def parse_date_flexible(
    value: Any,
    locale_hint: DateLocaleHint = "dmy",
) -> Optional[date]:
    # datetime/date pass-through.
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    # Excel serial number.
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return _excel_serial_to_date(float(value))

    if not isinstance(value, str):
        return None

    s = value.strip()
    if not s:
        return None

    # ── ISO: yyyy-mm-dd (unambiguous) ──
    m = _ISO_RE.match(s)
    if m:
        return _make_date_if_valid(int(m.group(1)), int(m.group(2)), int(m.group(3)))

    # ── Month-name forms ──
    m = _DAY_FIRST_MONTHNAME_RE.match(s)
    if m:
        mi = _MONTH_NAMES.get(m.group(2).lower())
        if mi is not None:
            return _make_date_if_valid(_normalize_year(int(m.group(3))), mi, int(m.group(1)))

    m = _MONTH_FIRST_MONTHNAME_RE.match(s)
    if m:
        mi = _MONTH_NAMES.get(m.group(1).lower())
        if mi is not None:
            return _make_date_if_valid(_normalize_year(int(m.group(3))), mi, int(m.group(2)))

    # ── Numeric d-m-y or m-d-y ──
    m = _NUMERIC_DATE_RE.match(s)
    if m:
        a, b, c = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if a >= 1000:
            # Year-leading form we somehow didn't catch above (e.g. dotted).
            return _make_date_if_valid(a, b, c)
        if a > 12 and b <= 12:
            day, month, year = a, b, c
        elif b > 12 and a <= 12:
            day, month, year = b, a, c
        else:
            if locale_hint == "mdy":
                month, day, year = a, b, c
            else:
                day, month, year = a, b, c
        return _make_date_if_valid(_normalize_year(year), month, day)

    return None


def format_date_dmy(d: date) -> str:
    return d.strftime("%d/%m/%Y")


def format_date_mdy(d: date) -> str:
    return d.strftime("%m/%d/%Y")
