"""parse_number_flexible — tolerant numeric parser.

Python mirror of frontend/src/engine/utils/parseNumberFlexible.ts. Same
behavior, same test coverage — see backend/tests/test_parse_utilities.py.

Accepts every realistic number-as-text representation:
  - Native numbers: 1234, 1234.56, bool
  - US with separators: 1,234.56
  - EU with separators: 1.234,56
  - Indian: 1,23,456.78
  - Currency prefix: $1,234, €1.234,56, ₪1,234
  - ISO code suffix: "100 USD", "100 ILS"
  - Percent: "50%" → 0.5
  - Parenthesis negative: "(100)" → -100
  - Trailing-minus: "100-" → -100
  - Scientific: "1.23E+06"

Returns None on anything that doesn't cleanly resolve.
"""

from __future__ import annotations

import re
from typing import Any, Literal, Optional

NumberLocaleHint = Literal["us", "eu", "auto"]

# Whitespace variants to strip (regular, NBSP, narrow NBSP, thin space).
_WS_RE = re.compile(r"[\s\u00A0\u202F\u2009]")

# Currency symbols we strip as a prefix/suffix.
_CURRENCY_SYMBOLS_RE = re.compile(r"[$€£¥₪₤₹¢₩]")

# ISO currency codes (word-boundary to avoid over-matching).
_CURRENCY_CODES_RE = re.compile(
    r"\b(USD|EUR|GBP|JPY|ILS|INR|KRW|CNY|CHF|CAD|AUD|NZD|SEK|NOK|DKK|"
    r"PLN|CZK|HUF|MXN|BRL|ZAR|TRY|RUB|HKD|SGD|THB|IDR|PHP)\b",
    re.IGNORECASE,
)

_SCIENTIFIC_RE = re.compile(r"^(.+?)[eE]([+-]?\d+)$")
_NUMERIC_ONLY_RE = re.compile(r"^[\d,.]+$")


def parse_number_flexible(
    value: Any,
    locale_hint: NumberLocaleHint = "auto",
) -> Optional[float]:
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, (int, float)):
        v = float(value)
        return v if _finite(v) else None
    if not isinstance(value, str):
        return None

    # 1. Strip currency symbols and ISO codes while word boundaries are intact.
    s = _CURRENCY_SYMBOLS_RE.sub("", value)
    s = _CURRENCY_CODES_RE.sub("", s)

    # 2. Strip all whitespace (including NBSP variants).
    s = _WS_RE.sub("", s)
    if not s:
        return None

    # 3. Trailing percent.
    percent_divisor = 1.0
    if s.endswith("%"):
        percent_divisor = 100.0
        s = s[:-1]

    # 4. Parenthesis-negative.
    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1]

    # 5. Leading + or -.
    if s.startswith("+"):
        s = s[1:]
    if s.startswith("-"):
        negative = not negative
        s = s[1:]

    # 6. Trailing-minus accounting style.
    if s.endswith("-"):
        negative = not negative
        s = s[:-1]

    if not s:
        return None

    # 7. Scientific notation — split off the exponent.
    exponent = 0
    sci_m = _SCIENTIFIC_RE.match(s)
    if sci_m:
        s = sci_m.group(1)
        exponent = int(sci_m.group(2))

    # 8. Mantissa must be digits + `,` / `.` only.
    if not _NUMERIC_ONLY_RE.match(s):
        return None

    dot_count = s.count(".")
    comma_count = s.count(",")

    if dot_count > 0 and comma_count > 0:
        # Rightmost separator is the decimal.
        if s.rfind(".") > s.rfind(","):
            normalized = s.replace(",", "")
        else:
            normalized = s.replace(".", "").replace(",", ".")
    elif dot_count == 1 and comma_count == 0:
        left, right = s.split(".")
        if locale_hint == "eu":
            normalized = s.replace(".", "") if len(right) == 3 and len(left) >= 1 else s
        elif locale_hint == "auto" and len(right) == 3 and 1 <= len(left) <= 3:
            normalized = s.replace(".", "")
        else:
            normalized = s
    elif comma_count == 1 and dot_count == 0:
        left, right = s.split(",")
        if locale_hint == "us":
            normalized = s.replace(",", "") if len(right) == 3 and len(left) >= 1 else s.replace(",", ".")
        elif locale_hint == "auto" and len(right) == 3 and 1 <= len(left) <= 3:
            normalized = s.replace(",", "")
        else:
            normalized = s.replace(",", ".")
    elif dot_count > 1 and comma_count == 0:
        parts = s.split(".")
        if not all(re.fullmatch(r"\d{3}", p) for p in parts[1:]):
            return None
        normalized = s.replace(".", "")
    elif comma_count > 1 and dot_count == 0:
        parts = s.split(",")
        if not all(re.fullmatch(r"\d{2,3}", p) for p in parts[1:]):
            return None
        normalized = s.replace(",", "")
    else:
        normalized = s

    try:
        n = float(normalized)
    except ValueError:
        return None
    if not _finite(n):
        return None

    result = (-n if negative else n) * (10.0 ** exponent) / percent_divisor
    return result


def _finite(x: float) -> bool:
    return x == x and x != float("inf") and x != float("-inf")
