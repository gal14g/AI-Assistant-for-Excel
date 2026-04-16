"""normalize_string — aggressive canonicalization for equality and comparison.

Python mirror of frontend/src/engine/utils/normalizeString.ts. Same behavior
and test coverage — see backend/tests/test_parse_utilities.py.

Covers the frequent reasons two strings that LOOK identical don't compare equal:
  - Leading/trailing whitespace
  - Internal whitespace collapsed (double spaces, NBSP, narrow-NBSP, tabs)
  - Unicode NFC normalization (composed vs decomposed)
  - Bidirectional control marks stripped (LRM, RLM, LRE/RLE/PDF, isolation marks)
  - Zero-width characters stripped (ZWJ, ZWNJ, BOM, ZWSP)
"""

from __future__ import annotations

import re
import unicodedata
from typing import Any

# Bidi control + zero-width characters.
_INVISIBLE_RE = re.compile(
    r"[\u200B-\u200F\u202A-\u202E\u2066-\u2069\uFEFF]"
)

# Whitespace variants: regular, tab, newline, NBSP, narrow NBSP, thin, hair, ideographic.
_WS_RE = re.compile(r"[\s\u00A0\u202F\u2009\u200A\u3000]+")


def normalize_string(value: Any) -> str:
    if value is None:
        return ""
    s = value if isinstance(value, str) else str(value)
    nfc = unicodedata.normalize("NFC", s)
    visible = _INVISIBLE_RE.sub("", nfc)
    return _WS_RE.sub(" ", visible).strip()


def normalize_for_compare(value: Any) -> str:
    """Case-insensitive canonicalization. Hebrew has no case so it's a no-op there."""
    return normalize_string(value).lower()
