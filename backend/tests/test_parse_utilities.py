"""Tests for the shared parse/normalize utilities.

Mirrors frontend/src/__tests__/parseDateFlexible.test.ts +
parseNumberFlexible.test.ts + normalizeString.test.ts so the two sides
stay behaviorally aligned.
"""

from __future__ import annotations

from datetime import date, datetime

import pytest

from app.execution.utils import (
    normalize_for_compare,
    normalize_string,
    parse_date_flexible,
    parse_number_flexible,
)
from app.execution.utils.parse_date_flexible import format_date_dmy, format_date_mdy


# ══════════════════════════════════════════════════════════════════════════
# parse_date_flexible
# ══════════════════════════════════════════════════════════════════════════


class TestParseDateISO:
    def test_yyyy_mm_dd(self):
        assert parse_date_flexible("2026-03-15") == date(2026, 3, 15)

    def test_yyyy_slash_mm_slash_dd(self):
        assert parse_date_flexible("2026/03/15") == date(2026, 3, 15)

    def test_iso_with_time(self):
        assert parse_date_flexible("2026-03-15T14:30:00") == date(2026, 3, 15)

    def test_iso_single_digit_components(self):
        assert parse_date_flexible("2026-3-5") == date(2026, 3, 5)


class TestParseDateDMY:
    def test_slash_sep(self):
        assert parse_date_flexible("15/03/2026") == date(2026, 3, 15)

    def test_dash_sep(self):
        assert parse_date_flexible("15-03-2026") == date(2026, 3, 15)

    def test_dot_sep(self):
        assert parse_date_flexible("15.03.2026") == date(2026, 3, 15)

    def test_day_over_12_unambiguous(self):
        # Even with mdy hint, 25 is clearly the day.
        assert parse_date_flexible("25/03/2026", "mdy") == date(2026, 3, 25)

    def test_month_over_12_flips(self):
        assert parse_date_flexible("03/25/2026", "dmy") == date(2026, 3, 25)


class TestParseDateMDY:
    def test_with_hint(self):
        assert parse_date_flexible("03/15/2026", "mdy") == date(2026, 3, 15)

    def test_ambiguous_mdy(self):
        # "05/04/2026" with mdy hint: month=05, day=04 → May 4.
        assert parse_date_flexible("05/04/2026", "mdy") == date(2026, 5, 4)

    def test_ambiguous_dmy_default(self):
        # With dmy default: day=05, month=04 → April 5.
        assert parse_date_flexible("05/04/2026") == date(2026, 4, 5)


class TestParseDateTwoDigitYears:
    def test_26_becomes_2026(self):
        assert parse_date_flexible("15/03/26").year == 2026

    def test_98_becomes_1998(self):
        assert parse_date_flexible("15/03/98").year == 1998

    def test_49_becomes_2049(self):
        assert parse_date_flexible("15/03/49").year == 2049

    def test_50_becomes_1950(self):
        assert parse_date_flexible("15/03/50").year == 1950


class TestParseDateMonthName:
    def test_day_first_short(self):
        assert parse_date_flexible("15 Mar 2026") == date(2026, 3, 15)

    def test_day_first_full(self):
        assert parse_date_flexible("15 March 2026") == date(2026, 3, 15)

    def test_month_first_comma(self):
        assert parse_date_flexible("March 15, 2026") == date(2026, 3, 15)

    def test_month_first_no_comma(self):
        assert parse_date_flexible("Mar 15 2026") == date(2026, 3, 15)

    def test_month_name_with_two_digit_year(self):
        assert parse_date_flexible("15 Mar 26") == date(2026, 3, 15)


class TestParseDateExcelSerial:
    def test_known_ref(self):
        # 2025-03-15 = Excel serial 45731
        assert parse_date_flexible(45731) == date(2025, 3, 15)

    def test_reject_negative(self):
        assert parse_date_flexible(-1) is None

    def test_reject_out_of_range(self):
        assert parse_date_flexible(10_000_000) is None


class TestParseDateObjects:
    def test_datetime_to_date(self):
        assert parse_date_flexible(datetime(2026, 3, 15, 14, 30)) == date(2026, 3, 15)

    def test_date_passthrough(self):
        d = date(2026, 3, 15)
        assert parse_date_flexible(d) == d


class TestParseDateInvalid:
    def test_empty(self):
        assert parse_date_flexible("") is None

    def test_whitespace_only(self):
        assert parse_date_flexible("   ") is None

    def test_gibberish(self):
        assert parse_date_flexible("hello world") is None

    def test_feb_31(self):
        assert parse_date_flexible("31/02/2026") is None

    def test_zero_date(self):
        assert parse_date_flexible("00/00/2026") is None

    def test_none(self):
        assert parse_date_flexible(None) is None

    def test_list(self):
        assert parse_date_flexible([]) is None


class TestFormatters:
    def test_dmy(self):
        assert format_date_dmy(date(2026, 3, 15)) == "15/03/2026"

    def test_mdy(self):
        assert format_date_mdy(date(2026, 3, 15)) == "03/15/2026"


# ══════════════════════════════════════════════════════════════════════════
# parse_number_flexible
# ══════════════════════════════════════════════════════════════════════════


class TestParseNumberNative:
    def test_int(self):
        assert parse_number_flexible(1234) == 1234.0

    def test_float(self):
        assert parse_number_flexible(1234.56) == 1234.56

    def test_negative(self):
        assert parse_number_flexible(-1234) == -1234.0

    def test_zero(self):
        assert parse_number_flexible(0) == 0.0

    def test_reject_nan(self):
        assert parse_number_flexible(float("nan")) is None

    def test_reject_inf(self):
        assert parse_number_flexible(float("inf")) is None

    def test_bool_true(self):
        assert parse_number_flexible(True) == 1.0

    def test_bool_false(self):
        assert parse_number_flexible(False) == 0.0


class TestParseNumberPlain:
    def test_int_string(self):
        assert parse_number_flexible("1234") == 1234.0

    def test_decimal_string(self):
        assert parse_number_flexible("1234.56") == 1234.56

    def test_leading_zeros(self):
        assert parse_number_flexible("0042") == 42.0

    def test_positive_prefix(self):
        assert parse_number_flexible("+100") == 100.0

    def test_negative_prefix(self):
        assert parse_number_flexible("-100") == -100.0


class TestParseNumberUSFormat:
    def test_thousand_sep(self):
        assert parse_number_flexible("1,234") == 1234.0

    def test_thousand_and_decimal(self):
        assert parse_number_flexible("1,234.56") == 1234.56

    def test_multi_thousand(self):
        assert parse_number_flexible("1,234,567.89") == 1234567.89

    def test_indian_lakh(self):
        assert parse_number_flexible("1,23,456.78") == 123456.78


class TestParseNumberEUFormat:
    def test_unambiguous_eu(self):
        assert parse_number_flexible("1.234,56") == 1234.56

    def test_multi_thousand_eu(self):
        assert parse_number_flexible("1.234.567,89") == 1234567.89

    def test_comma_decimal_with_eu_hint(self):
        assert parse_number_flexible("1,5", "eu") == 1.5

    def test_dot_thousand_with_eu_hint(self):
        assert parse_number_flexible("1.234", "eu") == 1234.0


class TestParseNumberAutoDetect:
    def test_auto_thousand_dot(self):
        assert parse_number_flexible("1.234", "auto") == 1234.0

    def test_auto_decimal_dot(self):
        assert parse_number_flexible("1.23", "auto") == 1.23

    def test_auto_thousand_comma(self):
        assert parse_number_flexible("1,234", "auto") == 1234.0

    def test_auto_decimal_comma(self):
        assert parse_number_flexible("1,5", "auto") == 1.5


class TestParseNumberCurrency:
    def test_usd_prefix(self):
        assert parse_number_flexible("$1,234.56") == 1234.56

    def test_eur_prefix(self):
        assert parse_number_flexible("€1.234,56") == 1234.56

    def test_ils_prefix(self):
        assert parse_number_flexible("₪1,234") == 1234.0

    def test_gbp_prefix(self):
        assert parse_number_flexible("£100") == 100.0

    def test_iso_code_suffix_usd(self):
        assert parse_number_flexible("100 USD") == 100.0

    def test_iso_code_suffix_ils(self):
        assert parse_number_flexible("100 ILS") == 100.0

    def test_symbol_with_space(self):
        assert parse_number_flexible("₪ 1,234") == 1234.0


class TestParseNumberPercent:
    def test_fifty_percent(self):
        assert parse_number_flexible("50%") == 0.5

    def test_hundred_percent(self):
        assert parse_number_flexible("100%") == 1.0

    def test_small_percent(self):
        assert parse_number_flexible("0.5%") == 0.005

    def test_negative_percent(self):
        assert parse_number_flexible("-25%") == -0.25


class TestParseNumberNegatives:
    def test_paren_negative(self):
        assert parse_number_flexible("(100)") == -100.0

    def test_paren_with_seps(self):
        assert parse_number_flexible("(1,234.56)") == -1234.56

    def test_paren_with_currency(self):
        assert parse_number_flexible("($1,234)") == -1234.0

    def test_trailing_minus(self):
        assert parse_number_flexible("100-") == -100.0


class TestParseNumberScientific:
    def test_positive_exponent(self):
        assert parse_number_flexible("1.23E+06") == 1230000.0

    def test_negative_exponent(self):
        assert parse_number_flexible("1.23e-03") == pytest.approx(0.00123)

    def test_no_sign_exponent(self):
        assert parse_number_flexible("5E4") == 50000.0


class TestParseNumberWhitespace:
    def test_nbsp_thousand(self):
        assert parse_number_flexible("1\u00A0234") == 1234.0

    def test_narrow_nbsp_thousand(self):
        assert parse_number_flexible("1\u202F234.56") == 1234.56

    def test_padded(self):
        assert parse_number_flexible("  42  ") == 42.0


class TestParseNumberInvalid:
    def test_empty(self):
        assert parse_number_flexible("") is None

    def test_letters(self):
        assert parse_number_flexible("abc") is None

    def test_unknown_currency_word(self):
        assert parse_number_flexible("100 dollars") is None

    def test_multiple_dots_invalid(self):
        assert parse_number_flexible("1.2.3") is None

    def test_none(self):
        assert parse_number_flexible(None) is None


class TestParseNumberCombined:
    def test_paren_currency_seps(self):
        assert parse_number_flexible("($1,234.56)") == -1234.56

    def test_percent_negative(self):
        assert parse_number_flexible("-50%") == -0.5

    def test_eu_currency(self):
        assert parse_number_flexible("€1.234,56") == 1234.56

    def test_eu_percent(self):
        assert parse_number_flexible("50,5%") == pytest.approx(0.505)


# ══════════════════════════════════════════════════════════════════════════
# normalize_string
# ══════════════════════════════════════════════════════════════════════════


class TestNormalizeStringWhitespace:
    def test_trim(self):
        assert normalize_string("  hello  ") == "hello"

    def test_collapse_double(self):
        assert normalize_string("hello    world") == "hello world"

    def test_collapse_tabs(self):
        assert normalize_string("hello\t\nworld") == "hello world"

    def test_nbsp(self):
        assert normalize_string("hello\u00A0world") == "hello world"

    def test_narrow_nbsp(self):
        assert normalize_string("hello\u202Fworld") == "hello world"

    def test_thin_space(self):
        assert normalize_string("hello\u2009world") == "hello world"

    def test_ideographic_space(self):
        assert normalize_string("hello\u3000world") == "hello world"


class TestNormalizeStringBidi:
    def test_strip_lrm(self):
        assert normalize_string("abc\u200Edef") == "abcdef"

    def test_strip_rlm(self):
        assert normalize_string("abc\u200Fdef") == "abcdef"

    def test_strip_zwj(self):
        assert normalize_string("abc\u200Ddef") == "abcdef"

    def test_strip_zwnj(self):
        assert normalize_string("abc\u200Cdef") == "abcdef"

    def test_strip_zwsp(self):
        assert normalize_string("abc\u200Bdef") == "abcdef"

    def test_strip_bom(self):
        assert normalize_string("\uFEFFhello") == "hello"

    def test_strip_bidi_embedding(self):
        assert normalize_string("abc\u202Bdef\u202Cghi") == "abcdefghi"

    def test_strip_bidi_isolation(self):
        assert normalize_string("abc\u2066def\u2069ghi") == "abcdefghi"


class TestNormalizeStringHebrew:
    def test_plain_hebrew(self):
        assert normalize_string("שלום עולם") == "שלום עולם"

    def test_strip_trailing_rlm(self):
        assert normalize_string("דוד\u200F") == "דוד"

    def test_trim_trailing_spaces(self):
        assert normalize_string("דוד   ") == "דוד"

    def test_mixed_ascii(self):
        assert normalize_string("דוד 123") == "דוד 123"


class TestNormalizeStringNFC:
    def test_e_combining_acute(self):
        decomposed = "e\u0301"  # e + combining acute
        composed = "\u00E9"  # é
        assert normalize_string(decomposed) == composed

    def test_composed_equals_decomposed(self):
        decomposed = "e\u0301"
        composed = "\u00E9"
        assert normalize_string(decomposed) == normalize_string(composed)


class TestNormalizeStringEdges:
    def test_empty(self):
        assert normalize_string("") == ""

    def test_whitespace_only(self):
        assert normalize_string("   \t\n  ") == ""

    def test_none(self):
        assert normalize_string(None) == ""

    def test_int(self):
        assert normalize_string(123) == "123"

    def test_bool(self):
        assert normalize_string(True) == "True"

    def test_preserve_case(self):
        assert normalize_string("Hello World") == "Hello World"


class TestNormalizeForCompare:
    def test_lowercase(self):
        assert normalize_for_compare("Hello World") == "hello world"

    def test_hebrew_passthrough(self):
        assert normalize_for_compare("שלום") == "שלום"

    def test_whitespace_invisible_equivalence(self):
        a = "דוד\u200F"
        b = "  דוד  "
        assert normalize_for_compare(a) == normalize_for_compare(b)

    def test_nbsp_equivalence(self):
        assert normalize_for_compare("hello\u00A0world") == normalize_for_compare("hello world")
