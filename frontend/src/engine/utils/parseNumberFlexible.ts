/**
 * parseNumberFlexible — tolerant numeric parser.
 *
 * Accepts every realistic number-as-text representation encountered in
 * user data (CSV imports, copy-pastes from web pages, Hebrew/European
 * locale formats, accounting-style negatives, currency-prefixed cells):
 *
 *   - Native numbers                 1234,  1234.56
 *   - US format with separators      1,234.56
 *   - EU format with separators      1.234,56   (via localeHint='eu' or auto-detect)
 *   - Indian format                  1,23,456.78
 *   - Currency prefix                $1,234.56   €1.234,56   ₪1,234   £100
 *   - Currency suffix / ISO code     100 USD   100 ILS   100₪
 *   - Percent                        50%        → 0.5
 *   - Parenthesis negative           (1,234.56) → -1234.56
 *   - Trailing-minus                 100-       → -100
 *   - Scientific                     1.23E+06
 *   - Plus-prefixed                  +100
 *   - Whitespace-padded, with NBSP, narrow-NBSP, thin-space
 *
 * Returns `null` on any input that doesn't cleanly resolve to a number.
 *
 * Locale disambiguation: when a value has only ONE separator (just `,` or
 * just `.`) and the separator could be either thousand or decimal, we use
 * `localeHint`:
 *   - 'us'   → `.` = decimal, `,` = thousand
 *   - 'eu'   → `,` = decimal, `.` = thousand
 *   - 'auto' → prefer US unless the input has exactly 3 digits after the
 *             separator AND the separator is the only one, in which case
 *             treat as thousand (e.g. '1.234' is more likely 1234 than 1.234).
 */

export type NumberLocaleHint = "us" | "eu" | "auto";

// Characters we strip as whitespace (regular space + NBSP + narrow-NBSP + thin).
const WS_CHARS = "[\\s\\u00A0\\u202F\\u2009]";

// Currency symbols and ISO codes we strip. Keep the list conservative — the
// goal is high precision on common prefixes, not exhaustive coverage.
const CURRENCY_SYMBOLS = /[$€£¥₪₤₹¢₩]/g;
const CURRENCY_CODES = /\b(USD|EUR|GBP|JPY|ILS|INR|KRW|CNY|CHF|CAD|AUD|NZD|SEK|NOK|DKK|PLN|CZK|HUF|MXN|BRL|ZAR|TRY|RUB|HKD|SGD|THB|IDR|PHP)\b/gi;

export function parseNumberFlexible(
  value: unknown,
  localeHint: NumberLocaleHint = "auto",
): number | null {
  // Already numeric?
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }
  if (typeof value === "boolean") {
    return value ? 1 : 0;
  }
  if (typeof value !== "string") return null;

  // Strip currency symbols and ISO codes FIRST (while word boundaries are
  // still intact — "100 ILS" needs to match \bILS\b). After that we strip
  // all whitespace variants.
  let s = value.replace(CURRENCY_SYMBOLS, "").replace(CURRENCY_CODES, "");
  s = s.replace(new RegExp(WS_CHARS, "g"), "");
  if (!s) return null;

  // Percent: pop the trailing % and remember to divide.
  let percentDivisor = 1;
  if (s.endsWith("%")) {
    percentDivisor = 100;
    s = s.slice(0, -1);
  }

  // Parenthesis negative: "(100)" or "(1,234.56)" → negate.
  let negative = false;
  if (s.startsWith("(") && s.endsWith(")")) {
    negative = true;
    s = s.slice(1, -1);
  }

  // Leading + or - sign.
  if (s.startsWith("+")) s = s.slice(1);
  if (s.startsWith("-")) { negative = !negative; s = s.slice(1); }

  // Trailing-minus accounting style: "100-"
  if (s.endsWith("-")) { negative = !negative; s = s.slice(0, -1); }

  if (!s) return null;

  // Scientific notation: leave E exponent alone but keep a separator-aware
  // handling for the mantissa. We split on E first.
  const sciMatch = s.match(/^(.+?)[eE]([+-]?\d+)$/);
  let exponent = 0;
  if (sciMatch) {
    s = sciMatch[1];
    exponent = Number(sciMatch[2]);
  }

  // Now `s` is the mantissa — only digits + possibly `,` and `.` separators.
  if (!/^[\d,.]+$/.test(s)) return null;

  const dotCount = (s.match(/\./g) ?? []).length;
  const commaCount = (s.match(/,/g) ?? []).length;

  let normalized: string;

  if (dotCount > 0 && commaCount > 0) {
    // Both separators present → the RIGHTMOST is the decimal, the other is
    // thousand-sep. Common in US "1,234.56" and EU "1.234,56".
    const lastDot = s.lastIndexOf(".");
    const lastComma = s.lastIndexOf(",");
    if (lastDot > lastComma) {
      // US form: comma = thousand, dot = decimal.
      normalized = s.replace(/,/g, "");
    } else {
      // EU form: dot = thousand, comma = decimal.
      normalized = s.replace(/\./g, "").replace(/,/g, ".");
    }
  } else if (dotCount === 1 && commaCount === 0) {
    // Only one dot — could be decimal (US) or thousand (EU "1.234").
    const [left, right] = s.split(".");
    if (localeHint === "eu") {
      // EU: dot is thousand separator, and it's valid only if right has
      // exactly 3 digits. Otherwise treat as decimal anyway (don't silently
      // misparse "1.5" as 15).
      normalized = right.length === 3 && left.length >= 1 ? s.replace(".", "") : s;
    } else if (localeHint === "auto" && right.length === 3 && left.length >= 1 && left.length <= 3) {
      // Ambiguous heuristic: "1.234" → 1234 (thousand). "1.23" → 1.23 (decimal).
      normalized = s.replace(".", "");
    } else {
      // US/default: dot is decimal.
      normalized = s;
    }
  } else if (commaCount === 1 && dotCount === 0) {
    // Only one comma — could be decimal (EU) or thousand (US "1,234").
    const [left, right] = s.split(",");
    if (localeHint === "us") {
      normalized = right.length === 3 && left.length >= 1 ? s.replace(",", "") : s.replace(",", ".");
    } else if (localeHint === "auto" && right.length === 3 && left.length >= 1 && left.length <= 3) {
      normalized = s.replace(",", ""); // thousand form
    } else {
      // EU/default: comma is decimal.
      normalized = s.replace(",", ".");
    }
  } else if (dotCount > 1 && commaCount === 0) {
    // Multiple dots, no comma → must be EU thousand form like "1.234.567".
    // Validate: every segment after the first is exactly 3 digits, otherwise
    // this is malformed (e.g. "1.2.3" has 1-digit segments → null).
    const parts = s.split(".");
    const allMiddleAreTriplets = parts.slice(1).every((p) => /^\d{3}$/.test(p));
    if (!allMiddleAreTriplets) return null;
    normalized = s.replace(/\./g, "");
  } else if (commaCount > 1 && dotCount === 0) {
    // Multiple commas, no dot → US thousand form like "1,234,567" OR Indian
    // "1,23,456". Validate: after the first segment, each must be 2 or 3
    // digits (Indian lakh uses 2, Western uses 3).
    const parts = s.split(",");
    const allMiddleOK = parts.slice(1).every((p) => /^\d{2,3}$/.test(p));
    if (!allMiddleOK) return null;
    normalized = s.replace(/,/g, "");
  } else {
    normalized = s;
  }

  const n = parseFloat(normalized);
  if (!Number.isFinite(n)) return null;

  const withSign = negative ? -n : n;
  const withExp = withSign * Math.pow(10, exponent);
  return withExp / percentDivisor;
}
