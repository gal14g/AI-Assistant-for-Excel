/**
 * normalizeString — aggressive canonicalization for equality and comparison.
 *
 * Covers the frequent reasons two strings that LOOK identical don't compare
 * equal in JS / Excel:
 *   - Leading/trailing whitespace
 *   - Internal whitespace collapsed (double spaces, NBSP, narrow-NBSP, tabs)
 *   - Unicode NFC normalization (composed vs decomposed Hebrew niqud, accents)
 *   - Bidirectional control marks stripped (LRM U+200E, RLM U+200F, LRE/RLE/PDF)
 *   - Zero-width characters stripped (ZWJ U+200D, ZWNJ U+200C, BOM U+FEFF, ZWSP U+200B)
 *
 * Never changes letter-case. Call `.toLowerCase()` yourself if you need
 * case-insensitive compare — case-folding rules for non-Latin scripts are
 * locale-dependent, and this utility stays agnostic.
 *
 * Accepts any value; non-strings return an empty string (defensive).
 */

// Bidi control + zero-width + BOM characters — all invisible glue that breaks
// equality checks when one side has them and the other doesn't.
// U+200B ZWSP, U+200C ZWNJ, U+200D ZWJ, U+200E LRM, U+200F RLM,
// U+202A-U+202E bidi embedding/override, U+2066-U+2069 bidi isolation,
// U+FEFF BOM / word-joiner.
const INVISIBLE_RE = /[\u200B-\u200F\u202A-\u202E\u2066-\u2069\uFEFF]/g;

// Whitespace variants to collapse: regular space, tab, CR/LF, NBSP (U+00A0),
// narrow NBSP (U+202F), thin space (U+2009), hair space (U+200A),
// ideographic space (U+3000).
const WS_RE = /[\s\u00A0\u202F\u2009\u200A\u3000]+/g;

export function normalizeString(value: unknown): string {
  if (value === null || value === undefined) return "";
  const s = typeof value === "string" ? value : String(value);

  // Unicode NFC — combining marks fold into their composed form.
  // This matters for Hebrew (niqud) and Latin with accents (é vs e + ´).
  const nfc = s.normalize("NFC");

  // Strip invisible control characters.
  const visible = nfc.replace(INVISIBLE_RE, "");

  // Collapse whitespace to single spaces, then trim.
  return visible.replace(WS_RE, " ").trim();
}

/** Case-insensitive normalize for equality checks — trims, strips invisibles,
 *  collapses whitespace, lowercases. Handles Hebrew correctly (Hebrew has
 *  no case distinction, so `.toLowerCase()` is a no-op there). */
export function normalizeForCompare(value: unknown): string {
  return normalizeString(value).toLowerCase();
}
