/**
 * normalizeString / normalizeForCompare — canonicalization coverage.
 */

import { normalizeString, normalizeForCompare } from "../engine/utils/normalizeString";

describe("normalizeString — whitespace", () => {
  it("trims leading/trailing whitespace", () => {
    expect(normalizeString("  hello  ")).toBe("hello");
  });

  it("collapses internal double spaces", () => {
    expect(normalizeString("hello    world")).toBe("hello world");
  });

  it("collapses tabs and newlines", () => {
    expect(normalizeString("hello\t\nworld")).toBe("hello world");
  });

  it("replaces NBSP (U+00A0) with regular space + collapses", () => {
    expect(normalizeString("hello\u00A0world")).toBe("hello world");
  });

  it("replaces narrow NBSP (U+202F)", () => {
    expect(normalizeString("hello\u202Fworld")).toBe("hello world");
  });

  it("replaces thin space (U+2009)", () => {
    expect(normalizeString("hello\u2009world")).toBe("hello world");
  });

  it("replaces ideographic space (U+3000)", () => {
    expect(normalizeString("hello\u3000world")).toBe("hello world");
  });
});

describe("normalizeString — bidi and invisibles", () => {
  it("strips LRM (U+200E)", () => {
    expect(normalizeString("abc\u200Edef")).toBe("abcdef");
  });

  it("strips RLM (U+200F)", () => {
    expect(normalizeString("abc\u200Fdef")).toBe("abcdef");
  });

  it("strips ZWJ (U+200D)", () => {
    expect(normalizeString("abc\u200Ddef")).toBe("abcdef");
  });

  it("strips ZWNJ (U+200C)", () => {
    expect(normalizeString("abc\u200Cdef")).toBe("abcdef");
  });

  it("strips ZWSP (U+200B)", () => {
    expect(normalizeString("abc\u200Bdef")).toBe("abcdef");
  });

  it("strips BOM (U+FEFF)", () => {
    expect(normalizeString("\uFEFFhello")).toBe("hello");
  });

  it("strips bidi embedding marks (U+202A-U+202E)", () => {
    expect(normalizeString("abc\u202Bdef\u202Cghi")).toBe("abcdefghi");
  });

  it("strips bidi isolation marks (U+2066-U+2069)", () => {
    expect(normalizeString("abc\u2066def\u2069ghi")).toBe("abcdefghi");
  });
});

describe("normalizeString — Hebrew", () => {
  it("preserves plain Hebrew", () => {
    expect(normalizeString("שלום עולם")).toBe("שלום עולם");
  });

  it("strips trailing RLM from Hebrew name", () => {
    expect(normalizeString("דוד\u200F")).toBe("דוד");
  });

  it("handles Hebrew with trailing whitespace", () => {
    expect(normalizeString("דוד   ")).toBe("דוד");
  });

  it("NFC normalizes Hebrew with niqud (composed vs decomposed)", () => {
    // Decomposed form: ב (U+05D1) + hiriq (U+05B4)
    const decomposed = "\u05D1\u05B4";
    // Composed form would be the same after NFC.
    const normalized = normalizeString(decomposed);
    expect(normalized).toBe(normalized.normalize("NFC"));
    expect(normalized.length).toBeGreaterThan(0);
  });

  it("Hebrew with mixed ASCII works", () => {
    expect(normalizeString("דוד 123")).toBe("דוד 123");
  });
});

describe("normalizeString — NFC normalization", () => {
  it("e + combining acute → é", () => {
    const decomposed = "e\u0301"; // e + U+0301 combining acute
    const composed = "\u00E9"; // é precomposed
    expect(normalizeString(decomposed)).toBe(composed);
  });

  it("composed and decomposed forms match after normalization", () => {
    const decomposed = "e\u0301";
    const composed = "\u00E9";
    expect(normalizeString(decomposed)).toBe(normalizeString(composed));
  });
});

describe("normalizeString — edge cases", () => {
  it("empty string", () => {
    expect(normalizeString("")).toBe("");
  });

  it("whitespace-only → empty", () => {
    expect(normalizeString("   \t\n  ")).toBe("");
  });

  it("null → empty", () => {
    expect(normalizeString(null)).toBe("");
  });

  it("undefined → empty", () => {
    expect(normalizeString(undefined)).toBe("");
  });

  it("number coerces", () => {
    expect(normalizeString(123)).toBe("123");
  });

  it("boolean coerces", () => {
    expect(normalizeString(true)).toBe("true");
  });

  it("preserves case by default", () => {
    expect(normalizeString("Hello World")).toBe("Hello World");
  });
});

describe("normalizeForCompare — case-insensitive variant", () => {
  it("lowercases Latin", () => {
    expect(normalizeForCompare("Hello World")).toBe("hello world");
  });

  it("handles Hebrew (no case concept)", () => {
    expect(normalizeForCompare("שלום")).toBe("שלום");
  });

  it("equates variants that differ only in whitespace/marks", () => {
    const a = "דוד\u200F";
    const b = "  דוד  ";
    expect(normalizeForCompare(a)).toBe(normalizeForCompare(b));
  });

  it("equates NBSP and space", () => {
    expect(normalizeForCompare("hello\u00A0world")).toBe(normalizeForCompare("hello world"));
  });
});
