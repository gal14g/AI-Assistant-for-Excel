/**
 * parseNumberFlexible — exhaustive format coverage.
 */

import { parseNumberFlexible } from "../engine/utils/parseNumberFlexible";

describe("parseNumberFlexible — native numerics", () => {
  it("passes through native numbers", () => {
    expect(parseNumberFlexible(1234)).toBe(1234);
    expect(parseNumberFlexible(1234.56)).toBe(1234.56);
    expect(parseNumberFlexible(-1234)).toBe(-1234);
    expect(parseNumberFlexible(0)).toBe(0);
  });

  it("rejects NaN / Infinity as native numbers", () => {
    expect(parseNumberFlexible(NaN)).toBeNull();
    expect(parseNumberFlexible(Infinity)).toBeNull();
    expect(parseNumberFlexible(-Infinity)).toBeNull();
  });

  it("coerces boolean to 0/1", () => {
    expect(parseNumberFlexible(true)).toBe(1);
    expect(parseNumberFlexible(false)).toBe(0);
  });
});

describe("parseNumberFlexible — plain text numbers", () => {
  it("'1234' → 1234", () => {
    expect(parseNumberFlexible("1234")).toBe(1234);
  });

  it("'1234.56' → 1234.56", () => {
    expect(parseNumberFlexible("1234.56")).toBe(1234.56);
  });

  it("leading zeros OK", () => {
    expect(parseNumberFlexible("0042")).toBe(42);
  });

  it("negative with - prefix", () => {
    expect(parseNumberFlexible("-100")).toBe(-100);
  });

  it("positive with + prefix", () => {
    expect(parseNumberFlexible("+100")).toBe(100);
  });
});

describe("parseNumberFlexible — US format with separators", () => {
  it("'1,234' → 1234", () => {
    expect(parseNumberFlexible("1,234")).toBe(1234);
  });

  it("'1,234.56' → 1234.56", () => {
    expect(parseNumberFlexible("1,234.56")).toBe(1234.56);
  });

  it("'1,234,567.89' → 1234567.89", () => {
    expect(parseNumberFlexible("1,234,567.89")).toBe(1234567.89);
  });

  it("Indian '1,23,456.78' → 123456.78", () => {
    expect(parseNumberFlexible("1,23,456.78")).toBe(123456.78);
  });
});

describe("parseNumberFlexible — EU format", () => {
  it("'1.234,56' unambiguous EU → 1234.56", () => {
    expect(parseNumberFlexible("1.234,56")).toBe(1234.56);
  });

  it("'1.234.567,89' → 1234567.89", () => {
    expect(parseNumberFlexible("1.234.567,89")).toBe(1234567.89);
  });

  it("'1,5' with localeHint='eu' → 1.5", () => {
    expect(parseNumberFlexible("1,5", "eu")).toBe(1.5);
  });

  it("'1.234' with localeHint='eu' → 1234 (thousand group)", () => {
    expect(parseNumberFlexible("1.234", "eu")).toBe(1234);
  });

  it("'1.23' with localeHint='eu' → 1.23 (too short to be thousand group)", () => {
    expect(parseNumberFlexible("1.23", "eu")).toBe(1.23);
  });
});

describe("parseNumberFlexible — auto-detect ambiguous cases", () => {
  it("'1.234' auto → 1234 (3 digits after dot → thousand)", () => {
    expect(parseNumberFlexible("1.234", "auto")).toBe(1234);
  });

  it("'1.23' auto → 1.23 (2 digits after dot → decimal)", () => {
    expect(parseNumberFlexible("1.23", "auto")).toBe(1.23);
  });

  it("'1,234' auto → 1234 (3 digits after comma → thousand)", () => {
    expect(parseNumberFlexible("1,234", "auto")).toBe(1234);
  });

  it("'1,5' auto → 1.5 (2 digits after comma → decimal)", () => {
    expect(parseNumberFlexible("1,5", "auto")).toBe(1.5);
  });
});

describe("parseNumberFlexible — currency prefixes and suffixes", () => {
  it("'$1,234.56' → 1234.56", () => {
    expect(parseNumberFlexible("$1,234.56")).toBe(1234.56);
  });

  it("'€1.234,56' → 1234.56", () => {
    expect(parseNumberFlexible("€1.234,56")).toBe(1234.56);
  });

  it("'₪1,234' → 1234 (Israeli shekel)", () => {
    expect(parseNumberFlexible("₪1,234")).toBe(1234);
  });

  it("'£100' → 100", () => {
    expect(parseNumberFlexible("£100")).toBe(100);
  });

  it("'100 USD' → 100 (ISO code suffix)", () => {
    expect(parseNumberFlexible("100 USD")).toBe(100);
  });

  it("'100 ILS' → 100", () => {
    expect(parseNumberFlexible("100 ILS")).toBe(100);
  });

  it("spacing after currency symbol: '₪ 1,234'", () => {
    expect(parseNumberFlexible("₪ 1,234")).toBe(1234);
  });
});

describe("parseNumberFlexible — percent", () => {
  it("'50%' → 0.5", () => {
    expect(parseNumberFlexible("50%")).toBe(0.5);
  });

  it("'100%' → 1", () => {
    expect(parseNumberFlexible("100%")).toBe(1);
  });

  it("'0.5%' → 0.005", () => {
    expect(parseNumberFlexible("0.5%")).toBe(0.005);
  });

  it("'-25%' → -0.25", () => {
    expect(parseNumberFlexible("-25%")).toBe(-0.25);
  });
});

describe("parseNumberFlexible — accounting-style negatives", () => {
  it("parenthesis negative '(100)' → -100", () => {
    expect(parseNumberFlexible("(100)")).toBe(-100);
  });

  it("parenthesis with separators '(1,234.56)' → -1234.56", () => {
    expect(parseNumberFlexible("(1,234.56)")).toBe(-1234.56);
  });

  it("parenthesis with currency '($1,234)' → -1234", () => {
    expect(parseNumberFlexible("($1,234)")).toBe(-1234);
  });

  it("trailing minus '100-' → -100", () => {
    expect(parseNumberFlexible("100-")).toBe(-100);
  });
});

describe("parseNumberFlexible — scientific notation", () => {
  it("'1.23E+06' → 1230000", () => {
    expect(parseNumberFlexible("1.23E+06")).toBe(1230000);
  });

  it("'1.23e-03' → 0.00123", () => {
    expect(parseNumberFlexible("1.23e-03")).toBeCloseTo(0.00123, 10);
  });

  it("'5E4' → 50000", () => {
    expect(parseNumberFlexible("5E4")).toBe(50000);
  });
});

describe("parseNumberFlexible — whitespace variants", () => {
  it("NBSP thousand separator '1\\u00A0234'", () => {
    expect(parseNumberFlexible("1\u00A0234")).toBe(1234);
  });

  it("narrow NBSP '1\\u202F234.56'", () => {
    expect(parseNumberFlexible("1\u202F234.56")).toBe(1234.56);
  });

  it("leading/trailing whitespace '  42  '", () => {
    expect(parseNumberFlexible("  42  ")).toBe(42);
  });
});

describe("parseNumberFlexible — invalid", () => {
  it("empty string", () => {
    expect(parseNumberFlexible("")).toBeNull();
  });

  it("letters", () => {
    expect(parseNumberFlexible("abc")).toBeNull();
  });

  it("mixed text '100 dollars' (unknown suffix)", () => {
    expect(parseNumberFlexible("100 dollars")).toBeNull();
  });

  it("multiple decimal points without thousand", () => {
    // "1.2.3" is not a valid number; we return null because after dot removal
    // we'd still have "1.23" or similar which is ambiguous and wrong.
    expect(parseNumberFlexible("1.2.3")).toBeNull();
  });

  it("null / undefined / object", () => {
    expect(parseNumberFlexible(null)).toBeNull();
    expect(parseNumberFlexible(undefined)).toBeNull();
    expect(parseNumberFlexible({})).toBeNull();
    expect(parseNumberFlexible([])).toBeNull();
  });
});

describe("parseNumberFlexible — combined features", () => {
  it("parens + currency + separators: '($1,234.56)'", () => {
    expect(parseNumberFlexible("($1,234.56)")).toBe(-1234.56);
  });

  it("percent + negative: '-50%'", () => {
    expect(parseNumberFlexible("-50%")).toBe(-0.5);
  });

  it("EU + currency: '€1.234,56'", () => {
    expect(parseNumberFlexible("€1.234,56")).toBe(1234.56);
  });

  it("EU + percent: '50,5%' → 0.505", () => {
    expect(parseNumberFlexible("50,5%")).toBe(0.505);
  });
});
