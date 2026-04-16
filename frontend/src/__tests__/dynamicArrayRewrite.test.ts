/**
 * Tests for the dynamic-array → legacy formula rewriter.
 *
 * These verify that each of the five 365-only functions
 * (FILTER, UNIQUE, XLOOKUP, SORT, SEQUENCE) is rewritten into a legacy
 * equivalent that works on Excel 2016/2019.
 */

import {
  usesDynamicArray,
  rewriteDynamicArrayFormula,
} from "../engine/capabilities/fallbacks/dynamicArrayRewrite";

describe("usesDynamicArray", () => {
  it("detects FILTER / UNIQUE / XLOOKUP / SORT / SEQUENCE", () => {
    expect(usesDynamicArray("=FILTER(A:A, B:B>0)")).toBe(true);
    expect(usesDynamicArray("=UNIQUE(A:A)")).toBe(true);
    expect(usesDynamicArray("=XLOOKUP(A1,B:B,C:C)")).toBe(true);
    expect(usesDynamicArray("=SORT(A:A)")).toBe(true);
    expect(usesDynamicArray("=SEQUENCE(10)")).toBe(true);
  });

  it("doesn't false-positive on longer identifiers containing a match", () => {
    // "MYFILTER" or "XLOOKUPPY" should NOT be detected as FILTER/XLOOKUP.
    expect(usesDynamicArray("=MYFILTER(A:A)")).toBe(false);
    expect(usesDynamicArray("=XLOOKUPPY(1,2,3)")).toBe(false);
  });

  it("returns false for plain SUM/IF/INDEX formulas", () => {
    expect(usesDynamicArray("=SUM(A:A)")).toBe(false);
    expect(usesDynamicArray("=INDEX(A:A,MATCH(1,B:B,0))")).toBe(false);
    expect(usesDynamicArray("=IF(A1>0,A1,0)")).toBe(false);
  });
});

describe("rewriteDynamicArrayFormula", () => {
  it("rewrites XLOOKUP into INDEX/MATCH with IFERROR when if_not_found provided", () => {
    const { formula, changes } = rewriteDynamicArrayFormula(
      '=XLOOKUP(A1,B:B,C:C,"not found")',
    );
    expect(formula).toContain("INDEX(C:C,MATCH(A1,B:B,0))");
    expect(formula).toContain("IFERROR");
    expect(changes).toContain("XLOOKUP → INDEX/MATCH");
  });

  it("rewrites XLOOKUP into plain INDEX/MATCH when no if_not_found", () => {
    const { formula } = rewriteDynamicArrayFormula("=XLOOKUP(A1,B:B,C:C)");
    expect(formula).toBe("=INDEX(C:C,MATCH(A1,B:B,0))");
  });

  it("rewrites SEQUENCE(n) into ROW(INDIRECT)", () => {
    const { formula, changes } = rewriteDynamicArrayFormula("=SEQUENCE(10)");
    expect(formula).toContain('ROW(INDIRECT("1:"&10))');
    expect(changes).toContain("SEQUENCE → ROW(INDIRECT)");
  });

  it("rewrites FILTER into INDEX/SMALL/IF array-formula", () => {
    const { formula, changes, requiresArrayEntry } =
      rewriteDynamicArrayFormula("=FILTER(A:A,B:B>0)");
    expect(formula).toContain("INDEX(A:A");
    expect(formula).toContain("SMALL(IF(B:B>0");
    expect(requiresArrayEntry).toBe(true);
    expect(changes).toContain("FILTER → INDEX/SMALL/IF array formula");
  });

  it("rewrites UNIQUE into INDEX/MATCH/COUNTIF", () => {
    const { formula, requiresArrayEntry } =
      rewriteDynamicArrayFormula("=UNIQUE(A:A)");
    expect(formula).toContain("INDEX(A:A");
    expect(formula).toContain("COUNTIF");
    expect(requiresArrayEntry).toBe(true);
  });

  it("rewrites SORT descending via LARGE instead of SMALL", () => {
    const { formula } = rewriteDynamicArrayFormula("=SORT(A:A,1,-1)");
    expect(formula).toContain("LARGE(A:A");
  });

  it("rewrites SORT ascending via SMALL by default", () => {
    const { formula } = rewriteDynamicArrayFormula("=SORT(A:A)");
    expect(formula).toContain("SMALL(A:A");
  });

  it("leaves non-dynamic-array formulas alone", () => {
    const { formula, changes } = rewriteDynamicArrayFormula("=SUM(A:A)+1");
    expect(formula).toBe("=SUM(A:A)+1");
    expect(changes).toHaveLength(0);
  });

  it("emits performance warnings for FILTER/UNIQUE/SORT", () => {
    const { warnings } = rewriteDynamicArrayFormula("=FILTER(A:A,B:B>0)");
    expect(warnings.join(" ")).toMatch(/array-formula|drag|fill/i);
  });
});
