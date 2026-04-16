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

describe("LET rewrite (Excel 2016/2019 compatibility)", () => {
  it("detects LET as a dynamic-array-family function needing rewrite", () => {
    expect(usesDynamicArray("=LET(x, 1, x+2)")).toBe(true);
  });

  it("inlines a single binding", () => {
    const { formula, changes } = rewriteDynamicArrayFormula("=LET(x,5,x+1)");
    expect(formula).toBe("=(5)+1");
    expect(changes).toContain("LET → inlined bindings");
  });

  it("inlines multiple bindings in order", () => {
    const { formula } = rewriteDynamicArrayFormula("=LET(a,1,b,a+2,b*3)");
    // a → 1, then b → (a)+2 → (1)+2 → written as (1)+2 before body subst,
    // then b in body → ((1)+2), so body becomes ((1)+2)*3.
    expect(formula).toContain("*3");
    expect(formula).toContain("1");
    expect(formula).not.toContain("LET(");
  });

  it("doesn't substitute inside string literals", () => {
    const { formula } = rewriteDynamicArrayFormula('=LET(x,"test","x-label")');
    // `x` inside "x-label" must stay literal.
    expect(formula).toBe('="x-label"');
  });

  it("respects identifier word boundaries (doesn't match substring)", () => {
    const { formula } = rewriteDynamicArrayFormula("=LET(x,1,xyz+x)");
    // `xyz` must NOT have its `x` prefix replaced.
    expect(formula).toBe("=xyz+(1)");
  });

  it("handles the Hebrew self-join example from the few-shot seed", () => {
    const f =
      '=IFERROR(LET(prevIdx,MATCH(1,(A$2:A$100=A2)*(C$2:C$100=B2-1),0),' +
      'prevD,INDEX(D$2:D$100,prevIdx),' +
      'IF(AND(prevD="נציבותי",D2<>"נציבותי"),"עזיבה",' +
      'IF(AND(prevD<>"נציבותי",D2="נציבותי"),"קליטה",""))),"")';
    const { formula, changes } = rewriteDynamicArrayFormula(f);
    // Both LET and the INDEX/MATCH(1,…,0) rewrites should fire.
    expect(changes).toContain("LET → inlined bindings");
    expect(changes.join(" ")).toContain("INDEX/MATCH(1,array,0)");
    expect(formula).not.toContain("LET(");
    expect(formula).not.toContain("prevIdx");
    expect(formula).not.toContain("prevD");
    expect(formula).toContain("LOOKUP(2,1/(");
  });

  it("leaves malformed LET (even arg count) alone", () => {
    // LET must have odd arg count (name/expr pairs + body).
    const { formula } = rewriteDynamicArrayFormula("=LET(x,1,y,2)");
    expect(formula).toContain("LET(");
  });
});

describe("INDEX/MATCH(1,array,0) → LOOKUP rewrite", () => {
  it("rewrites the two-criteria self-join pattern", () => {
    const f = "=INDEX(D1:D10,MATCH(1,(A1:A10=\"k\")*(B1:B10=5),0))";
    const { formula, changes } = rewriteDynamicArrayFormula(f);
    expect(formula).toBe('=LOOKUP(2,1/((A1:A10="k")*(B1:B10=5)),D1:D10)');
    expect(changes).toContain("INDEX/MATCH(1,array,0) → LOOKUP(2,1/array,range)");
  });

  it("emits a semantic warning about first-vs-last-match difference", () => {
    const { warnings } = rewriteDynamicArrayFormula(
      "=INDEX(A:A,MATCH(1,(B:B=1)*(C:C=2),0))",
    );
    expect(warnings.join(" ")).toMatch(/LAST matching|last match/i);
  });

  it("leaves plain INDEX/MATCH (scalar lookup value) alone", () => {
    // Single-criteria MATCH doesn't need this rewrite — it works on all versions.
    const { formula, changes } = rewriteDynamicArrayFormula(
      "=INDEX(B:B,MATCH(\"exact\",A:A,0))",
    );
    expect(formula).toBe('=INDEX(B:B,MATCH("exact",A:A,0))');
    expect(changes.filter((c) => c.includes("LOOKUP"))).toHaveLength(0);
  });

  it("leaves INDEX with 3 args (row + col) alone", () => {
    // INDEX(range, row, col) has a different shape — our rewriter only
    // touches the 2-arg form.
    const { formula } = rewriteDynamicArrayFormula(
      "=INDEX(A1:C10,MATCH(1,(A:A=1),0),2)",
    );
    expect(formula).toContain("INDEX(A1:C10,");
    expect(formula).not.toContain("LOOKUP");
  });
});
