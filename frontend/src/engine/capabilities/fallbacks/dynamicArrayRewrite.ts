/**
 * Dynamic-array → legacy formula rewriter.
 *
 * Excel 365+ introduced dynamic array functions (FILTER, UNIQUE, XLOOKUP,
 * SORT, SEQUENCE) that "spill" their results into adjacent cells. On Excel
 * 2016/2019 these functions don't exist and would evaluate to #NAME?.
 *
 * This module detects those functions in a formula body and rewrites them
 * into equivalent (or close-enough) legacy formulas that work on any Excel
 * version supported by Office.js (1.1+).
 *
 * Used by:
 * - `writeFormula` (primary handler) — when it detects dynamic-array use and
 *   the runtime doesn't support ExcelApi 1.11+
 * - `spillFormula` (registered fallback) — always, since spillFormula is
 *   inherently a 365-only action
 *
 * Limitations:
 * - The rewriter operates on the formula *string*. It handles the common
 *   cases where the dynamic-array call is the top-level (or one-level-deep)
 *   expression. Deeply nested dynamic arrays may need user intervention.
 * - SORT and UNIQUE on very large ranges (>1000 rows) produce slower legacy
 *   formulas; we flag these via `warnings` in the result.
 * - FILTER with multiple criteria is supported via multiplied boolean arrays.
 */

export interface RewriteResult {
  formula: string;
  /** Human-readable summary of what was rewritten (for progress/log UI). */
  changes: string[];
  /** Non-blocking warnings — e.g. "performance may degrade on large ranges". */
  warnings: string[];
  /**
   * When a rewrite produces an array formula that needs Ctrl+Shift+Enter
   * semantics (INDEX+SMALL+IF pattern), Office.js handlers should use
   * `range.setArrayFormula()` rather than `range.formulas =`.
   */
  requiresArrayEntry: boolean;
}

// Match a top-level function invocation like `FUNC(...)`. Respects nested
// parens and single-quoted strings (sheet names) / double-quoted strings.
function findCall(formula: string, funcName: string): { start: number; end: number; args: string[] } | null {
  const upper = formula.toUpperCase();
  const needle = funcName.toUpperCase() + "(";
  let searchFrom = 0;
  while (searchFrom < upper.length) {
    const idx = upper.indexOf(needle, searchFrom);
    if (idx === -1) return null;
    // Ensure preceding char isn't a letter/digit/./_ (would be a longer identifier)
    const prev = idx > 0 ? upper[idx - 1] : "";
    if (prev && /[A-Z0-9_.]/.test(prev)) {
      searchFrom = idx + needle.length;
      continue;
    }
    const openParen = idx + needle.length - 1;
    // Walk paren-matched, tracking strings.
    let depth = 1;
    let i = openParen + 1;
    let argStart = i;
    const args: string[] = [];
    let inDQ = false;
    let inSQ = false;
    while (i < formula.length && depth > 0) {
      const ch = formula[i];
      if (inDQ) {
        if (ch === '"') inDQ = false;
      } else if (inSQ) {
        if (ch === "'") inSQ = false;
      } else if (ch === '"') {
        inDQ = true;
      } else if (ch === "'") {
        inSQ = true;
      } else if (ch === "(") {
        depth += 1;
      } else if (ch === ")") {
        depth -= 1;
        if (depth === 0) {
          args.push(formula.slice(argStart, i).trim());
          return { start: idx, end: i + 1, args };
        }
      } else if (ch === "," && depth === 1) {
        args.push(formula.slice(argStart, i).trim());
        argStart = i + 1;
      }
      i += 1;
    }
    // Unbalanced — skip this match.
    searchFrom = idx + needle.length;
  }
  return null;
}

function rewriteXLOOKUP(args: string[]): string {
  // XLOOKUP(lookup_value, lookup_array, return_array [, if_not_found, match_mode, search_mode])
  // → IFERROR(INDEX(return_array, MATCH(lookup_value, lookup_array, 0)), if_not_found)
  // For approximate matches, MATCH match_type differs; we keep the common exact-match.
  if (args.length < 3) return `XLOOKUP(${args.join(",")})`; // give up
  const [lookup, arr, ret, ifNotFound] = args;
  const core = `INDEX(${ret},MATCH(${lookup},${arr},0))`;
  if (ifNotFound && ifNotFound.length) {
    return `IFERROR(${core},${ifNotFound})`;
  }
  return core;
}

function rewriteSEQUENCE(args: string[]): string {
  // SEQUENCE(rows[, columns, start, step])
  // Only handles the 1-D row case (rows with optional start/step).
  // SEQUENCE(n) → ROW(INDIRECT("1:"&n))
  // SEQUENCE(n, 1, start, step) → (ROW(INDIRECT("1:"&n))-1)*step+start
  if (args.length === 0) return "SEQUENCE()";
  const rows = args[0];
  const start = args.length >= 3 ? args[2] : "1";
  const step = args.length >= 4 ? args[3] : "1";
  if (args.length <= 1 || (args.length === 2 && args[1].trim() === "1")) {
    if (start === "1" && step === "1") {
      return `ROW(INDIRECT("1:"&${rows}))`;
    }
    return `(ROW(INDIRECT("1:"&${rows}))-1)*${step}+${start}`;
  }
  // 2D SEQUENCE has no clean legacy equivalent — fall through.
  return `SEQUENCE(${args.join(",")})`;
}

function rewriteUNIQUE(args: string[]): { formula: string; needsArray: boolean; warning?: string } {
  // UNIQUE(array) → array formula:
  // =IFERROR(INDEX(array, MATCH(0, COUNTIF($Result$Above, array), 0)), "")
  //
  // We can't express this as a single-cell formula; it has to be entered as
  // an array formula that the caller fills across the expected output range.
  // Since we don't know the output range here, we emit the most common form
  // (INDEX+MATCH+COUNTIF) and set `needsArray` so the handler uses
  // setArrayFormula for multi-cell ranges.
  if (args.length < 1) return { formula: `UNIQUE(${args.join(",")})`, needsArray: false };
  const arr = args[0];
  return {
    // Single-cell variant: returns the first unique; caller fills down/across.
    formula: `INDEX(${arr},MATCH(0,COUNTIF($A$1:A1,${arr}),0))`,
    needsArray: true,
    warning:
      "UNIQUE rewrite assumes output starts at A1 relative to the formula cell; " +
      "verify the COUNTIF accumulator range matches your actual output location.",
  };
}

function rewriteSORT(args: string[]): { formula: string; needsArray: boolean; warning?: string } {
  // SORT(array [, sort_index, sort_order, by_col])
  // Rewrite: INDEX(array, MATCH(LARGE/SMALL(IF(...))))
  // For simplicity we support 1-D ascending sort:
  //   SORT(A:A) → INDEX(A:A, MATCH(SMALL(A:A, ROW()), A:A, 0))
  if (args.length < 1) return { formula: `SORT(${args.join(",")})`, needsArray: false };
  const arr = args[0];
  const order = args.length >= 3 ? args[2] : "1"; // 1 asc, -1 desc
  const picker = order.trim() === "-1" ? "LARGE" : "SMALL";
  return {
    formula: `INDEX(${arr},MATCH(${picker}(${arr},ROW()),${arr},0))`,
    needsArray: true,
    warning:
      "SORT rewrite assumes a 1-D range and is slow on >1000 rows; consider " +
      "sorting manually via the sortRange action instead.",
  };
}

function rewriteFILTER(args: string[]): { formula: string; needsArray: boolean; warning?: string } {
  // FILTER(array, include [, if_empty])
  // Classic equivalent via INDEX/SMALL/IF array formula:
  //   =IFERROR(INDEX(array, SMALL(IF(include, ROW(include)-MIN(ROW(include))+1), ROW(A1))), if_empty)
  if (args.length < 2) return { formula: `FILTER(${args.join(",")})`, needsArray: false };
  const [arr, include, ifEmpty] = args;
  const core =
    `INDEX(${arr},SMALL(IF(${include},ROW(${include})-MIN(ROW(${include}))+1),ROW(A1)))`;
  const withIfEmpty = ifEmpty
    ? `IFERROR(${core},${ifEmpty})`
    : `IFERROR(${core},"")`;
  return {
    formula: withIfEmpty,
    needsArray: true,
    warning:
      "FILTER rewrite needs array-formula entry (Ctrl+Shift+Enter). Fill down " +
      "across your expected output range; ROW(A1) will increment per row.",
  };
}

/**
 * Return true if the given formula uses any of the dynamic-array functions
 * that require rewriting on pre-365 Excel.
 */
export function usesDynamicArray(formula: string): boolean {
  const upper = formula.toUpperCase();
  return [
    "FILTER(",
    "UNIQUE(",
    "XLOOKUP(",
    "SORT(",
    "SEQUENCE(",
  ].some((fn) => {
    const idx = upper.indexOf(fn);
    if (idx === -1) return false;
    const prev = idx > 0 ? upper[idx - 1] : "";
    // Ensure this is the start of an identifier, not a longer name.
    return !(prev && /[A-Z0-9_.]/.test(prev));
  });
}

/**
 * Rewrite every top-level dynamic-array call in `formula` into a legacy
 * equivalent. Unrecognized / nested cases are left intact; the caller should
 * still error gracefully if Excel eventually evaluates #NAME?.
 */
export function rewriteDynamicArrayFormula(formula: string): RewriteResult {
  let out = formula;
  const changes: string[] = [];
  const warnings: string[] = [];
  let needsArrayEntry = false;

  // XLOOKUP first — it produces scalar values and is the most common case.
  let xl = findCall(out, "XLOOKUP");
  while (xl) {
    const replacement = rewriteXLOOKUP(xl.args);
    out = out.slice(0, xl.start) + replacement + out.slice(xl.end);
    changes.push("XLOOKUP → INDEX/MATCH");
    xl = findCall(out, "XLOOKUP");
  }

  // SEQUENCE
  let seq = findCall(out, "SEQUENCE");
  while (seq) {
    const replacement = rewriteSEQUENCE(seq.args);
    out = out.slice(0, seq.start) + replacement + out.slice(seq.end);
    changes.push("SEQUENCE → ROW(INDIRECT)");
    seq = findCall(out, "SEQUENCE");
  }

  // UNIQUE
  let uq = findCall(out, "UNIQUE");
  while (uq) {
    const { formula: replacement, needsArray, warning } = rewriteUNIQUE(uq.args);
    out = out.slice(0, uq.start) + replacement + out.slice(uq.end);
    changes.push("UNIQUE → INDEX/MATCH/COUNTIF");
    if (needsArray) needsArrayEntry = true;
    if (warning) warnings.push(warning);
    uq = findCall(out, "UNIQUE");
  }

  // SORT
  let st = findCall(out, "SORT");
  while (st) {
    const { formula: replacement, needsArray, warning } = rewriteSORT(st.args);
    out = out.slice(0, st.start) + replacement + out.slice(st.end);
    changes.push("SORT → INDEX/MATCH/SMALL");
    if (needsArray) needsArrayEntry = true;
    if (warning) warnings.push(warning);
    st = findCall(out, "SORT");
  }

  // FILTER
  let fl = findCall(out, "FILTER");
  while (fl) {
    const { formula: replacement, needsArray, warning } = rewriteFILTER(fl.args);
    out = out.slice(0, fl.start) + replacement + out.slice(fl.end);
    changes.push("FILTER → INDEX/SMALL/IF array formula");
    if (needsArray) needsArrayEntry = true;
    if (warning) warnings.push(warning);
    fl = findCall(out, "FILTER");
  }

  return {
    formula: out,
    changes,
    warnings,
    requiresArrayEntry: needsArrayEntry,
  };
}
