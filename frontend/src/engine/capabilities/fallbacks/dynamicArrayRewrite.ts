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

/**
 * LET(name1, expr1, name2, expr2, …, body) — Excel 365-only function that
 * binds names to expressions and evaluates `body` with them in scope.
 *
 * Rewrite strategy: textually inline each name. `LET(prevIdx, X, prevD, Y, BODY)`
 * becomes `BODY` with every whole-word `prevIdx` replaced by `(X)` and every
 * `prevD` replaced by `(Y)`. The parens preserve operator precedence. Because
 * later LET bindings can reference earlier ones, we apply substitutions
 * in-order so expr_{i+1} sees the already-inlined form of expr_i.
 *
 * Strings (double- or single-quoted — single-quoted is for sheet names like
 * 'My Sheet') are never rewritten. Function identifiers like SUM, INDEX etc.
 * are matched by Excel's identifier regex `[A-Za-z_][A-Za-z0-9_.]*` which our
 * substitution respects via word-boundary lookahead/lookbehind.
 */
function rewriteLET(args: string[]): string {
  // LET needs an ODD number of args: (name, expr) pairs + body.
  if (args.length < 3 || args.length % 2 === 0) {
    return `LET(${args.join(",")})`; // malformed — give up, Excel will #VALUE! it
  }
  const body = args[args.length - 1];
  const pairs: Array<[string, string]> = [];
  for (let i = 0; i < args.length - 1; i += 2) {
    const name = args[i].trim();
    const expr = args[i + 1];
    if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name)) {
      // Unexpected — a LET binding name must be a valid identifier.
      return `LET(${args.join(",")})`;
    }
    pairs.push([name, expr]);
  }

  // Substitute bindings progressively. Later bindings' expressions may
  // reference earlier names; inline those FIRST so later substitutions see
  // the resolved form.
  const resolved: Array<[string, string]> = [];
  for (const [name, expr] of pairs) {
    let inlined = expr;
    for (const [prevName, prevExpr] of resolved) {
      inlined = inlineIdentifier(inlined, prevName, prevExpr);
    }
    resolved.push([name, inlined]);
  }
  // Finally, inline every binding into the body.
  let outBody = body;
  for (const [name, expr] of resolved) {
    outBody = inlineIdentifier(outBody, name, expr);
  }
  return outBody;
}

/**
 * Replace every occurrence of `name` in `formula` with `(replacement)`,
 * respecting Excel identifier word boundaries AND skipping matches inside
 * string literals (double-quoted and single-quoted).
 */
function inlineIdentifier(formula: string, name: string, replacement: string): string {
  const pattern = new RegExp(`(?<![A-Za-z0-9_.])${name}(?![A-Za-z0-9_.])`, "g");
  let out = "";
  let i = 0;
  let inDQ = false;
  let inSQ = false;
  while (i < formula.length) {
    const ch = formula[i];
    if (inDQ) {
      out += ch;
      if (ch === '"') inDQ = false;
      i += 1;
      continue;
    }
    if (inSQ) {
      out += ch;
      if (ch === "'") inSQ = false;
      i += 1;
      continue;
    }
    if (ch === '"') { inDQ = true; out += ch; i += 1; continue; }
    if (ch === "'") { inSQ = true; out += ch; i += 1; continue; }

    // Outside any string — try to match `name` starting at i.
    pattern.lastIndex = i;
    const m = pattern.exec(formula);
    if (m && m.index === i) {
      out += `(${replacement})`;
      i += m[0].length;
      continue;
    }
    out += ch;
    i += 1;
  }
  return out;
}

/**
 * INDEX(range, MATCH(1, <array_expr>, 0)) → LOOKUP(2, 1/(<array_expr>), range)
 *
 * On Excel 365 the original works natively. On pre-365 Excel, `MATCH(1,
 * (cond1)*(cond2), 0)` needs array-formula entry (Ctrl+Shift+Enter) or it
 * silently returns #N/A. The LOOKUP(2, 1/X, range) idiom is the classic
 * workaround: LOOKUP is inherently array-aware, errors from 1/FALSE are
 * skipped, and lookup_value=2 greater than the max possible result (1)
 * returns the last matching position.
 *
 * Caveat: LOOKUP returns the LAST matching row; MATCH(1, …, 0) returns the
 * FIRST. For self-join scenarios with a unique predecessor per row, they
 * agree. For multi-match lookups the legacy behavior differs — we emit a
 * warning.
 */
function rewriteIndexMatchArray(formula: string): { formula: string; changes: string[]; warnings: string[] } {
  let out = formula;
  const changes: string[] = [];
  const warnings: string[] = [];
  let index = findCall(out, "INDEX");
  while (index) {
    // INDEX(range, row_num [, col_num]) — we only touch the 2-arg form whose
    // row_num is a MATCH(1, X, 0) call.
    if (index.args.length !== 2) {
      // Advance past this INDEX so we don't loop.
      const skipFrom = index.end;
      const next = findCall(out.slice(skipFrom), "INDEX");
      if (!next) break;
      index = { start: next.start + skipFrom, end: next.end + skipFrom, args: next.args };
      continue;
    }
    const [rangeArg, rowArg] = index.args;
    // Does rowArg look like MATCH(1, X, 0) ?
    const matchInner = tryParseMatchOneZero(rowArg);
    if (!matchInner) {
      const skipFrom = index.end;
      const next = findCall(out.slice(skipFrom), "INDEX");
      if (!next) break;
      index = { start: next.start + skipFrom, end: next.end + skipFrom, args: next.args };
      continue;
    }
    const replacement = `LOOKUP(2,1/(${matchInner}),${rangeArg})`;
    out = out.slice(0, index.start) + replacement + out.slice(index.end);
    changes.push("INDEX/MATCH(1,array,0) → LOOKUP(2,1/array,range)");
    // Restart from the top since indices shifted.
    index = findCall(out, "INDEX");
  }
  if (changes.length) {
    warnings.push(
      "Rewrote INDEX/MATCH array-criteria lookup to LOOKUP(2, 1/(…), range) for " +
      "pre-365 compatibility. This returns the LAST matching row (the 365 original " +
      "returned the FIRST). For self-join or unique-key scenarios the behavior is " +
      "identical; for multi-match lookups verify the result.",
    );
  }
  return { formula: out, changes, warnings };
}

/** Strip outer parens that wrap an entire expression: "((x))" → "x".
 *  Only strips if the outer parens are a matched pair enclosing the whole
 *  string — never strips through something like "(a)+(b)". */
function stripOuterParens(expr: string): string {
  let s = expr.trim();
  while (s.startsWith("(") && s.endsWith(")")) {
    let depth = 0;
    let matched = true;
    for (let i = 0; i < s.length; i++) {
      if (s[i] === "(") depth += 1;
      else if (s[i] === ")") {
        depth -= 1;
        if (depth === 0 && i < s.length - 1) { matched = false; break; }
      }
    }
    if (!matched) break;
    s = s.slice(1, -1).trim();
  }
  return s;
}

/** If `expr` is syntactically `MATCH(1, X, 0)` (or the same wrapped in
 *  parens from a prior LET inlining) return X, else null. */
function tryParseMatchOneZero(expr: string): string | null {
  const trimmed = stripOuterParens(expr);
  if (!/^MATCH\s*\(/i.test(trimmed)) return null;
  const call = findCall(trimmed, "MATCH");
  if (!call || call.start !== 0 || call.end !== trimmed.length) return null;
  if (call.args.length !== 3) return null;
  const [lookup, arr, mtype] = call.args.map((s) => s.trim());
  if (lookup !== "1" || mtype !== "0") return null;
  return arr;
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
    // LET is 365-only too (evaluates to #NAME? on 2016/2019); we inline its
    // bindings during the rewrite pass.
    "LET(",
  ].some((fn) => {
    const idx = upper.indexOf(fn);
    if (idx === -1) return false;
    const prev = idx > 0 ? upper[idx - 1] : "";
    // Ensure this is the start of an identifier, not a longer name.
    return !(prev && /[A-Z0-9_.]/.test(prev));
  });
}

/**
 * Return true if the formula uses an INDEX(range, MATCH(1, array-product, 0))
 * pattern — this evaluates only in array-formula mode on pre-365 Excel,
 * so we preemptively rewrite it to the LOOKUP(2, 1/…) idiom that works
 * everywhere without Ctrl+Shift+Enter.
 */
export function usesArrayIndexMatch(formula: string): boolean {
  // Heuristic: has `INDEX(`, and within some `INDEX(…)` the second arg is
  // `MATCH(1,…,0)`. Cheap pre-check without full parse — the real check runs
  // inside rewriteIndexMatchArray.
  return /INDEX\s*\([^)]*MATCH\s*\(\s*1\s*,/i.test(formula);
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

  // LET first — inlining its bindings may expose XLOOKUP/INDEX/MATCH
  // patterns that need further rewriting. Guard against infinite loop if
  // the rewrite declines to touch a malformed LET (returns the input shape).
  let letCall = findCall(out, "LET");
  while (letCall) {
    const before = out;
    const replacement = rewriteLET(letCall.args);
    out = out.slice(0, letCall.start) + replacement + out.slice(letCall.end);
    if (out === before) break; // malformed LET left as-is; don't re-scan it
    changes.push("LET → inlined bindings");
    letCall = findCall(out, "LET");
  }

  // XLOOKUP — scalar-returning 365 function, most common dynamic-array case.
  let xl = findCall(out, "XLOOKUP");
  while (xl) {
    const replacement = rewriteXLOOKUP(xl.args);
    out = out.slice(0, xl.start) + replacement + out.slice(xl.end);
    changes.push("XLOOKUP → INDEX/MATCH");
    xl = findCall(out, "XLOOKUP");
  }

  // INDEX/MATCH with array criteria (`MATCH(1, (cond)*(cond), 0)`) needs
  // array-formula entry on pre-365 Excel. Rewrite to the LOOKUP(2, 1/…)
  // idiom that works without Ctrl+Shift+Enter.
  const imRewrite = rewriteIndexMatchArray(out);
  out = imRewrite.formula;
  for (const c of imRewrite.changes) changes.push(c);
  for (const w of imRewrite.warnings) warnings.push(w);

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
