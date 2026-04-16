/**
 * Shared utility: detect full-column references in a formula and rewrite
 * them to their target sheet's used-range bounds.
 *
 * Why: a dynamic-array formula whose source is `B:B` spills 1,048,576 rows.
 * Even a non-dynamic formula like `=SUMPRODUCT((A:A="X")*B:B)` evaluates
 * against 1M rows and is slow. Clipping to used-range is almost always
 * what the user meant.
 *
 * The regex matches column-letter-to-column-letter (optionally prefixed by
 * a sheet name + optional $ absolute markers) while refusing to match
 * bounded references like "A1:A10" (non-alphanumeric boundaries around the
 * match reject digits immediately adjacent to the colon).
 *
 * Matches:     A:A, $A:$C, AA:AB, Sheet1!B:D, 'My Sheet'!A:A
 * Does NOT match: A1:A10, Sheet1!B2:B100, $A$1:$A$10
 */

const FULL_COL_RE =
  /(?<![A-Za-z0-9$_])(?:('[^']+'|[A-Za-z\u0590-\u05FF][\w\u0590-\u05FF]*)!)?(\$?[A-Z]+)\s*:\s*(\$?[A-Z]+)(?![A-Za-z0-9_])/g;

export interface ClipResult {
  formula: string;
  clippedCount: number;
  clippedRefs: string[];
  warnings: string[];
}

/** Rewrite every full-column reference to the sheet's used-range row bounds.
 *  Returns the formula unchanged if no full-column refs are found. */
export async function clipFullColumnRefs(
  context: Excel.RequestContext,
  formula: string,
  defaultSheetName: string | null,
): Promise<ClipResult> {
  // Reset RegExp.lastIndex — a global RE keeps state across calls.
  FULL_COL_RE.lastIndex = 0;
  const matches: RegExpExecArray[] = [];
  let m: RegExpExecArray | null;
  while ((m = FULL_COL_RE.exec(formula)) !== null) matches.push(m);
  if (matches.length === 0) {
    return { formula, clippedCount: 0, clippedRefs: [], warnings: [] };
  }

  // Collect unique sheet targets. Unqualified refs resolve to
  // defaultSheetName or the active sheet.
  const sheetKeyFor = (raw: string | undefined): string => {
    if (!raw) return defaultSheetName ?? "__ACTIVE__";
    return raw.replace(/^'|'$/g, "");
  };
  const uniqueSheets = new Set<string>(matches.map((mm) => sheetKeyFor(mm[1])));

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const usedMap = new Map<string, any>();
  for (const key of uniqueSheets) {
    const ws =
      key === "__ACTIVE__"
        ? context.workbook.worksheets.getActiveWorksheet()
        : context.workbook.worksheets.getItemOrNullObject(key);
    const used = ws.getUsedRangeOrNullObject(true);
    used.load(["isNullObject", "rowCount", "address"]);
    usedMap.set(key, { ws, used });
  }
  await context.sync();

  const lastRowBySheet = new Map<string, number>();
  for (const [key, { used }] of usedMap) {
    if (used.isNullObject || !used.rowCount) {
      lastRowBySheet.set(key, 1);
    } else {
      const addr: string = used.address ?? "";
      const tail = addr.match(/(\d+)\s*$/);
      lastRowBySheet.set(key, tail ? Number(tail[1]) : used.rowCount);
    }
  }

  // Apply substitutions right-to-left so earlier positions aren't shifted.
  let out = formula;
  const clippedRefs: string[] = [];
  const warnings: string[] = [];
  for (let i = matches.length - 1; i >= 0; i--) {
    const mm = matches[i];
    const fullMatch = mm[0];
    const sheetPrefix = mm[1];
    const col1 = mm[2];
    const col2 = mm[3];
    const key = sheetKeyFor(sheetPrefix);
    const lastRow = lastRowBySheet.get(key) ?? 1;
    const bounded =
      (sheetPrefix ? `${sheetPrefix}!` : "") + `${col1}1:${col2}${lastRow}`;
    const start = mm.index;
    out = out.slice(0, start) + bounded + out.slice(start + fullMatch.length);
    clippedRefs.push(fullMatch);
    if (lastRow <= 1) {
      warnings.push(
        `${fullMatch} → ${bounded} (target sheet appears empty; bound to row 1 so the formula doesn't spill 1M cells)`,
      );
    }
  }

  return {
    formula: out,
    clippedCount: matches.length,
    clippedRefs,
    warnings,
  };
}
