/**
 * matchRecords – Lookup/match records between two ranges.
 *
 * Formula strategy (preferFormula = true, default):
 *   Writes XLOOKUP formulas so the user gets native, recalculating lookups.
 *   On first use the capability probes whether XLOOKUP is available.
 *   If not (Excel 2016/2019), it automatically falls back to VLOOKUP.
 *   The detection result is cached for the session so subsequent calls
 *   go straight to the right formula without any extra round-trip.
 *
 * Value strategy (preferFormula = false):
 *   Reads both ranges, performs the match in JS, and writes result values.
 *   Works on every Excel version. Cells don't auto-update when source changes.
 *
 * XLOOKUP vs VLOOKUP difference:
 *   XLOOKUP – lookup and return arrays are separate, any order, match mode arg.
 *   VLOOKUP – lookup col must be leftmost; return col is an index into the table.
 *   We build the VLOOKUP table_array dynamically to span key→return col.
 */

import {
  CapabilityMeta,
  MatchRecordsParams,
  StepResult,
  ExecutionOptions,
} from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange, resolveSheet, stripWorkbookQualifier, quoteSheetInRef } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "matchRecords",
  description: "Lookup and match records between ranges using XLOOKUP (with VLOOKUP fallback) or computed match",
  mutates: true,
  affectsFormatting: false,
};

/**
 * Session-level cache for XLOOKUP availability.
 * null  = not yet probed
 * true  = XLOOKUP works on this Excel
 * false = XLOOKUP not available; use VLOOKUP
 */
let xlookupAvailable: boolean | null = null;

async function handler(
  context: Excel.RequestContext,
  params: MatchRecordsParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const { lookupRange, sourceRange, outputRange, preferFormula = true } = params;
  // Default to returning the first non-key column (index 1 = first column of sourceRange)
  const returnColumns = params.returnColumns?.length ? params.returnColumns : [1];

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would match records from ${lookupRange} against ${sourceRange}, output to ${outputRange}`,
    };
  }

  // Multi-column composite key matching: when either range spans >1 column
  // or a constant writeValue is requested, use deterministic JS matching —
  // XLOOKUP/VLOOKUP only support single-column lookups.
  const lookupCols = countColumnsInAddr(stripWorkbookQualifier(lookupRange));
  const sourceCols = countColumnsInAddr(stripWorkbookQualifier(sourceRange));

  if (lookupCols > 1 || sourceCols > 1 || params.writeValue !== undefined) {
    return await compositeKeyMatch(context, { ...params, returnColumns }, options);
  }

  if (preferFormula) {
    return await formulaMatch(context, { ...params, returnColumns }, options);
  } else {
    return await computedMatch(context, { ...params, returnColumns }, options);
  }
}

// ---------------------------------------------------------------------------
// Formula-based match (XLOOKUP → VLOOKUP fallback)
// ---------------------------------------------------------------------------

async function formulaMatch(
  context: Excel.RequestContext,
  params: MatchRecordsParams & { returnColumns: number[] },
  options: ExecutionOptions
): Promise<StepResult> {
  options.onProgress?.("Building lookup formulas...");

  // Strip workbook qualifiers — formula strings reference ranges within the
  // same workbook, so "[WorkbookName.xlsx]Sheet1!A:A" → "Sheet1!A:A".
  const lookupAddr = stripWorkbookQualifier(params.lookupRange);
  const sourceAddr = stripWorkbookQualifier(params.sourceRange);
  const outputAddr = stripWorkbookQualifier(params.outputRange);
  const { returnColumns } = params;
  const matchMode = params.matchType === "exact" ? "0" : "1";

  // Determine actual row count.
  // Full-column refs like "A:A" have rowCount = 1,048,576.
  // getUsedRange(true) returns only the data-filled sub-range.
  let rowCount = 0;
  try {
    const lookupRng = resolveRange(context, params.lookupRange);
    const lookupUsed = lookupRng.getUsedRange(true);
    lookupUsed.load("rowCount");
    await context.sync();
    rowCount = lookupUsed.rowCount;
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    if (
      msg.includes("doesn't exist") ||
      msg.includes("ItemNotFound") ||
      msg.includes("not found")
    ) {
      return {
        stepId: "",
        status: "success",
        message: `No data found in lookup range: ${params.lookupRange}`,
      };
    }
    throw new Error(`Failed to read lookup range "${params.lookupRange}": ${msg}`);
  }

  if (rowCount === 0) {
    return { stepId: "", status: "success", message: "No rows to process." };
  }

  const preciseOutputAddr = buildOutputRange(outputAddr, rowCount, returnColumns.length);
  const outputRng = resolveRange(context, preciseOutputAddr);

  // If we already know XLOOKUP isn't available, skip straight to VLOOKUP.
  if (xlookupAvailable === false) {
    outputRng.formulas = buildVlookupFormulas(lookupAddr, sourceAddr, returnColumns, matchMode, rowCount);
    await context.sync();
    options.onProgress?.(`Wrote ${rowCount} VLOOKUP formulas`);
    return {
      stepId: "",
      status: "success",
      message: `Created ${rowCount} VLOOKUP formulas in ${outputAddr}`,
    };
  }

  // Try XLOOKUP (first time or already confirmed available).
  outputRng.formulas = buildXlookupFormulas(lookupAddr, sourceAddr, returnColumns, matchMode, rowCount);
  await context.sync();

  // First time: probe whether XLOOKUP was accepted by reading the first output cell.
  // If the cell shows "#NAME?" Excel doesn't know XLOOKUP — fall back to VLOOKUP.
  if (xlookupAvailable === null) {
    options.onProgress?.("Checking XLOOKUP availability...");
    const checkCell = resolveRange(context, buildOutputRange(outputAddr, 1, 1));
    checkCell.load("values");
    await context.sync();

    const firstVal = checkCell.values?.[0]?.[0];
    xlookupAvailable = firstVal !== "#NAME?";

    if (!xlookupAvailable) {
      // Rewrite with VLOOKUP
      outputRng.formulas = buildVlookupFormulas(lookupAddr, sourceAddr, returnColumns, matchMode, rowCount);
      await context.sync();
      options.onProgress?.(`Rewrote ${rowCount} formulas as VLOOKUP (XLOOKUP not available)`);
      return {
        stepId: "",
        status: "success",
        message: `Created ${rowCount} VLOOKUP formulas in ${outputAddr} (XLOOKUP not supported on this Excel version — automatically used VLOOKUP instead)`,
      };
    }
  }

  options.onProgress?.(`Wrote ${rowCount} XLOOKUP formulas`);
  return {
    stepId: "",
    status: "success",
    message: `Created ${rowCount} XLOOKUP formulas in ${outputAddr}`,
  };
}

// ---------------------------------------------------------------------------
// Value-based match (pure JS — works on all Excel versions)
// ---------------------------------------------------------------------------

async function computedMatch(
  context: Excel.RequestContext,
  params: MatchRecordsParams & { returnColumns: number[] },
  options: ExecutionOptions
): Promise<StepResult> {
  options.onProgress?.("Reading source data...");

  // Use getUsedRange(false) to avoid loading 1M rows for full-column refs.
  // Load address on the lookup range so we know its actual start row.
  let lookupUsed: Excel.Range;
  let sourceUsed: Excel.Range;
  try {
    const lookupRng = resolveRange(context, params.lookupRange);
    const sourceRng = resolveRange(context, params.sourceRange);
    lookupUsed = lookupRng.getUsedRange(false);
    sourceUsed = sourceRng.getUsedRange(false);
    lookupUsed.load("values, address");
    sourceUsed.load("values");
    await context.sync();
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    if (msg.includes("ItemNotFound") || msg.includes("not found") || msg.includes("doesn't exist")) {
      return { stepId: "", status: "success", message: `No data found in lookup range: ${params.lookupRange}` };
    }
    throw new Error(`Failed to read ranges: ${msg}`);
  }

  const lookupValues = lookupUsed.values ?? [];
  const sourceValues = sourceUsed.values ?? [];
  const { returnColumns } = params;

  // Parse actual start row from the bounding-box address
  const startRow = (() => {
    const raw = lookupUsed.address ?? "";
    const cellPart = raw.includes("!") ? raw.split("!").pop()! : raw;
    const m = cellPart.replace(/\$/g, "").match(/[A-Z]+(\d+)/);
    return m ? parseInt(m[1], 10) : 1;
  })();

  options.onProgress?.("Matching records...");

  // Build index: source key (column 0) → full row
  const index = new Map<string, (string | number | boolean)[]>();
  for (const row of sourceValues) {
    const key = String(row[0]).toLowerCase();
    if (!index.has(key)) {
      index.set(key, row as (string | number | boolean)[]);
    }
  }

  const results: (string | number | boolean | null)[][] = [];
  for (const lookupRow of lookupValues) {
    const key = String(lookupRow[0]).toLowerCase();
    const sourceRow = index.get(key);
    const resultRow: (string | number | boolean | null)[] = [];
    for (const colIdx of returnColumns) {
      resultRow.push(sourceRow ? (sourceRow[colIdx - 1] ?? null) : null);
    }
    results.push(resultRow);
  }

  // Write output starting from the actual start row of the lookup bounding box
  const outputAddr = stripWorkbookQualifier(params.outputRange);
  const outputParts = outputAddr.split("!");
  const outSheet = outputParts.length > 1 ? outputParts[0] + "!" : "";
  const outRef   = outputParts.length > 1 ? outputParts[1] : outputParts[0];
  const outCol   = outRef.match(/[A-Z]+/)?.[0] ?? "A";
  const endCol   = offsetColumn(outCol, returnColumns.length - 1);
  const preciseOutputAddr = `${outSheet}${outCol}${startRow}:${endCol}${startRow + results.length - 1}`;

  const outputRng = resolveRange(context, preciseOutputAddr);
  outputRng.values = results;
  await context.sync();

  const matched = results.filter((r) => r.some((v) => v !== null)).length;
  return {
    stepId: "",
    status: "success",
    message: `Matched ${matched}/${lookupValues.length} records, wrote to ${preciseOutputAddr}`,
  };
}

// ---------------------------------------------------------------------------
// Composite key match (multi-column deterministic — no formulas)
// ---------------------------------------------------------------------------

/**
 * Matches rows by comparing tuples across multiple columns.
 * Handles sparse data and merged cells via getUsedRange(false) + fill-down.
 *
 * For each matched row, writes params.writeValue to the EXACT sheet row in
 * outputRange that corresponds to that lookup row.  Non-matched rows are left
 * completely untouched — no empty strings are written over existing content.
 */
async function compositeKeyMatch(
  context: Excel.RequestContext,
  params: MatchRecordsParams & { returnColumns: number[] },
  options: ExecutionOptions
): Promise<StepResult> {
  const writeValue = params.writeValue ?? "match";
  const isContains = params.matchType === "contains" || params.matchType === "approximate";
  options.onProgress?.("Reading ranges for composite match...");

  // Resolve ranges then load via the worksheet's used range as a fallback.
  // Calling getUsedRange(false) directly on whole-column refs like "B:B"
  // throws "invalid argument" on some Office.js builds — we work around this
  // by loading the worksheet's usedRange row count first and bounding manually.
  let sourceUsed: Excel.Range;
  let lookupUsed: Excel.Range;
  try {
    const sourceSheet = resolveSheet(context, params.sourceRange);
    const lookupSheet = resolveSheet(context, params.lookupRange);
    // Use OrNullObject variant so empty sheets don't throw on context.sync().
    const wsSourceUsed = sourceSheet.getUsedRangeOrNullObject(false);
    const wsLookupUsed = lookupSheet.getUsedRangeOrNullObject(false);
    wsSourceUsed.load("isNullObject, rowCount");
    wsLookupUsed.load("isNullObject, rowCount");
    await context.sync();

    // If a sheet has no data at all, fall back to a safe large bound (10 000 rows).
    const sourceRowCount = wsSourceUsed.isNullObject ? 10000 : (wsSourceUsed.rowCount || 1);
    const lookupRowCount = wsLookupUsed.isNullObject ? 10000 : (wsLookupUsed.rowCount || 1);

    // Bound any full-column address to the worksheet's used row count so
    // getUsedRange on the resulting range is guaranteed to succeed.
    const boundAddr = (addr: string, maxRow: number): string => {
      const stripped = stripWorkbookQualifier(addr);
      const parts = stripped.includes("!") ? stripped.split("!") : ["", stripped];
      const ref = parts[parts.length - 1];
      // Full-column pattern: "B:B" or "A:C" (no digits)
      if (/^[A-Z]+:[A-Z]+$/i.test(ref)) {
        const prefix = parts.length > 1 ? parts[0] + "!" : "";
        const cols = ref.split(":");
        return `${prefix}${cols[0]}1:${cols[1]}${maxRow}`;
      }
      return addr;
    };

    const boundedSource = boundAddr(params.sourceRange, sourceRowCount);
    const boundedLookup = boundAddr(params.lookupRange, lookupRowCount);

    const sourceRng = resolveRange(context, boundedSource);
    const lookupRng = resolveRange(context, boundedLookup);
    sourceUsed = sourceRng.getUsedRange(false);
    lookupUsed = lookupRng.getUsedRange(false);
    sourceUsed.load("values");
    lookupUsed.load("values, address");
    await context.sync();
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    if (msg.includes("ItemNotFound") || msg.includes("not found") || msg.includes("doesn't exist")) {
      return { stepId: "", status: "success", message: `No data found in range: ${msg}` };
    }
    throw new Error(`Failed to read ranges: ${msg}`);
  }

  // Parse the first row number from the bounding-box address.
  // Address looks like: "[Book.xlsx]Sheet!$A$2:$B$9" or "'תוכנה'!$A$2:$B$9"
  // We want the first digit group, e.g. 2 from "$A$2".
  const lookupStartRow = (() => {
    const raw = (lookupUsed as Excel.Range).address ?? "";
    // After "!" take only the cell-address part; strip $ signs; find first integer
    const cellPart = raw.includes("!") ? raw.split("!").pop()! : raw;
    const m = cellPart.replace(/\$/g, "").match(/[A-Z]+(\d+)/);
    return m ? parseInt(m[1], 10) : 1;
  })();

  type CellVal = string | number | boolean | null;
  const sourceVals = sourceUsed.values as CellVal[][];
  const lookupVals = lookupUsed.values as CellVal[][];

  options.onProgress?.("Building composite key index...");

  // Normalize: trim, lowercase, coerce to string. Join with null-byte separator
  // so "A\x00B" cannot collide with "AB".
  const normalize = (v: CellVal): string => String(v ?? "").trim().toLowerCase();
  const toKey = (row: CellVal[]): string => row.map(normalize).join("\x00");
  const isEmptyRow = (row: CellVal[]): boolean => row.every((v) => v === null || v === "");

  // Forward-fill merged cells: Office.js returns "" for every row in a merge
  // group after the first. Fill them down with the last non-empty row's values
  // so composite key matching works correctly on sheets with merged cells.
  // A row stays empty (null-row) only when there is no prior non-empty row to
  // inherit from (i.e. it is a genuine leading gap, not a merge continuation).
  const fillDown = (vals: CellVal[][]): { filled: CellVal[][]; wasEmpty: boolean[] } => {
    const filled: CellVal[][] = [];
    const wasEmpty: boolean[] = [];
    let last: CellVal[] | null = null;
    for (const row of vals) {
      if (isEmptyRow(row)) {
        wasEmpty.push(true);
        filled.push(last !== null ? [...last] : row);
      } else {
        last = row;
        wasEmpty.push(false);
        filled.push(row);
      }
    }
    return { filled, wasEmpty };
  };

  const { filled: filledSource } = fillDown(sourceVals);
  const { filled: filledLookup, wasEmpty: lookupWasEmpty } = fillDown(lookupVals);

  // Build index from forward-filled source values (exact) or flat list (contains).
  const sourceKeys = new Set<string>();
  const sourceList: string[] = [];
  for (const row of filledSource) {
    if (!isEmptyRow(row)) {
      const k = toKey(row);
      sourceKeys.add(k);
      if (isContains) sourceList.push(k);
    }
  }

  // Matcher: exact set lookup, or contains check (lookup value substring-matches
  // any source value, OR any source value is a substring of the lookup value).
  const matches = isContains
    ? (rowKey: string): boolean =>
        sourceList.some((s) => s.includes(rowKey) || rowKey.includes(s))
    : (rowKey: string): boolean => sourceKeys.has(rowKey);

  options.onProgress?.(`Matching ${filledLookup.length} rows against ${sourceKeys.size} source keys (${isContains ? "contains" : "exact"})...`);

  // Resolve the output worksheet and column letter.
  // Strip any workbook qualifier and sheet-name quotes so getItem() works.
  const outputAddrStripped = stripWorkbookQualifier(params.outputRange);
  const outputParts = outputAddrStripped.split("!");
  const outSheetName = outputParts.length > 1
    ? outputParts[0].replace(/^'|'$/g, "")   // remove surrounding single-quotes from Hebrew names
    : null;
  const outRef = outputParts.length > 1 ? outputParts[1] : outputParts[0];
  const outCol = outRef.match(/[A-Z]+/)?.[0] ?? "G";

  const outWs = outSheetName
    ? context.workbook.worksheets.getItem(outSheetName)
    : context.workbook.worksheets.getActiveWorksheet();

  // Write writeValue to the EXACT sheet row of each matched lookup row.
  // Non-matched rows are left completely untouched.
  let matchCount = 0;
  for (let i = 0; i < filledLookup.length; i++) {
    const row = filledLookup[i];
    if (isEmptyRow(row) && lookupWasEmpty[i]) continue; // genuine leading gap — skip
    if (matches(toKey(row))) {
      const sheetRow = lookupStartRow + i;
      outWs.getRange(`${outCol}${sheetRow}`).values = [[writeValue]];
      matchCount++;
    }
  }
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Composite match: ${matchCount}/${filledLookup.length} rows matched — wrote "${writeValue}" to ${outSheetName ?? "active sheet"} column ${outCol}`,
  };
}

// ---------------------------------------------------------------------------
// Formula builders
// ---------------------------------------------------------------------------

function buildXlookupFormulas(
  lookupAddr: string,
  sourceAddr: string,
  returnColumns: number[],
  matchMode: string,
  rowCount: number
): string[][] {
  const formulas: string[][] = [];
  for (let row = 0; row < rowCount; row++) {
    const rowFormulas: string[] = [];
    for (const colIdx of returnColumns) {
      // quoteSheetInRef ensures Hebrew/non-ASCII sheet names are single-quoted
      // in the formula string, e.g. "גיליון1!A:A" → "'גיליון1'!A:A"
      const lookupCell = quoteSheetInRef(getRelCellRef(lookupAddr, row));
      const lookupArr  = quoteSheetInRef(getColumnRef(sourceAddr, 0));
      const returnArr  = quoteSheetInRef(getColumnRef(sourceAddr, colIdx - 1));
      rowFormulas.push(
        `=IFERROR(XLOOKUP(${lookupCell},${lookupArr},${returnArr},"",${matchMode}),"")`
      );
    }
    formulas.push(rowFormulas);
  }
  return formulas;
}

/**
 * Build VLOOKUP formulas as a fallback for Excel versions without XLOOKUP.
 *
 * VLOOKUP requires the lookup column to be the leftmost column of the
 * table_array, and the return column is specified as an index into that array.
 * We build the table_array dynamically: Sheet1!A:B spans key col (A) to
 * return col (B), and col_index_num = colIdx (e.g. 2 for column B).
 */
function buildVlookupFormulas(
  lookupAddr: string,
  sourceAddr: string,
  returnColumns: number[],
  matchMode: string,
  rowCount: number
): string[][] {
  // VLOOKUP range_lookup: 0 = exact match, 1 = approximate (sorted)
  const exactMatch = matchMode === "0" ? "0" : "1";

  // Derive the sheet prefix and key column letter from sourceAddr
  const keyColFull = getColumnRef(sourceAddr, 0); // e.g. "Sheet1!A:A"
  const sheetPrefix = keyColFull.includes("!")
    ? keyColFull.split("!")[0] + "!"
    : "";
  const keyCol = columnLetterFromColRef(keyColFull); // e.g. "A"

  const formulas: string[][] = [];
  for (let row = 0; row < rowCount; row++) {
    const rowFormulas: string[] = [];
    for (const colIdx of returnColumns) {
      const lookupCell = quoteSheetInRef(getRelCellRef(lookupAddr, row));
      const retColFull = getColumnRef(sourceAddr, colIdx - 1); // e.g. "גיליון1!B:B"
      const retCol = columnLetterFromColRef(retColFull); // e.g. "B"
      // table_array spans key col → return col; quote sheet name if Hebrew/non-ASCII
      const rawTableArray = sheetPrefix ? `${sheetPrefix}${keyCol}:${retCol}` : `${keyCol}:${retCol}`;
      const tableArray = quoteSheetInRef(rawTableArray);
      rowFormulas.push(
        `=IFERROR(VLOOKUP(${lookupCell},${tableArray},${colIdx},${exactMatch}),"")`
      );
    }
    formulas.push(rowFormulas);
  }
  return formulas;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Extract just the column letter(s) from a full column ref like "Sheet1!B:B" */
function columnLetterFromColRef(colRef: string): string {
  const part = colRef.includes("!") ? colRef.split("!")[1] : colRef;
  return part.split(":")[0]; // "B:B" → "B"
}

/** Get a cell reference at a row offset within a range's first column.
 *  "Sheet1!A:A", offset 0 → "Sheet1!A1"
 *  "Sheet1!A2:A100", offset 2 → "Sheet1!A4"
 */
function getRelCellRef(rangeAddress: string, rowOffset: number): string {
  const parts = rangeAddress.split("!");
  const ref   = parts.length > 1 ? parts[1] : parts[0];
  const col   = ref.match(/[A-Z]+/)?.[0] ?? "A";
  const startRow = parseInt(ref.match(/\d+/)?.[0] ?? "1", 10);
  const prefix   = parts.length > 1 ? parts[0] + "!" : "";
  return `${prefix}${col}${startRow + rowOffset}`;
}

/** Get a full column reference at a column offset from a range's start column.
 *  "Sheet1!A:A", offset 1 → "Sheet1!B:B"
 */
function getColumnRef(rangeAddress: string, colOffset: number): string {
  const parts  = rangeAddress.split("!");
  const ref    = parts.length > 1 ? parts[1] : parts[0];
  const startCol = ref.match(/[A-Z]+/)?.[0] ?? "A";
  const col    = offsetColumn(startCol, colOffset);
  const prefix = parts.length > 1 ? parts[0] + "!" : "";
  return `${prefix}${col}:${col}`;
}

function offsetColumn(col: string, offset: number): string {
  let num = 0;
  for (let i = 0; i < col.length; i++) {
    num = num * 26 + (col.charCodeAt(i) - 64);
  }
  num += offset;
  let result = "";
  while (num > 0) {
    const rem = (num - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    num = Math.floor((num - 1) / 26);
  }
  return result || "A";
}

/** Convert a column letter like "A", "B", "AA" to a 1-based index. */
function colLetterToIndex(col: string): number {
  let n = 0;
  for (let i = 0; i < col.length; i++) {
    n = n * 26 + (col.charCodeAt(i) - 64);
  }
  return n;
}

/**
 * Count the number of columns spanned by a range address.
 * "A:A" → 1, "C:D" → 2, "A1:C10" → 3
 */
function countColumnsInAddr(addr: string): number {
  const ref = addr.includes("!") ? addr.split("!")[1] : addr;
  const cols = ref.match(/[A-Z]+/g) ?? ["A"];
  if (cols.length < 2) return 1;
  return colLetterToIndex(cols[1]) - colLetterToIndex(cols[0]) + 1;
}

/**
 * Build a precise output address like "Sheet2!C1:D10" so that assigning a
 * fixed-size formulas array to a range never hits a dimension mismatch.
 */
function buildOutputRange(outputAddr: string, rowCount: number, colCount: number): string {
  const parts    = outputAddr.split("!");
  const ref      = parts.length > 1 ? parts[1] : parts[0];
  const prefix   = parts.length > 1 ? parts[0] + "!" : "";
  const startCol = ref.match(/[A-Z]+/)?.[0] ?? "A";
  const startRow = parseInt(ref.match(/\d+/)?.[0] ?? "1", 10);
  const endCol   = offsetColumn(startCol, colCount - 1);
  const endRow   = startRow + rowCount - 1;
  return `${prefix}${startCol}${startRow}:${endCol}${endRow}`;
}

registry.register(meta, handler as any);
export { meta };
