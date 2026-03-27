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
import { resolveRange, stripWorkbookQualifier } from "./rangeUtils";

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

  const lookupRng = resolveRange(context, params.lookupRange);
  const sourceRng = resolveRange(context, params.sourceRange);
  lookupRng.load("values");
  sourceRng.load("values");
  await context.sync();

  const lookupValues = lookupRng.values ?? [];
  const sourceValues = sourceRng.values ?? [];
  const { returnColumns } = params;

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

  const outputAddr = stripWorkbookQualifier(params.outputRange);
  const outputRng = resolveRange(
    context,
    buildOutputRange(outputAddr, results.length, returnColumns.length)
  );
  outputRng.values = results;
  await context.sync();

  const matched = results.filter((r) => r.some((v) => v !== null)).length;
  return {
    stepId: "",
    status: "success",
    message: `Matched ${matched}/${lookupValues.length} records, wrote to ${params.outputRange}`,
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
      const lookupCell = getRelCellRef(lookupAddr, row);
      const lookupArr  = getColumnRef(sourceAddr, 0);
      const returnArr  = getColumnRef(sourceAddr, colIdx - 1);
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
      const lookupCell = getRelCellRef(lookupAddr, row);
      const retColFull = getColumnRef(sourceAddr, colIdx - 1); // e.g. "Sheet1!B:B"
      const retCol = columnLetterFromColRef(retColFull); // e.g. "B"
      // table_array spans from key col to return col: e.g. "Sheet1!A:B"
      const tableArray = `${sheetPrefix}${keyCol}:${retCol}`;
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
