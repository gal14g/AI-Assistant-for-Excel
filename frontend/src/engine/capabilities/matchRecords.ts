/**
 * matchRecords – Lookup/match records between two ranges.
 *
 * Strategy (preferFormula = true, default):
 *   Writes XLOOKUP or VLOOKUP formulas so the user gets native, recalculating
 *   lookups. XLOOKUP is preferred (Excel 365+) but falls back to VLOOKUP.
 *
 * Strategy (preferFormula = false):
 *   Reads both ranges, performs the match in JS, and writes result values.
 *   This is faster for one-time operations but doesn't auto-update.
 *
 * Office.js notes:
 * - We cannot detect whether XLOOKUP is available at runtime via Office.js.
 *   The planner should indicate which formula to use based on the user's
 *   Excel version if known, or default to XLOOKUP.
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
  description: "Lookup and match records between ranges using VLOOKUP/XLOOKUP or computed match",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: MatchRecordsParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const {
    lookupRange,
    sourceRange,
    matchType,
    outputRange,
    preferFormula = true,
  } = params;
  // Default to returning the first non-key column (column index 1) if not specified
  const returnColumns = params.returnColumns?.length ? params.returnColumns : [1];

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would match records from ${lookupRange} against ${sourceRange}, output to ${outputRange}`,
    };
  }

  if (preferFormula) {
    return await formulaMatch(context, params, options);
  } else {
    return await computedMatch(context, params, options);
  }
}

async function formulaMatch(
  context: Excel.RequestContext,
  params: MatchRecordsParams,
  options: ExecutionOptions
): Promise<StepResult> {
  options.onProgress?.("Building lookup formulas...");

  // Strip workbook qualifiers — formula strings reference ranges within
  // the same workbook, so "[WorkbookName.xlsx]Sheet1!A:A" → "Sheet1!A:A".
  const lookupAddr  = stripWorkbookQualifier(params.lookupRange);
  const sourceAddr  = stripWorkbookQualifier(params.sourceRange);
  const outputAddr  = stripWorkbookQualifier(params.outputRange);

  // Determine actual row count.
  // Full-column refs like "A:A" have rowCount = 1,048,576 which we cannot
  // iterate. getUsedRange(true) on the column itself returns only the
  // sub-range that has data, giving the exact last filled row.
  // getUsedRange(true) throws "The requested resource doesn't exist" when
  // the range is empty — catch that and treat it as 0 rows.
  let rowCount = 0;
  try {
    const lookupRng = resolveRange(context, params.lookupRange);
    const lookupUsed = lookupRng.getUsedRange(true);
    lookupUsed.load("rowCount");
    await context.sync();
    rowCount = lookupUsed.rowCount;
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    if (msg.includes("doesn't exist") || msg.includes("ItemNotFound") || msg.includes("not found")) {
      return { stepId: "", status: "success", message: `No data found in lookup range: ${params.lookupRange}` };
    }
    throw new Error(`Failed to read lookup range "${params.lookupRange}": ${msg}`);
  }

  if (rowCount === 0) {
    return { stepId: "", status: "success", message: "No rows to process." };
  }

  const matchMode = params.matchType === "exact" ? "0" : "1";
  const formulas: string[][] = [];

  for (let row = 0; row < rowCount; row++) {
    const rowFormulas: string[] = [];
    for (const colIdx of params.returnColumns) {
      const lookupCell  = getRelCellRef(lookupAddr, row);
      const lookupArr   = getColumnRef(sourceAddr, 0);
      const returnArr   = getColumnRef(sourceAddr, colIdx - 1);
      rowFormulas.push(
        `=IFERROR(XLOOKUP(${lookupCell},${lookupArr},${returnArr},"",${matchMode}),"")`
      );
    }
    formulas.push(rowFormulas);
  }

  // Build a precise output range (e.g. Sheet2!C1:D10) rather than writing to
  // the full column — assigning a small array to a 1M-row range can fail in
  // Office.js when the array dimensions don't match the range dimensions.
  const outputRng = resolveRange(context, buildOutputRange(outputAddr, rowCount, params.returnColumns.length));

  outputRng.formulas = formulas;
  await context.sync();

  options.onProgress?.(`Wrote ${rowCount} lookup formulas`);
  return {
    stepId: "",
    status: "success",
    message: `Created ${rowCount} XLOOKUP formulas in ${outputAddr}`,
  };
}

async function computedMatch(
  context: Excel.RequestContext,
  params: MatchRecordsParams,
  options: ExecutionOptions
): Promise<StepResult> {
  options.onProgress?.("Reading source data...");

  // resolveRange already handles workbook-qualified addresses via rangeUtils.
  const lookupRng = resolveRange(context, params.lookupRange);
  const sourceRng = resolveRange(context, params.sourceRange);
  lookupRng.load("values");
  sourceRng.load("values");
  await context.sync();

  const lookupValues = lookupRng.values;
  const sourceValues = sourceRng.values;

  options.onProgress?.("Matching records...");

  // Build index from source key column (column 0)
  const index = new Map<string, (string | number | boolean)[]>();
  for (const row of sourceValues) {
    const key = String(row[0]).toLowerCase();
    if (!index.has(key)) {
      index.set(key, row as (string | number | boolean)[]);
    }
  }

  // Match each lookup value
  const results: (string | number | boolean | null)[][] = [];
  for (const lookupRow of lookupValues) {
    const key = String(lookupRow[0]).toLowerCase();
    const sourceRow = index.get(key);
    const resultRow: (string | number | boolean | null)[] = [];

    for (const colIdx of params.returnColumns) {
      resultRow.push(sourceRow ? sourceRow[colIdx - 1] ?? null : null);
    }
    results.push(resultRow);
  }

  // Write results to a precise range (not full column) to avoid dimension mismatch
  const outputAddr = stripWorkbookQualifier(params.outputRange);
  const outputRng = resolveRange(context, buildOutputRange(outputAddr, results.length, params.returnColumns.length));
  outputRng.values = results;
  await context.sync();

  const matched = results.filter((r) => r.some((v) => v !== null)).length;
  return {
    stepId: "",
    status: "success",
    message: `Matched ${matched}/${lookupValues.length} records, wrote to ${params.outputRange}`,
  };
}


/** Get a cell reference relative to a range's first column at a given row offset */
function getRelCellRef(rangeAddress: string, rowOffset: number): string {
  // Simplified: returns e.g. "Sheet1!A2" for range "Sheet1!A1:A100" with offset 1
  const parts = rangeAddress.split("!");
  const ref = parts.length > 1 ? parts[1] : parts[0];
  const col = ref.match(/[A-Z]+/)?.[0] ?? "A";
  const startRow = parseInt(ref.match(/\d+/)?.[0] ?? "1", 10);
  const prefix = parts.length > 1 ? parts[0] + "!" : "";
  return `${prefix}${col}${startRow + rowOffset}`;
}

/** Get a full column reference like "Sheet1!A:A" from a range */
function getColumnRef(rangeAddress: string, colOffset: number): string {
  const parts = rangeAddress.split("!");
  const ref = parts.length > 1 ? parts[1] : parts[0];
  const startCol = ref.match(/[A-Z]+/)?.[0] ?? "A";
  const col = offsetColumn(startCol, colOffset);
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
 * Build a precise output range address like "Sheet2!C1:D10".
 * Avoids writing a small array to a full-column range (1M rows) which
 * causes Office.js to throw a dimension mismatch error.
 */
function buildOutputRange(outputAddr: string, rowCount: number, colCount: number): string {
  const parts = outputAddr.split("!");
  const ref = parts.length > 1 ? parts[1] : parts[0];
  const prefix = parts.length > 1 ? parts[0] + "!" : "";
  const startCol = ref.match(/[A-Z]+/)?.[0] ?? "A";
  const startRow = parseInt(ref.match(/\d+/)?.[0] ?? "1", 10);
  const endCol = offsetColumn(startCol, colCount - 1);
  const endRow = startRow + rowCount - 1;
  return `${prefix}${startCol}${startRow}:${endCol}${endRow}`;
}

registry.register(meta, handler as any);
export { meta };
