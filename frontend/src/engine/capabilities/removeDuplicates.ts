/**
 * removeDuplicates – Remove duplicate rows from a range.
 *
 * Office.js notes:
 * - Range.removeDuplicates() is available in ExcelApi 1.9+.
 * - columnIndexes specifies which columns to compare (0-based).
 * - If omitted, all columns are compared.
 * - This modifies the range in-place; rows are deleted.
 */

import { CapabilityMeta, RemoveDuplicatesParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";
import { ensureUnmerged } from "../utils/mergedCells";

const meta: CapabilityMeta = {
  action: "removeDuplicates",
  description: "Remove duplicate rows from a range",
  mutates: true,
  affectsFormatting: false,
  requiresApiSet: "ExcelApi 1.9",
};

async function handler(
  context: Excel.RequestContext,
  params: RemoveDuplicatesParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const { range: address, columnIndexes } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would remove duplicates from ${address}`,
    };
  }

  options.onProgress?.("Removing duplicates...");

  // Note: Office.js Range.removeDuplicates() is efficient on its own — it only
  // processes rows that have data regardless of the range address. No need to
  // call getUsedRange() here (and doing so without syncing inside the try block
  // would fail silently since Office.js proxy errors surface at context.sync time).
  const range = resolveRange(context, address);

  // Office.js removeDuplicates() errors on ranges with merged cells.
  // Auto-unmerge so the dedupe can proceed.
  const mergeReport = await ensureUnmerged(context, range, {
    operation: "removeDuplicates",
    policy: "refuseWithError",
  });
  if (mergeReport.error) return mergeReport.error;

  const result = range.removeDuplicates(
    columnIndexes ?? [],
    true // includesHeader – default to true
  );
  result.load(["removed", "uniqueRemaining"]);
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Removed ${result.removed} duplicate rows; ${result.uniqueRemaining} unique rows remain${mergeReport.warning ?? ""}`,
    outputs: { range: address, removedCount: result.removed },
  };
}


// ── Legacy-Excel fallback (ExcelApi < 1.9) ────────────────────────────────────
// Range.removeDuplicates requires 1.9. We reproduce the behavior in JS: read
// the range values, dedupe rows by the chosen column indexes (or all columns),
// write the result back, and clear any leftover rows at the bottom of the
// original range. Headers are always preserved.
async function fallback(
  context: Excel.RequestContext,
  params: RemoveDuplicatesParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { range: address, columnIndexes } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would dedupe ${address} in JS (legacy fallback).`,
    };
  }

  options.onProgress?.("Legacy-Excel mode: deduping in JS (Range.removeDuplicates unavailable)...");

  // Clip via getUsedRange(false) so "A:A" or "Sheet1!A:D" doesn't load ~1M
  // empty rows on Excel 2016. All subsequent reads/writes go against the
  // bounded range, mirroring Office.js's own internal clipping on 1.9+.
  const rawRange = resolveRange(context, address);
  const range = rawRange.getUsedRange(false);
  range.load(["values", "rowCount", "columnCount", "address"]);
  await context.sync();

  const values = (range.values ?? []) as (string | number | boolean | null)[][];
  if (values.length === 0) {
    return { stepId: "", status: "success", message: "Range is empty — nothing to dedupe." };
  }

  // Always treat row 0 as the header and dedupe among rows [1..].
  const header = values[0];
  const dataRows = values.slice(1);

  // Decide which columns form the dedupe key. If no indexes given, compare all columns.
  const keyCols = (columnIndexes && columnIndexes.length > 0)
    ? columnIndexes
    : header.map((_, i) => i);

  const seen = new Set<string>();
  const uniqueRows: (string | number | boolean | null)[][] = [];
  let removed = 0;
  for (const row of dataRows) {
    const key = keyCols.map((i) => JSON.stringify(row[i] ?? null)).join("\u0001");
    if (seen.has(key)) {
      removed += 1;
    } else {
      seen.add(key);
      uniqueRows.push(row);
    }
  }

  const colCount = range.columnCount;
  const rowCount = range.rowCount;

  // Write deduped rows back. Pad any missing rows with nulls to overwrite
  // the previous content of the range.
  const out: (string | number | boolean | null)[][] = [header];
  for (const r of uniqueRows) out.push(r);
  while (out.length < rowCount) {
    const blank: (string | number | boolean | null)[] = [];
    for (let c = 0; c < colCount; c++) blank.push(null);
    out.push(blank);
  }
  // Range.values doesn't accept `null` per the TS types, but Office.js treats
  // null as "empty cell" at runtime. Cast through unknown so the types let us.
  range.values = out as unknown as (string | number | boolean)[][];
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message:
      `Removed ${removed} duplicate rows in JS; ${uniqueRows.length} unique rows remain ` +
      `(legacy-Excel fallback — Range.removeDuplicates requires ExcelApi 1.9+).`,
    outputs: { range: address, removedCount: removed },
  };
}

registry.register(meta, handler as any, fallback as any);
export { meta };
