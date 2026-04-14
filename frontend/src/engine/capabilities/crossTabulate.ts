/**
 * crossTabulate – Build a cross-tab (contingency) matrix from raw data.
 *
 * Example: Count candidates per role per department
 *   Row field:    Department
 *   Column field: Role
 *   Value:        Candidate name (count)
 *
 * Output is a static matrix written to outputRange.
 */

import { CapabilityMeta, CrossTabulateParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "crossTabulate",
  description: "Build a cross-tab matrix counting or summing values across two dimensions",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: CrossTabulateParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { sourceRange, rowField, columnField, valueField, aggregation = "count", outputRange } = params;

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would cross-tabulate ${sourceRange} (row: col ${rowField}, col: col ${columnField})` };
  }

  options.onProgress?.("Reading source data...");
  const srcRng = resolveRange(context, sourceRange);
  srcRng.load("values");
  await context.sync();

  const data = (srcRng.values ?? []) as (string | number | boolean | null)[][];
  if (data.length < 2) return { stepId: "", status: "success", message: "Not enough data." };

  const rowIdx = rowField - 1;
  const colIdx = columnField - 1;
  const valIdx = valueField - 1;

  options.onProgress?.("Building cross-tab matrix...");

  // Collect unique row keys and column keys (in order of first appearance)
  const rowKeys: string[] = [];
  const colKeys: string[] = [];
  const rowSet = new Set<string>();
  const colSet = new Set<string>();

  for (let r = 1; r < data.length; r++) {
    const rk = String(data[r][rowIdx] ?? "");
    const ck = String(data[r][colIdx] ?? "");
    if (!rowSet.has(rk)) { rowSet.add(rk); rowKeys.push(rk); }
    if (!colSet.has(ck)) { colSet.add(ck); colKeys.push(ck); }
  }

  // Accumulate values: matrix[rowKey][colKey] = { sum, count }
  type Cell = { sum: number; count: number };
  const matrix: Map<string, Map<string, Cell>> = new Map();
  for (const rk of rowKeys) {
    const row = new Map<string, Cell>();
    for (const ck of colKeys) row.set(ck, { sum: 0, count: 0 });
    matrix.set(rk, row);
  }

  for (let r = 1; r < data.length; r++) {
    const rk = String(data[r][rowIdx] ?? "");
    const ck = String(data[r][colIdx] ?? "");
    const v = Number(data[r][valIdx]) || (aggregation === "count" ? 1 : 0);
    const cell = matrix.get(rk)?.get(ck);
    if (cell) { cell.sum += v; cell.count++; }
  }

  // Build output 2D array
  const outRows: (string | number)[][] = [];
  // Header row: blank top-left, then column keys
  outRows.push(["", ...colKeys, "Total"]);
  // Data rows
  for (const rk of rowKeys) {
    const row: (string | number)[] = [rk];
    let rowTotal = 0;
    for (const ck of colKeys) {
      const cell = matrix.get(rk)?.get(ck) ?? { sum: 0, count: 0 };
      const v = aggregation === "count" ? cell.count : aggregation === "sum" ? cell.sum : cell.count ? cell.sum / cell.count : 0;
      row.push(v);
      rowTotal += aggregation === "count" ? cell.count : cell.sum;
    }
    row.push(rowTotal);
    outRows.push(row);
  }
  // Total row
  const totalsRow: (string | number)[] = ["Total"];
  let grandTotal = 0;
  for (let ci = 0; ci < colKeys.length; ci++) {
    let colTotal = 0;
    for (const rk of rowKeys) {
      const cell = matrix.get(rk)?.get(colKeys[ci]) ?? { sum: 0, count: 0 };
      colTotal += aggregation === "count" ? cell.count : cell.sum;
    }
    totalsRow.push(colTotal);
    grandTotal += colTotal;
  }
  totalsRow.push(grandTotal);
  outRows.push(totalsRow);

  const outRng = resolveRange(context, outputRange);
  outRng.getResizedRange(outRows.length - 1, outRows[0].length - 1).values = outRows as any;
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Cross-tab: ${rowKeys.length} rows × ${colKeys.length} columns (${aggregation}) written to ${outputRange}`,
    outputs: { outputRange },
  };
}

registry.register(meta, handler as any);
export { meta };
