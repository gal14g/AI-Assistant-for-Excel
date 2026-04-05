/**
 * compareSheets – Diff two ranges and report/highlight differences.
 *
 * Two modes:
 *   highlightDiffs = true  → highlights differing cells in rangeA in-place
 *   highlightDiffs = false → writes a diff report to outputRange (or new sheet)
 *
 * The report has columns: Row | Col | Value A | Value B
 */

import { CapabilityMeta, CompareSheetsParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "compareSheets",
  description: "Compare two ranges and highlight or report differences",
  mutates: true,
  affectsFormatting: true,
};

async function handler(
  context: Excel.RequestContext,
  params: CompareSheetsParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { rangeA, rangeB, highlightDiffs = false, highlightColor = "#FFD966" } = params;

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would compare ${rangeA} vs ${rangeB}` };
  }

  options.onProgress?.("Reading both ranges...");
  const rngA = resolveRange(context, rangeA);
  const rngB = resolveRange(context, rangeB);
  rngA.load("values, address");
  rngB.load("values");
  await context.sync();

  const valsA = (rngA.values ?? []) as (string | number | boolean | null)[][];
  const valsB = (rngB.values ?? []) as (string | number | boolean | null)[][];

  const rows = Math.max(valsA.length, valsB.length);
  const cols = Math.max(valsA[0]?.length ?? 0, valsB[0]?.length ?? 0);

  options.onProgress?.(`Comparing ${rows} rows × ${cols} columns...`);

  type Diff = { row: number; col: number; valA: string; valB: string };
  const diffs: Diff[] = [];

  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const a = String(valsA[r]?.[c] ?? "");
      const b = String(valsB[r]?.[c] ?? "");
      if (a !== b) diffs.push({ row: r + 1, col: c + 1, valA: a, valB: b });
    }
  }

  if (highlightDiffs) {
    // Highlight differing cells in rangeA
    for (const d of diffs) {
      const cell = rngA.getCell(d.row - 1, d.col - 1);
      cell.format.fill.color = highlightColor;
    }
    await context.sync();
    return {
      stepId: "",
      status: "success",
      message: `Highlighted ${diffs.length} difference(s) in ${rangeA}`,
    };
  }

  // Write diff report
  const cellPart = rngA.address.includes("!") ? rngA.address.split("!").pop()! : rngA.address;
  const startRow = parseInt((cellPart.replace(/\$/g, "").match(/[A-Z]+(\d+)/) ?? ["", "1"])[1], 10);
  const startCol = (cellPart.replace(/\$/g, "").match(/([A-Z]+)\d+/) ?? ["", "A"])[1];

  let outRng: Excel.Range;
  if (params.outputRange) {
    outRng = resolveRange(context, params.outputRange);
  } else {
    // Write to a new sheet named "Diff_Report"
    const newSheet = context.workbook.worksheets.add("Diff_Report");
    outRng = newSheet.getRange("A1");
  }

  const reportRows: (string | number)[][] = [["Row", "Column", rangeA, rangeB]];
  for (const d of diffs) {
    reportRows.push([startRow + d.row - 1, startCol + (d.col - 1), d.valA, d.valB]);
  }

  if (diffs.length === 0) reportRows.push(["No differences found", "", "", ""]);
  outRng.getResizedRange(reportRows.length - 1, 3).values = reportRows as any;
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Found ${diffs.length} difference(s) between ${rangeA} and ${rangeB}`,
  };
}

registry.register(meta, handler as any);
export { meta };
