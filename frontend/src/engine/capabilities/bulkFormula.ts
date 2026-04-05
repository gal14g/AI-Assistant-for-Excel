/**
 * bulkFormula – Write a formula to an entire column, stopping at the last data row.
 *
 * The formula template should reference the first data row (e.g. "=A2*B2").
 * The executor auto-increments row numbers for each subsequent row.
 *
 * Example:
 *   formula:     "=A2*B2"
 *   dataRange:   "A:B"          ← used to detect last row
 *   outputRange: "C:C"          ← where formulas are written
 *   hasHeaders:  true           ← skip row 1
 */

import { CapabilityMeta, BulkFormulaParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange, resolveSheet, stripWorkbookQualifier } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "bulkFormula",
  description: "Write a formula to an entire column up to the last data row",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: BulkFormulaParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { formula, outputRange, dataRange, hasHeaders = true } = params;

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would write formula "${formula}" to ${outputRange} based on data in ${dataRange}` };
  }

  options.onProgress?.("Detecting last data row...");

  // Get the last used row from dataRange
  const ws = resolveSheet(context, dataRange);
  const dataRng = resolveRange(context, dataRange);
  let lastRow = 1;
  try {
    const used = dataRng.getUsedRange(false);
    used.load("rowCount, address");
    await context.sync();
    const cellPart = used.address.includes("!") ? used.address.split("!").pop()! : used.address;
    const m = cellPart.replace(/\$/g, "").match(/[A-Z]+(\d+)/);
    const startRow = m ? parseInt(m[1], 10) : 1;
    lastRow = startRow + used.rowCount - 1;
  } catch {
    ws.getUsedRange(false).load("rowCount");
    await context.sync();
    lastRow = ws.getUsedRange(false).rowCount;
  }

  const firstDataRow = hasHeaders ? 2 : 1;
  if (lastRow < firstDataRow) {
    return { stepId: "", status: "success", message: "No data rows found." };
  }
  const rowCount = lastRow - firstDataRow + 1;
  options.onProgress?.(`Writing formula to ${rowCount} rows...`);

  // Determine output column and start row
  const stripped = stripWorkbookQualifier(outputRange);
  const ref = stripped.includes("!") ? stripped.split("!")[1] : stripped;
  const outCol = ref.match(/[A-Z]+/)?.[0] ?? "C";

  // Build formula array, incrementing row numbers from the template
  // Template: "=A2*B2" → for row 3: "=A3*B3"
  const templateRow = formula.match(/\d+/)
    ? parseInt(formula.match(/\d+/)![0], 10)
    : firstDataRow;

  const formulas: string[][] = [];
  for (let r = firstDataRow; r <= lastRow; r++) {
    const offset = r - templateRow;
    const f = offset === 0
      ? formula
      : formula.replace(/([A-Z]+)(\d+)/g, (_, col, row) => `${col}${parseInt(row, 10) + offset}`);
    formulas.push([f]);
  }

  const outRng = ws.getRange(`${outCol}${firstDataRow}:${outCol}${lastRow}`);
  outRng.formulas = formulas;
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Wrote "${formula}" (adjusted) to ${rowCount} rows in column ${outCol}`,
  };
}

registry.register(meta, handler as any);
export { meta };
