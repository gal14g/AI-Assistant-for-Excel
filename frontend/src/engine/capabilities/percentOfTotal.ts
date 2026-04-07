/**
 * percentOfTotal – Write a percentage-of-total formula for each row.
 *
 * For each data row, writes: =C2/SUM($C$2:$C$N)
 * Optionally formats the output column as "0.0%".
 */

import { CapabilityMeta, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange, resolveSheet, stripWorkbookQualifier } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "percentOfTotal",
  description: "Write a percentage-of-total formula for each row",
  mutates: true,
  affectsFormatting: true,
};

async function handler(
  context: Excel.RequestContext,
  params: any,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { sourceRange, outputRange, hasHeaders = true, formatAsPercent = true } = params;

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would write percent-of-total formulas from ${sourceRange} to ${outputRange}` };
  }

  options.onProgress?.("Detecting last data row...");

  // Get the last used row from sourceRange
  const ws = resolveSheet(context, sourceRange);
  const srcRng = resolveRange(context, sourceRange);
  let lastRow = 1;
  try {
    const used = srcRng.getUsedRange(false);
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
  options.onProgress?.(`Writing ${rowCount} percent-of-total formulas...`);

  // Determine source and output column letters
  const strippedSrc = stripWorkbookQualifier(sourceRange);
  const srcRef = strippedSrc.includes("!") ? strippedSrc.split("!")[1] : strippedSrc;
  const srcCol = srcRef.match(/[A-Z]+/)?.[0] ?? "A";

  const strippedOut = stripWorkbookQualifier(outputRange);
  const outRef = strippedOut.includes("!") ? strippedOut.split("!")[1] : strippedOut;
  const outCol = outRef.match(/[A-Z]+/)?.[0] ?? "B";

  // Write header if applicable
  if (hasHeaders) {
    ws.getRange(`${outCol}1`).values = [["% of Total"]];
  }

  // Build percent-of-total formulas
  const formulas: string[][] = [];
  for (let r = firstDataRow; r <= lastRow; r++) {
    formulas.push([`=${srcCol}${r}/SUM($${srcCol}$${firstDataRow}:$${srcCol}$${lastRow})`]);
  }

  const outRng = ws.getRange(`${outCol}${firstDataRow}:${outCol}${lastRow}`);
  outRng.formulas = formulas;

  // Apply percent format if requested
  if (formatAsPercent) {
    outRng.numberFormat = Array.from({ length: rowCount }, () => ["0.0%"]);
  }

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Created ${rowCount} percent-of-total formulas in ${outputRange}`,
  };
}

registry.register(meta, handler as any);
export { meta };
