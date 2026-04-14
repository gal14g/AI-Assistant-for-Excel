/**
 * growthRate – Calculate period-over-period growth rate.
 *
 * For each row after the first data row, writes: =IFERROR((C3-C2)/C2,"")
 * The first data row gets an empty string (no prior period).
 * Optionally formats the output column as "0.0%".
 */

import { CapabilityMeta, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange, resolveSheet, stripWorkbookQualifier } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "growthRate",
  description: "Calculate period-over-period growth rate formulas",
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
    return { stepId: "", status: "success", message: `Would write growth rate formulas from ${sourceRange} to ${outputRange}` };
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
  options.onProgress?.(`Writing ${rowCount} growth rate formulas...`);

  // Determine source and output column letters
  const strippedSrc = stripWorkbookQualifier(sourceRange);
  const srcRef = strippedSrc.includes("!") ? strippedSrc.split("!")[1] : strippedSrc;
  const srcCol = srcRef.match(/[A-Z]+/)?.[0] ?? "A";

  const strippedOut = stripWorkbookQualifier(outputRange);
  const outRef = strippedOut.includes("!") ? strippedOut.split("!")[1] : strippedOut;
  const outCol = outRef.match(/[A-Z]+/)?.[0] ?? "B";

  // Write header if applicable
  if (hasHeaders) {
    ws.getRange(`${outCol}1`).values = [["Growth %"]];
  }

  // Build growth rate formulas
  const formulas: string[][] = [];
  // First data row: no prior period
  formulas.push([""]);
  // Subsequent rows: =IFERROR((Cn-Cn-1)/Cn-1,"")
  for (let r = firstDataRow + 1; r <= lastRow; r++) {
    formulas.push([`=IFERROR((${srcCol}${r}-${srcCol}${r - 1})/${srcCol}${r - 1},"")`]);
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
    message: `Created ${rowCount} growth rate formulas in ${outputRange}`,
    outputs: { outputRange },
  };
}

registry.register(meta, handler as any);
export { meta };
