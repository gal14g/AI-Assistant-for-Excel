/**
 * rankColumn – Write RANK formulas for values in a column.
 *
 * For each row, writes: =RANK(C2,$C$2:$C$N, order)
 * where order = 0 for descending, 1 for ascending.
 *
 * If hasHeaders is true, skips row 1 and writes a "Rank" header.
 */

import { CapabilityMeta, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange, resolveSheet, stripWorkbookQualifier } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "rankColumn",
  description: "Write RANK formulas for values in a column",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: any,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { sourceRange, outputRange, order = "descending", hasHeaders = true } = params;

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would write ${order} rank formulas from ${sourceRange} to ${outputRange}` };
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
  options.onProgress?.(`Writing ${rowCount} rank formulas...`);

  // Determine source and output column letters
  const strippedSrc = stripWorkbookQualifier(sourceRange);
  const srcRef = strippedSrc.includes("!") ? strippedSrc.split("!")[1] : strippedSrc;
  const srcCol = srcRef.match(/[A-Z]+/)?.[0] ?? "A";

  const strippedOut = stripWorkbookQualifier(outputRange);
  const outRef = strippedOut.includes("!") ? strippedOut.split("!")[1] : strippedOut;
  const outCol = outRef.match(/[A-Z]+/)?.[0] ?? "B";

  const rankOrder = order === "ascending" ? 1 : 0;

  // Write header if applicable
  if (hasHeaders) {
    ws.getRange(`${outCol}1`).values = [["Rank"]];
  }

  // Build RANK formulas
  const formulas: string[][] = [];
  for (let r = firstDataRow; r <= lastRow; r++) {
    formulas.push([`=RANK(${srcCol}${r},$${srcCol}$${firstDataRow}:$${srcCol}$${lastRow},${rankOrder})`]);
  }

  const outRng = ws.getRange(`${outCol}${firstDataRow}:${outCol}${lastRow}`);
  outRng.formulas = formulas;
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Created ${rowCount} rank formulas in ${outputRange}`,
    outputs: { outputRange },
  };
}

registry.register(meta, handler as any);
export { meta };
