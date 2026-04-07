/**
 * topN – Extract the top N or bottom N rows from a range sorted by a value column.
 *
 * Reads the full data range, sorts by the specified value column, takes the
 * first N rows, and writes them to the output range.
 */

import { CapabilityMeta, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "topN",
  description: "Extract the top N or bottom N rows from a range sorted by a value column",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: any,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { dataRange, valueColumn, n, position = "top", outputRange, hasHeaders = true } = params;

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would extract ${position} ${n} rows from ${dataRange} by column ${valueColumn} to ${outputRange}` };
  }

  options.onProgress?.("Reading data...");

  const rng = resolveRange(context, dataRange);
  rng.load("values");
  await context.sync();

  const vals = (rng.values ?? []) as (string | number | boolean | null)[][];
  if (vals.length < 2) {
    return { stepId: "", status: "success", message: "Not enough rows." };
  }

  // Separate header and data rows
  const headerRow = hasHeaders ? vals[0] : null;
  const dataRows = hasHeaders ? vals.slice(1) : vals.slice();
  const colIdx = valueColumn - 1;

  options.onProgress?.(`Sorting by column ${valueColumn} (${position})...`);

  // Sort data rows
  dataRows.sort((a, b) => {
    const aVal = Number(a[colIdx]) || 0;
    const bVal = Number(b[colIdx]) || 0;
    return position === "top" ? bVal - aVal : aVal - bVal;
  });

  // Take first N rows
  const taken = dataRows.slice(0, n);

  // Build output
  const output: (string | number | boolean | null)[][] = [];
  if (headerRow) {
    output.push(headerRow);
  }
  output.push(...taken);

  options.onProgress?.(`Writing ${taken.length} rows to ${outputRange}...`);

  const outRng = resolveRange(context, outputRange);
  outRng.getResizedRange(output.length - 1, (output[0]?.length ?? 1) - 1).values = output as any;
  await context.sync();

  const label = position === "top" ? "Top" : "Bottom";
  return {
    stepId: "",
    status: "success",
    message: `${label} ${taken.length} rows by column ${valueColumn} written to ${outputRange}`,
  };
}

registry.register(meta, handler as any);
export { meta };
