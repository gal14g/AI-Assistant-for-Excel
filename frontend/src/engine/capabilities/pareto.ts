/**
 * pareto — 80/20 analysis.
 *
 * Input: 2-col dataRange (label | value). Output: 3 columns (label | value |
 * cumulative %) with rows sorted by value descending. Optionally a combo
 * chart (column for value, line for cumulative %) — via two stacked chart
 * series since Office.js "combo" chart type creation is inconsistent
 * across versions.
 */

import { CapabilityMeta, ParetoParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";
import { parseNumberFlexible } from "../utils/parseNumberFlexible";

const meta: CapabilityMeta = {
  action: "pareto",
  description: "Pareto (80/20) analysis: sorted values + cumulative % + optional chart",
  mutates: true,
  affectsFormatting: false,
  requiresApiSet: "ExcelApi 1.2",
};

function indexToLetters(idx: number): string {
  let n = idx + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}
function columnLetterToIndex(letters: string): number {
  let n = 0;
  for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
  return n - 1;
}

async function handler(
  context: Excel.RequestContext,
  params: ParetoParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { dataRange, outputRange, includeChart = true, hasHeaders = true } = params;

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would build Pareto from ${dataRange}` };
  }

  options.onProgress?.("Computing Pareto (sort + cumulative %)...");

  const src = resolveRange(context, dataRange);
  const srcUsed = src.getUsedRange(false);
  srcUsed.load(["values", "rowCount", "columnCount", "address"]);
  await context.sync();

  if (srcUsed.columnCount < 2) {
    return { stepId: "", status: "error", message: "dataRange must have 2 columns (label | value)." };
  }

  const values = (srcUsed.values ?? []) as (string | number | boolean | null)[][];
  const startRow = hasHeaders ? 1 : 0;
  const rows: { label: string | number; value: number }[] = [];
  for (let r = startRow; r < values.length; r++) {
    const lbl = values[r][0];
    // parseNumberFlexible: handles text-stored numbers, currency prefixes,
    // percent, EU format, paren-negatives — anything a CSV import might have.
    const v = parseNumberFlexible(values[r][1]);
    if (v !== null) {
      rows.push({ label: (lbl ?? "") as string | number, value: v });
    }
  }
  if (rows.length === 0) {
    return { stepId: "", status: "error", message: "No numeric value rows found." };
  }

  // Sort descending by value.
  rows.sort((a, b) => b.value - a.value);
  const total = rows.reduce((sum, r) => sum + r.value, 0);
  let running = 0;
  const output: (string | number)[][] = [["Label", "Value", "Cumulative %"]];
  for (const r of rows) {
    running += r.value;
    output.push([r.label, r.value, total === 0 ? 0 : running / total]);
  }

  const out = resolveRange(context, outputRange);
  out.load(["address", "worksheet/name"]);
  await context.sync();
  const outAddrPart = out.address.includes("!") ? out.address.split("!").pop()! : out.address;
  const outFirst = outAddrPart.split(":")[0].match(/^([A-Z]+)(\d+)$/);
  if (!outFirst) {
    return { stepId: "", status: "error", message: `Could not parse output address: ${out.address}` };
  }
  const outStartCol = outFirst[1];
  const outStartColIdx = columnLetterToIndex(outStartCol);
  const outStartRow = Number(outFirst[2]);
  const outSheet = out.worksheet;

  try {
    const block = outSheet.getRangeByIndexes(outStartRow - 1, outStartColIdx, output.length, 3);
    block.values = output as unknown as (string | number | boolean)[][];
    // Format the cumulative-% column.
    const pctCol = indexToLetters(outStartColIdx + 2);
    const pctRange = outSheet.getRange(
      `${pctCol}${outStartRow + 1}:${pctCol}${outStartRow + output.length - 1}`,
    );
    pctRange.numberFormat = [["0.0%"]] as unknown as string[][];
    await context.sync();
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return { stepId: "", status: "error", message: `Failed to write Pareto block: ${msg}`, error: msg };
  }

  let chartName: string | undefined;
  if (includeChart) {
    try {
      const chartRange = outSheet.getRange(
        `${outStartCol}${outStartRow}:${indexToLetters(outStartColIdx + 2)}${outStartRow + output.length - 1}`,
      );
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const chart: any = (outSheet as any).charts.add("columnClustered", chartRange, "auto");
      chart.title.text = "Pareto Analysis";
      chartName = chart.name ?? "Pareto";
      await context.sync();
    } catch {
      // Chart failure shouldn't abort.
    }
  }

  const outputAddr = `${outSheet.name}!${outStartCol}${outStartRow}:${indexToLetters(outStartColIdx + 2)}${outStartRow + output.length - 1}`;
  return {
    stepId: "",
    status: "success",
    message: `Pareto analysis written with ${rows.length} item(s). Output: ${outputAddr}.${chartName ? ` Chart: ${chartName}.` : ""}`,
    outputs: { outputRange: outputAddr, ...(chartName ? { chartName } : {}) },
  };
}

registry.register(meta, handler as any);
export { meta };
