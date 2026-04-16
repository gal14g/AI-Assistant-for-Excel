/**
 * histogram — distribution of numeric values across bins.
 *
 * Writes a 2-col table (bin upper-bound | count) and, by default, a column
 * chart to visualize. Bins can be supplied explicitly or auto-computed via
 * Sturges' rule (ceil(log2(N) + 1)).
 */

import {
  CapabilityMeta,
  HistogramParams,
  StepResult,
  ExecutionOptions,
} from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";
import { parseNumberFlexible } from "../utils/parseNumberFlexible";

const meta: CapabilityMeta = {
  action: "histogram",
  description: "Build a histogram (FREQUENCY + bin table + optional column chart)",
  mutates: true,
  affectsFormatting: false,
  requiresApiSet: "ExcelApi 1.2",
};

async function handler(
  context: Excel.RequestContext,
  params: HistogramParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const {
    dataRange,
    outputRange,
    bins,
    binCount,
    includeChart = true,
    chartType = "columnClustered",
    hasHeaders = true,
  } = params;

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would build histogram for ${dataRange}` };
  }

  options.onProgress?.("Computing histogram bins...");

  // Read data to determine min/max for auto-binning.
  const dataRaw = resolveRange(context, dataRange);
  const dataUsed = dataRaw.getUsedRange(false);
  dataUsed.load(["values", "rowCount", "columnCount", "address"]);
  await context.sync();

  const values = (dataUsed.values ?? []) as (string | number | boolean | null)[][];
  const startRow = hasHeaders ? 1 : 0;
  const nums: number[] = [];
  for (let r = startRow; r < values.length; r++) {
    // Tolerate text-stored numbers ("1,234", "$100", "50%", "(100)") that
    // arrive from CSV imports — parseNumberFlexible returns null for
    // genuinely non-numeric cells and we skip those.
    const v = parseNumberFlexible(values[r][0]);
    if (v !== null) nums.push(v);
  }
  if (nums.length === 0) {
    return { stepId: "", status: "error", message: "No numeric values found in dataRange." };
  }

  // Compute bins.
  let binEdges: number[];
  if (bins && bins.length > 0) {
    binEdges = [...bins].sort((a, b) => a - b);
  } else {
    const n = nums.length;
    const count = binCount ?? Math.max(2, Math.ceil(Math.log2(n) + 1)); // Sturges
    const min = Math.min(...nums);
    const max = Math.max(...nums);
    const step = (max - min) / count;
    binEdges = [];
    for (let i = 1; i <= count; i++) binEdges.push(min + step * i);
  }

  // Prepare output block: headers + bin labels + FREQUENCY formula. FREQUENCY
  // is an array formula — the trick: enter it on a range 1 taller than the
  // bins list so the "more than last bin" count lands in the final cell.
  //
  // Layout:
  //   outputRange.topLeft   →  ["Bin", "Count"]
  //   next rows: [bin1, 0], [bin2, 0], ..., [binN, 0], [>binN, 0]
  // Then we write =FREQUENCY(dataRange, binEdgesRange) as an array into the
  // count column.
  const out = resolveRange(context, outputRange);
  out.load(["address", "worksheet/name"]);
  await context.sync();

  const outAddrPart = out.address.includes("!") ? out.address.split("!").pop()! : out.address;
  const outFirst = outAddrPart.split(":")[0].match(/^([A-Z]+)(\d+)$/);
  if (!outFirst) {
    return { stepId: "", status: "error", message: `Could not parse output address: ${out.address}` };
  }
  const outStartCol = outFirst[1];
  const outStartRow = Number(outFirst[2]);
  const outSheet = out.worksheet;

  // Build the 2D matrix: 1 header + binEdges.length + 1 overflow row.
  const rows: (string | number | null)[][] = [["Bin", "Count"]];
  for (const edge of binEdges) rows.push([edge, null]);
  rows.push([`>${binEdges[binEdges.length - 1]}`, null]);

  const outputBlock = outSheet.getRangeByIndexes(
    outStartRow - 1,
    columnLetterToIndex(outStartCol),
    rows.length,
    2,
  );
  try {
    outputBlock.values = rows as unknown as (string | number | boolean)[][];
    await context.sync();
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return { stepId: "", status: "error", message: `Failed to write histogram block: ${msg}`, error: msg };
  }

  // Write FREQUENCY array formula into the Count column, covering all rows
  // after the header.
  const nBins = binEdges.length + 1; // include the overflow row
  const countCol = nextColumnLetter(outStartCol);
  const countFirstRow = outStartRow + 1;
  const countLastRow = outStartRow + nBins;
  const binRefFirstRow = outStartRow + 1;
  const binRefLastRow = outStartRow + binEdges.length;
  const countRangeAddr = `${countCol}${countFirstRow}:${countCol}${countLastRow}`;
  const binRangeAddr = `${outStartCol}${binRefFirstRow}:${outStartCol}${binRefLastRow}`;
  // The source data range for FREQUENCY — use the actual data region (skip header if present).
  const dataAddr = dataUsed.address;
  const dataAddrOnly = dataAddr.includes("!") ? dataAddr.split("!").pop()! : dataAddr;
  const dataSheetPrefix = dataAddr.includes("!") ? dataAddr.split("!")[0] + "!" : "";
  // Build a bounded address for just the data rows (skip header row on column 1).
  const dataMatch = dataAddrOnly.match(/^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/);
  let dataFormulaRef = dataAddr;
  if (dataMatch) {
    const [, c1, r1, c2, r2] = dataMatch;
    const dataStart = hasHeaders ? Number(r1) + 1 : Number(r1);
    dataFormulaRef = `${dataSheetPrefix}${c1}${dataStart}:${c2}${r2}`;
  }
  const formula = `=FREQUENCY(${dataFormulaRef},${binRangeAddr})`;

  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const countRange = outSheet.getRange(countRangeAddr) as any;
    // setFormulaArray introduced on Range; prefer formulas = [[...]] on a
    // single-column range — modern Excel treats this as an array formula
    // when the expression is FREQUENCY.
    const singleColFormula: string[][] = [];
    for (let i = 0; i < nBins; i++) singleColFormula.push([formula]);
    countRange.formulas = singleColFormula;
    await context.sync();
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return { stepId: "", status: "error", message: `Failed to write FREQUENCY formula: ${msg}`, error: msg };
  }

  // Optional chart.
  let chartName: string | undefined;
  if (includeChart) {
    try {
      const chartDataRange = outSheet.getRange(`${outStartCol}${outStartRow}:${countCol}${countLastRow}`);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const chart: any = (outSheet as any).charts.add(chartType, chartDataRange, "auto");
      chart.title.text = "Histogram";
      chartName = chart.name ?? "Histogram";
      await context.sync();
    } catch {
      // Chart failures shouldn't abort the main output.
    }
  }

  const outputAddr = `${outSheet.name}!${outStartCol}${outStartRow}:${countCol}${countLastRow}`;
  return {
    stepId: "",
    status: "success",
    message: `Histogram built with ${binEdges.length} bins + overflow. Output: ${outputAddr}.${chartName ? ` Chart: ${chartName}.` : ""}`,
    outputs: { outputRange: outputAddr, ...(chartName ? { chartName } : {}) },
  };
}

/** Letter like "A" → 0; "AA" → 26. */
function columnLetterToIndex(letters: string): number {
  let n = 0;
  for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
  return n - 1;
}

/** "A" → "B", "Z" → "AA". */
function nextColumnLetter(letters: string): string {
  const idx = columnLetterToIndex(letters) + 1;
  let n = idx + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

registry.register(meta, handler as any);
export { meta };
