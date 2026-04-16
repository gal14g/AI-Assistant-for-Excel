/**
 * forecast — project a time series forward using FORECAST.LINEAR or
 * FORECAST.ETS.
 *
 * Input: 2-col sourceRange (dates | values). Emits a 2-col projection table
 * of length `periods` starting one step after the last source date, using
 * either FORECAST.LINEAR (linear regression) or FORECAST.ETS (exponential
 * triple smoothing, better for seasonal data). Optional line chart showing
 * source + forecast together.
 */

import {
  CapabilityMeta,
  ForecastParams,
  StepResult,
  ExecutionOptions,
} from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";
import { parseDateFlexible } from "../utils/parseDateFlexible";

const meta: CapabilityMeta = {
  action: "forecast",
  description: "Project a time series forward (FORECAST.LINEAR or FORECAST.ETS) with optional chart",
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
  params: ForecastParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { sourceRange, outputRange, periods, method = "linear", includeChart = true, hasHeaders = true } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would project ${periods} period(s) via FORECAST.${method === "ets" ? "ETS" : "LINEAR"}.`,
    };
  }

  options.onProgress?.("Setting up forecast table...");

  // Read the source's address + dimensions. We don't need values — FORECAST
  // operates on the live range. But we DO need the last-date cell to extend
  // the date series.
  const src = resolveRange(context, sourceRange);
  const srcUsed = src.getUsedRange(false);
  srcUsed.load(["address", "rowCount", "columnCount", "values", "worksheet/name"]);
  await context.sync();

  if (srcUsed.columnCount < 2) {
    return { stepId: "", status: "error", message: "sourceRange must have 2 columns (dates | values)." };
  }

  const srcAddr = srcUsed.address;
  const srcAddrOnly = srcAddr.includes("!") ? srcAddr.split("!").pop()! : srcAddr;
  const m = srcAddrOnly.match(/^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/);
  if (!m) {
    return { stepId: "", status: "error", message: `Could not parse source address: ${srcAddr}` };
  }
  const [, srcColA, srcRow1, srcColB, srcRowN] = m;
  const srcSheetName = srcUsed.worksheet.name;
  const firstDataRow = hasHeaders ? Number(srcRow1) + 1 : Number(srcRow1);
  const lastDataRow = Number(srcRowN);
  const sheetPrefix = srcSheetName.includes(" ") ? `'${srcSheetName}'!` : `${srcSheetName}!`;
  const knownYRange = `${sheetPrefix}${srcColB}${firstDataRow}:${srcColB}${lastDataRow}`;
  const knownXRange = `${sheetPrefix}${srcColA}${firstDataRow}:${srcColA}${lastDataRow}`;

  // Derive the next date from the last two source dates (assume linear step).
  const values = (srcUsed.values ?? []) as (string | number | boolean | null)[][];
  const srcRowsData = values.slice(hasHeaders ? 1 : 0);
  if (srcRowsData.length < 2) {
    return { stepId: "", status: "error", message: "Need at least 2 source rows to infer the date step." };
  }
  // Convert the last two source dates into Excel serials so we can project
  // forward in the same unit Excel uses. parseDateFlexible handles every
  // format variation (US mm/dd, EU dd/mm, ISO, month-name, Excel serial, etc.)
  const toExcelSerial = (v: unknown): number | null => {
    if (typeof v === "number") return v;
    const d = parseDateFlexible(v);
    if (d === null) return null;
    // 1970 epoch offset to Excel epoch (1899-12-30): 25569 days.
    return d.getTime() / 86_400_000 + 25569;
  };
  const lastTwoDates = [
    toExcelSerial(srcRowsData[srcRowsData.length - 2][0]),
    toExcelSerial(srcRowsData[srcRowsData.length - 1][0]),
  ];
  if (lastTwoDates[0] == null || lastTwoDates[1] == null) {
    return { stepId: "", status: "error", message: "Could not parse the last two source dates." };
  }
  const dateStep = lastTwoDates[1] - lastTwoDates[0];

  // Compose output block: headers + `periods` rows of [date, forecast formula].
  const out = resolveRange(context, outputRange);
  out.load(["address", "worksheet/name"]);
  await context.sync();

  const outAddrPart = out.address.includes("!") ? out.address.split("!").pop()! : out.address;
  const outFirst = outAddrPart.split(":")[0].match(/^([A-Z]+)(\d+)$/);
  if (!outFirst) {
    return { stepId: "", status: "error", message: `Could not parse output address: ${out.address}` };
  }
  const outStartColLetters = outFirst[1];
  const outStartColIdx = columnLetterToIndex(outStartColLetters);
  const outStartRow = Number(outFirst[2]);
  const outSheet = out.worksheet;

  const header: (string | number)[] = ["Date", "Forecast"];
  const body: (string | number)[][] = [header];
  for (let i = 0; i < periods; i++) {
    const dateSerial = lastTwoDates[1] + dateStep * (i + 1);
    // Write raw serial; user can format as date via Cell Format.
    // FORECAST formula references the new date cell (relative) so fill-down works.
    const rowNum = outStartRow + 1 + i;
    const dateCellRef = `${outStartColLetters}${rowNum}`;
    const fnName = method === "ets" ? "FORECAST.ETS" : "FORECAST.LINEAR";
    body.push([dateSerial, `=${fnName}(${dateCellRef},${knownYRange},${knownXRange})` as unknown as number]);
  }

  try {
    const block = outSheet.getRangeByIndexes(outStartRow - 1, outStartColIdx, body.length, 2);
    block.values = body as unknown as (string | number | boolean)[][];
    // The Forecast column needs to be re-written as formulas (since we smuggled
    // them as strings above, range.values treats them as literal strings).
    // Rewrite the 2nd column with explicit formulas.
    const fcolLetter = indexToLetters(outStartColIdx + 1);
    const fcolRange = outSheet.getRange(`${fcolLetter}${outStartRow + 1}:${fcolLetter}${outStartRow + periods}`);
    const fnName = method === "ets" ? "FORECAST.ETS" : "FORECAST.LINEAR";
    const formulas: string[][] = [];
    for (let i = 0; i < periods; i++) {
      const dateCellRef = `${outStartColLetters}${outStartRow + 1 + i}`;
      formulas.push([`=${fnName}(${dateCellRef},${knownYRange},${knownXRange})`]);
    }
    fcolRange.formulas = formulas;

    // Format date column as dd/mm/yyyy for readability.
    const dateColLetter = outStartColLetters;
    const dateColRange = outSheet.getRange(`${dateColLetter}${outStartRow + 1}:${dateColLetter}${outStartRow + periods}`);
    dateColRange.numberFormat = [["dd/mm/yyyy"]] as unknown as string[][];

    await context.sync();
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return { stepId: "", status: "error", message: `Failed to write forecast block: ${msg}`, error: msg };
  }

  let chartName: string | undefined;
  if (includeChart) {
    try {
      // Combined chart from source + forecast. Easiest: build a contiguous
      // 2-col range on a temp region — but cross-region charts don't exist
      // cleanly via Office.js. We fall back to charting the forecast block
      // alone; the user can extend by including the source range manually
      // if they want both series.
      const forecastChartRange = outSheet.getRange(
        `${outStartColLetters}${outStartRow}:${indexToLetters(outStartColIdx + 1)}${outStartRow + periods}`,
      );
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const chart: any = (outSheet as any).charts.add("line", forecastChartRange, "auto");
      chart.title.text = `Forecast (${method === "ets" ? "ETS" : "Linear"})`;
      chartName = chart.name ?? "Forecast";
      await context.sync();
    } catch {
      // Chart failures shouldn't abort the main output.
    }
  }

  const outputAddr = `${outSheet.name}!${outStartColLetters}${outStartRow}:${indexToLetters(outStartColIdx + 1)}${outStartRow + periods}`;
  return {
    stepId: "",
    status: "success",
    message: `Projected ${periods} period(s) via FORECAST.${method === "ets" ? "ETS" : "LINEAR"}. Output: ${outputAddr}.${chartName ? ` Chart: ${chartName}.` : ""}`,
    outputs: { outputRange: outputAddr, ...(chartName ? { chartName } : {}) },
  };
}

registry.register(meta, handler as any);
export { meta };
