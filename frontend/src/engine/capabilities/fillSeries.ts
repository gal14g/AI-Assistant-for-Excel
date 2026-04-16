/**
 * fillSeries — generate a series (number / date / weekday / repeat pattern).
 *
 * Replaces the common pattern where the LLM builds a manual 2D array of
 * sequential numbers or dates in writeValues. A dedicated primitive is safer
 * for long sequences and more discoverable.
 *
 * Single context.sync: build the series in JS, write back.
 */

import { CapabilityMeta, FillSeriesParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";
import { parseDateFlexible, formatDateDMY } from "../utils/parseDateFlexible";

const meta: CapabilityMeta = {
  action: "fillSeries",
  description: "Generate a number, date, weekday, or repeating-pattern series into a range",
  mutates: true,
  affectsFormatting: false,
};

function addDateStep(d: Date, step: number, unit: "day" | "week" | "month" | "year"): Date {
  const out = new Date(d);
  switch (unit) {
    case "day":   out.setUTCDate(out.getUTCDate() + step); break;
    case "week":  out.setUTCDate(out.getUTCDate() + step * 7); break;
    case "month": out.setUTCMonth(out.getUTCMonth() + step); break;
    case "year":  out.setUTCFullYear(out.getUTCFullYear() + step); break;
  }
  return out;
}

async function handler(
  context: Excel.RequestContext,
  params: FillSeriesParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const {
    range: address,
    seriesType,
    start = seriesType === "number" ? 1 : undefined,
    step = 1,
    pattern,
    dateUnit = "day",
    count,
    horizontal = false,
  } = params;

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would fillSeries ${seriesType} into ${address}` };
  }

  options.onProgress?.(`Generating ${seriesType} series...`);

  const range = resolveRange(context, address);
  range.load(["rowCount", "columnCount", "address"]);
  await context.sync();

  // How many cells we need to fill.
  const totalCells = (count !== undefined)
    ? count
    : horizontal
      ? range.columnCount
      : range.rowCount;

  const series: (string | number)[] = [];

  switch (seriesType) {
    case "number": {
      const s = typeof start === "number" ? start : Number(start ?? 1);
      for (let i = 0; i < totalCells; i++) series.push(s + i * step);
      break;
    }
    case "date": {
      const s = parseDateFlexible(start ?? new Date()) ?? new Date();
      for (let i = 0; i < totalCells; i++) {
        series.push(formatDateDMY(addDateStep(s, i * step, dateUnit)));
      }
      break;
    }
    case "weekday": {
      let d = parseDateFlexible(start ?? new Date()) ?? new Date();
      // Advance to first weekday if start lands on a weekend.
      while (d.getUTCDay() === 0 || d.getUTCDay() === 6) d = addDateStep(d, 1, "day");
      for (let i = 0; i < totalCells; i++) {
        series.push(formatDateDMY(d));
        do { d = addDateStep(d, step, "day"); } while (d.getUTCDay() === 0 || d.getUTCDay() === 6);
      }
      break;
    }
    case "repeatPattern": {
      if (!pattern || pattern.length === 0) {
        return { stepId: "", status: "error", message: "pattern is required for seriesType='repeatPattern'." };
      }
      for (let i = 0; i < totalCells; i++) {
        const v = pattern[i % pattern.length];
        series.push(typeof v === "boolean" ? (v ? "TRUE" : "FALSE") : (v as string | number));
      }
      break;
    }
  }

  // Shape into 2D for write.
  const grid: (string | number)[][] = horizontal
    ? [series]
    : series.map((v) => [v]);

  // Fit to the provided range dimensions, resizing if needed.
  const topLeft = range.getCell(0, 0);
  const rows = grid.length;
  const cols = grid[0]?.length ?? 1;
  const target = topLeft.getResizedRange(rows - 1, cols - 1);
  try {
    target.values = grid;
    await context.sync();
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return { stepId: "", status: "error", message: `Failed to write series: ${msg}`, error: msg };
  }

  return {
    stepId: "",
    status: "success",
    message: `Wrote ${totalCells} ${seriesType} value(s) to ${address}.`,
    outputs: { range: address, filledCount: totalCells },
  };
}

registry.register(meta, handler as any);
export { meta };
