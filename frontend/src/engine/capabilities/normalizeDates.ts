/**
 * normalizeDates – Standardize date formats in a column to a consistent format.
 *
 * Handles multiple input formats:
 *   dd/mm/yyyy, mm/dd/yyyy, yyyy-mm-dd, d-MMM-yy, Excel serial numbers, etc.
 *
 * Excel serial numbers are converted using the epoch offset (25569 days from
 * 1900-01-01 to 1970-01-01) and 86400000 ms per day.
 */

import { CapabilityMeta, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "normalizeDates",
  description: "Standardize date formats in a column to a consistent output format",
  mutates: true,
  affectsFormatting: false,
};

/** Months abbreviated for d-MMM-yy parsing */
const MONTH_ABBR: Record<string, number> = {
  jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
  jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11,
};

const EXCEL_EPOCH = 25569;
const MS_PER_DAY = 86400000;

function tryParseDate(val: string | number | boolean | null): Date | null {
  if (val === null || val === "") return null;

  // Excel serial number
  if (typeof val === "number") {
    if (val > 1 && val < 2958466) {
      // Valid Excel date serial range
      return new Date((val - EXCEL_EPOCH) * MS_PER_DAY);
    }
    return null;
  }

  const s = String(val).trim();

  // yyyy-mm-dd or yyyy/mm/dd
  const isoMatch = s.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/);
  if (isoMatch) {
    return new Date(+isoMatch[1], +isoMatch[2] - 1, +isoMatch[3]);
  }

  // dd/mm/yyyy or dd-mm-yyyy
  const dmy = s.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$/);
  if (dmy) {
    const day = +dmy[1];
    const month = +dmy[2];
    // Heuristic: if first number > 12, it must be day
    if (day > 12) {
      return new Date(+dmy[3], month - 1, day);
    }
    // If second number > 12, it must be day (mm/dd/yyyy)
    if (month > 12) {
      return new Date(+dmy[3], day - 1, month);
    }
    // Ambiguous — assume dd/mm/yyyy
    return new Date(+dmy[3], month - 1, day);
  }

  // d-MMM-yy or d-MMM-yyyy (e.g. 5-Jan-23, 15-Mar-2023)
  const mmmMatch = s.match(/^(\d{1,2})[/-]([A-Za-z]{3})[/-](\d{2,4})$/);
  if (mmmMatch) {
    const mon = MONTH_ABBR[mmmMatch[2].toLowerCase()];
    if (mon !== undefined) {
      let year = +mmmMatch[3];
      if (year < 100) year += year < 50 ? 2000 : 1900;
      return new Date(year, mon, +mmmMatch[1]);
    }
  }

  // Fallback: built-in Date.parse
  const fallback = Date.parse(s);
  if (!isNaN(fallback)) return new Date(fallback);

  return null;
}

function formatDate(d: Date, fmt: string): string {
  const pad = (n: number) => n.toString().padStart(2, "0");
  const yyyy = d.getFullYear().toString();
  const mm = pad(d.getMonth() + 1);
  const dd = pad(d.getDate());
  return fmt
    .replace(/yyyy/gi, yyyy)
    .replace(/mm/g, mm)
    .replace(/dd/g, dd);
}

async function handler(
  context: Excel.RequestContext,
  params: any,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { range, outputFormat, inputFormat: _inputFormat } = params as {
    range: string;
    outputFormat: string;
    inputFormat?: string;
  };

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would normalize dates in ${range} to format "${outputFormat}"`,
    };
  }

  options.onProgress?.("Reading range...");
  const rawRng = resolveRange(context, range);
  // Clip to used range — full-column refs like "A:A" would load ~1M rows
  let rng: Excel.Range;
  try {
    rng = rawRng.getUsedRange(false);
  } catch {
    rng = rawRng;
  }
  rng.load("values");
  await context.sync();

  const vals = (rng.values ?? []) as (string | number | boolean | null)[][];
  if (!vals.length) return { stepId: "", status: "success", message: "No data found." };

  options.onProgress?.("Normalizing dates...");
  let normalized = 0;
  const total = vals.reduce((sum, row) => sum + row.length, 0);
  const out: (string | number | boolean | null)[][] = vals.map((r) => [...r]);

  for (let r = 0; r < out.length; r++) {
    for (let c = 0; c < out[r].length; c++) {
      const parsed = tryParseDate(out[r][c]);
      if (parsed) {
        out[r][c] = formatDate(parsed, outputFormat);
        normalized++;
      }
    }
  }

  rng.values = out as any;
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Normalized ${normalized}/${total} dates to "${outputFormat}"`,
    outputs: { range },
  };
}

registry.register(meta, handler as any);
export { meta };
