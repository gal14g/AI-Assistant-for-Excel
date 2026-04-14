/**
 * coerceDataType – Convert values in a range from one type to another.
 *
 * Supported conversions:
 *   text → number   (strips currency symbols, commas, whitespace before parseFloat)
 *   text → date     (attempts Date.parse)
 *   number → text   (String(val))
 *   any → text/number/date
 */

import { CapabilityMeta, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "coerceDataType",
  description: "Convert values in a column from one type to another (text-to-number, text-to-date, number-to-text)",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: any,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { range, targetType, dateFormat, locale: _locale } = params as {
    range: string;
    targetType: "number" | "text" | "date";
    dateFormat?: string;
    locale?: string;
  };

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would convert values in ${range} to ${targetType}`,
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

  options.onProgress?.("Converting values...");
  let converted = 0;
  const total = vals.reduce((sum, row) => sum + row.length, 0);
  const out: (string | number | boolean | null)[][] = vals.map((r) => [...r]);

  for (let r = 0; r < out.length; r++) {
    for (let c = 0; c < out[r].length; c++) {
      const val = out[r][c];
      if (val === null || val === "") continue;

      if (targetType === "number") {
        // Strip currency symbols, commas, and whitespace before parsing
        const cleaned = String(val).replace(/[$€£¥,\s]/g, "");
        const num = parseFloat(cleaned);
        if (!isNaN(num)) {
          out[r][c] = num;
          converted++;
        }
      } else if (targetType === "text") {
        out[r][c] = String(val);
        converted++;
      } else if (targetType === "date") {
        const parsed = Date.parse(String(val));
        if (!isNaN(parsed)) {
          const d = new Date(parsed);
          // Default ISO format; honour dateFormat if provided
          const fmt = dateFormat ?? "yyyy-mm-dd";
          out[r][c] = formatDate(d, fmt);
          converted++;
        }
      }
    }
  }

  rng.values = out as any;
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Converted ${converted}/${total} cells to ${targetType}`,
    outputs: { range },
  };
}

/** Simple date formatter supporting common Excel-style tokens. */
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

registry.register(meta, handler as any);
export { meta };
