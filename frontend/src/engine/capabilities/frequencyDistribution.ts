/**
 * frequencyDistribution – Count occurrences of each unique value in a column
 * and write a frequency table.
 *
 * Output columns: Value, Count, Percent (optional).
 * Sorting by value or frequency, ascending or descending.
 */

import { CapabilityMeta, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "frequencyDistribution",
  description: "Count occurrences of each unique value and write a frequency table",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: any,
  options: ExecutionOptions,
): Promise<StepResult> {
  const {
    sourceRange,
    outputRange,
    sortBy = "frequency",
    ascending = false,
    includePercent = true,
  } = params as {
    sourceRange: string;
    outputRange: string;
    sortBy?: "value" | "frequency";
    ascending?: boolean;
    includePercent?: boolean;
  };

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would compute frequency distribution of ${sourceRange} → ${outputRange}`,
    };
  }

  options.onProgress?.("Reading source data...");
  const srcRng = resolveRange(context, sourceRange);
  srcRng.load("values");
  await context.sync();

  const vals = (srcRng.values ?? []) as (string | number | boolean | null)[][];
  if (!vals.length) return { stepId: "", status: "success", message: "No data found." };

  options.onProgress?.("Counting frequencies...");

  // Flatten to a single list of values
  const flat: (string | number | boolean)[] = [];
  for (const row of vals) {
    for (const cell of row) {
      if (cell !== null && cell !== "") flat.push(cell);
    }
  }

  // Count frequency (case-insensitive for strings)
  const counts = new Map<string, { display: string | number | boolean; count: number }>();
  for (const val of flat) {
    const key = typeof val === "string" ? val.toLowerCase() : String(val);
    if (!counts.has(key)) {
      counts.set(key, { display: val, count: 0 });
    }
    counts.get(key)!.count++;
  }

  // Convert to sorted array
  const entries = Array.from(counts.values());

  if (sortBy === "frequency") {
    entries.sort((a, b) => ascending ? a.count - b.count : b.count - a.count);
  } else {
    // Sort by value — numeric-aware
    entries.sort((a, b) => {
      const aStr = String(a.display);
      const bStr = String(b.display);
      const cmp = aStr.localeCompare(bStr, undefined, { numeric: true, sensitivity: "base" });
      return ascending ? cmp : -cmp;
    });
  }

  // Build output rows
  const totalCount = flat.length;
  const header: (string | number)[] = includePercent
    ? ["Value", "Count", "Percent"]
    : ["Value", "Count"];

  const dataRows: (string | number)[][] = entries.map((e) => {
    const row: (string | number)[] = [e.display as string | number, e.count];
    if (includePercent) {
      const pct = Math.round((e.count / totalCount) * 10000) / 100; // two decimals
      row.push(pct);
    }
    return row;
  });

  const output = [header, ...dataRows];

  options.onProgress?.("Writing frequency table...");
  const outRng = resolveRange(context, outputRange);
  outRng.getResizedRange(output.length - 1, output[0].length - 1).values = output as any;
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Found ${entries.length} unique values across ${totalCount} total entries`,
  };
}

registry.register(meta, handler as any);
export { meta };
