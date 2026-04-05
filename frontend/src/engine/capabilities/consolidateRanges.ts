/**
 * consolidateRanges – Merge data from multiple ranges/sheets into one.
 *
 * Modes:
 *   vertical   → stack all source ranges on top of each other (default)
 *   horizontal → join side-by-side (requires same number of rows)
 *
 * Options:
 *   addSourceLabel  → prepend a column with the source range name
 *   deduplicate     → remove duplicate rows after consolidating
 */

import { CapabilityMeta, ConsolidateRangesParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "consolidateRanges",
  description: "Merge data from multiple ranges or sheets into one consolidated range",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: ConsolidateRangesParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { sourceRanges, outputRange, direction = "vertical", addSourceLabel = false, deduplicate = false } = params;

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would consolidate ${sourceRanges.length} ranges to ${outputRange}` };
  }

  if (!sourceRanges.length) return { stepId: "", status: "success", message: "No source ranges provided." };

  options.onProgress?.(`Reading ${sourceRanges.length} source ranges...`);

  // Read all ranges
  const loaded: Excel.Range[] = [];
  for (const addr of sourceRanges) {
    const rng = resolveRange(context, addr);
    rng.load("values");
    loaded.push(rng);
  }
  await context.sync();

  type Row = (string | number | boolean | null)[];
  const tables: Row[][] = loaded.map((r) => (r.values ?? []) as Row[]);

  options.onProgress?.("Consolidating...");

  if (direction === "vertical") {
    let combined: Row[] = [];
    for (let i = 0; i < tables.length; i++) {
      const table = tables[i];
      if (!table.length) continue;
      // Include header only from first source
      const start = i === 0 ? 0 : 1;
      for (let r = start; r < table.length; r++) {
        const row: Row = addSourceLabel ? [sourceRanges[i], ...table[r]] : [...table[r]];
        combined.push(row);
      }
    }

    if (deduplicate) {
      const seen = new Set<string>();
      combined = combined.filter((row) => {
        const key = JSON.stringify(row);
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
    }

    if (!combined.length) return { stepId: "", status: "success", message: "No data to consolidate." };
    const outRng = resolveRange(context, outputRange);
    outRng.getResizedRange(combined.length - 1, combined[0].length - 1).values = combined as any;
    await context.sync();

    return {
      stepId: "",
      status: "success",
      message: `Consolidated ${tables.length} ranges → ${combined.length} rows in ${outputRange}${deduplicate ? " (deduped)" : ""}`,
    };
  } else {
    // horizontal: join side by side
    const maxRows = Math.max(...tables.map((t) => t.length));
    const combined: Row[] = Array.from({ length: maxRows }, () => []);
    for (const table of tables) {
      for (let r = 0; r < maxRows; r++) {
        const row = table[r] ?? [];
        combined[r].push(...row);
      }
    }
    if (!combined.length) return { stepId: "", status: "success", message: "No data to consolidate." };
    const outRng = resolveRange(context, outputRange);
    outRng.getResizedRange(combined.length - 1, combined[0].length - 1).values = combined as any;
    await context.sync();

    return {
      stepId: "",
      status: "success",
      message: `Horizontally joined ${tables.length} ranges → ${combined.length} rows × ${combined[0].length} columns`,
    };
  }
}

registry.register(meta, handler as any);
export { meta };
