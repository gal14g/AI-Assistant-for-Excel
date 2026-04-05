/**
 * fillBlanks – Fill empty cells in a range.
 *
 * fillMode:
 *   "down"     → copy the value from the cell above (most common — handles merged-cell exports)
 *   "up"       → copy the value from the cell below
 *   "constant" → fill every blank with constantValue
 */

import { CapabilityMeta, FillBlanksParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "fillBlanks",
  description: "Fill empty cells downward (or upward, or with a constant value)",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: FillBlanksParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { range, fillMode = "down", constantValue = "" } = params;

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would fill blanks in ${range} (mode: ${fillMode})` };
  }

  options.onProgress?.("Reading range...");
  const rng = resolveRange(context, range);
  const used = (() => { try { return rng.getUsedRange(false); } catch { return rng; } })();
  used.load("values");
  await context.sync();

  const vals = (used.values ?? []) as (string | number | boolean | null)[][];
  if (!vals.length) return { stepId: "", status: "success", message: "No data found." };

  options.onProgress?.("Filling blanks...");
  const isEmpty = (v: string | number | boolean | null): boolean => v === null || v === "";

  let filled = 0;
  const out: (string | number | boolean | null)[][] = vals.map((r) => [...r]);

  if (fillMode === "down") {
    for (let c = 0; c < out[0].length; c++) {
      let last: string | number | boolean | null = null;
      for (let r = 0; r < out.length; r++) {
        if (!isEmpty(out[r][c])) { last = out[r][c]; }
        else if (last !== null) { out[r][c] = last; filled++; }
      }
    }
  } else if (fillMode === "up") {
    for (let c = 0; c < out[0].length; c++) {
      let next: string | number | boolean | null = null;
      for (let r = out.length - 1; r >= 0; r--) {
        if (!isEmpty(out[r][c])) { next = out[r][c]; }
        else if (next !== null) { out[r][c] = next; filled++; }
      }
    }
  } else {
    // constant
    for (let r = 0; r < out.length; r++)
      for (let c = 0; c < out[r].length; c++)
        if (isEmpty(out[r][c])) { out[r][c] = constantValue; filled++; }
  }

  used.values = out as any;
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Filled ${filled} blank cell(s) in ${range} (mode: ${fillMode})`,
  };
}

registry.register(meta, handler as any);
export { meta };
