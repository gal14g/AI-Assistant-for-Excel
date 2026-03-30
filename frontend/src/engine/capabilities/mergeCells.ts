/**
 * mergeCells – Merge cells in a range.
 *
 * Office.js merge modes:
 *   range.merge(false)  → merge all cells in the range into one
 *   range.merge(true)   → merge each row separately (merge across)
 *
 * The MergeCellsParams.across flag maps directly to this boolean.
 *
 * Note: To unmerge, use range.unmerge() — not yet exposed as a separate
 * action but can be added to this handler in the future.
 */

import { CapabilityMeta, MergeCellsParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "mergeCells",
  description: "Merge cells in a range (fully or row-by-row)",
  mutates: true,
  affectsFormatting: true,
};

async function handler(
  context: Excel.RequestContext,
  params: MergeCellsParams,
  options: ExecutionOptions
): Promise<StepResult> {
  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would merge cells in ${params.range}` };
  }

  options.onProgress?.(`Merging cells in ${params.range}...`);

  const range = resolveRange(context, params.range);
  // across=true merges each row individually; false (default) merges the whole range.
  // Also accept mergeType from the backend Pydantic model ("mergeAcross" → across=true).
  const across = params.across ?? (params as any).mergeType === "mergeAcross";
  range.merge(across);
  await context.sync();

  const modeLabel = across ? "merge across" : "full merge";
  return {
    stepId: "",
    status: "success",
    message: `Merged cells in ${params.range} (${modeLabel})`,
  };
}

registry.register(meta, handler as any);
export { meta };
