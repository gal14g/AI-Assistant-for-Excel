/**
 * removeDuplicates – Remove duplicate rows from a range.
 *
 * Office.js notes:
 * - Range.removeDuplicates() is available in ExcelApi 1.9+.
 * - columnIndexes specifies which columns to compare (0-based).
 * - If omitted, all columns are compared.
 * - This modifies the range in-place; rows are deleted.
 */

import { CapabilityMeta, RemoveDuplicatesParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "removeDuplicates",
  description: "Remove duplicate rows from a range",
  mutates: true,
  affectsFormatting: false,
  requiresApiSet: "ExcelApi 1.9",
};

async function handler(
  context: Excel.RequestContext,
  params: RemoveDuplicatesParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const { range: address, columnIndexes } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would remove duplicates from ${address}`,
    };
  }

  options.onProgress?.("Removing duplicates...");

  const range = resolveRange(context, address);
  const result = range.removeDuplicates(
    columnIndexes ?? [],
    true // includesHeader – default to true
  );
  result.load(["removed", "uniqueRemaining"]);
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Removed ${result.removed} duplicate rows; ${result.uniqueRemaining} unique rows remain`,
  };
}


registry.register(meta, handler as any);
export { meta };
