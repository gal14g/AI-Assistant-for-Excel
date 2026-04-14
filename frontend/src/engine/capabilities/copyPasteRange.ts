/**
 * copyPasteRange – Copy a range and paste to another location.
 *
 * Supports: all, values only, formats only, formulas only.
 * Uses Range.copyFrom() which handles cross-sheet copies.
 */

import { CapabilityMeta, CopyPasteRangeParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "copyPasteRange",
  description: "Copy a range and paste to another location",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: CopyPasteRangeParams,
  options: ExecutionOptions
): Promise<StepResult> {
  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would copy ${params.sourceRange} to ${params.destinationRange}` };
  }

  options.onProgress?.(`Copying ${params.sourceRange} → ${params.destinationRange}...`);

  const source = resolveRange(context, params.sourceRange);
  const dest = resolveRange(context, params.destinationRange);

  const pasteTypeMap: Record<string, Excel.RangeCopyType> = {
    all: "All" as Excel.RangeCopyType,
    values: "Values" as Excel.RangeCopyType,
    formats: "Formats" as Excel.RangeCopyType,
    formulas: "Formulas" as Excel.RangeCopyType,
  };

  const copyType = pasteTypeMap[params.pasteType ?? "all"] ?? ("All" as Excel.RangeCopyType);
  dest.copyFrom(source, copyType);
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Copied ${params.sourceRange} → ${params.destinationRange} (${params.pasteType ?? "all"})`,
    outputs: { outputRange: params.destinationRange },
  };
}

registry.register(meta, handler as any);
export { meta };
