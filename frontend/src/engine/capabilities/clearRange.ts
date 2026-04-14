/**
 * clearRange – Clear contents, formatting, or both from a range.
 */

import { CapabilityMeta, ClearRangeParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "clearRange",
  description: "Clear a range's contents, formatting, or both",
  mutates: true,
  affectsFormatting: true,
};

async function handler(
  context: Excel.RequestContext,
  params: ClearRangeParams,
  options: ExecutionOptions
): Promise<StepResult> {
  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would clear ${params.clearType} from ${params.range}` };
  }

  options.onProgress?.(`Clearing ${params.clearType} from ${params.range}...`);

  const range = resolveRange(context, params.range);

  switch (params.clearType) {
    case "contents":
      range.clear("Contents" as Excel.ClearApplyTo);
      break;
    case "formats":
      range.clear("Formats" as Excel.ClearApplyTo);
      break;
    case "all":
    default:
      range.clear("All" as Excel.ClearApplyTo);
      break;
  }

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Cleared ${params.clearType} from ${params.range}`,
    outputs: { range: params.range },
  };
}

registry.register(meta, handler as any);
export { meta };
