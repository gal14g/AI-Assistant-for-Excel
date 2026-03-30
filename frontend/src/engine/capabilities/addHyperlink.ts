/**
 * addHyperlink – Insert a hyperlink in a cell.
 */

import { CapabilityMeta, AddHyperlinkParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "addHyperlink",
  description: "Insert a hyperlink in a cell",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: AddHyperlinkParams,
  options: ExecutionOptions
): Promise<StepResult> {
  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would add hyperlink to ${params.cell}` };
  }

  options.onProgress?.("Inserting hyperlink...");

  const range = resolveRange(context, params.cell);
  range.hyperlink = {
    address: params.url,
    textToDisplay: params.displayText ?? params.url,
  };
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Added hyperlink to ${params.cell}: ${params.url}`,
  };
}

registry.register(meta, handler as any);
export { meta };
