/**
 * hideShow – Hide or unhide rows, columns, or entire sheets.
 */

import { CapabilityMeta, HideShowParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";

const meta: CapabilityMeta = {
  action: "hideShow",
  description: "Hide or unhide rows, columns, or sheets",
  mutates: false,
  affectsFormatting: true,
};

async function handler(
  context: Excel.RequestContext,
  params: HideShowParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const verb = params.hide ? "Hide" : "Unhide";

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would ${verb.toLowerCase()} ${params.target}: ${params.rangeOrName}` };
  }

  options.onProgress?.(`${verb}ing ${params.target}...`);

  if (params.target === "sheet") {
    const sheet = context.workbook.worksheets.getItem(params.rangeOrName);
    sheet.visibility = params.hide
      ? ("Hidden" as Excel.SheetVisibility)
      : ("Visible" as Excel.SheetVisibility);
  } else if (params.target === "rows") {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(params.rangeOrName);
    range.rowHidden = params.hide;
  } else if (params.target === "columns") {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(params.rangeOrName);
    range.columnHidden = params.hide;
  }

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `${verb} ${params.target}: ${params.rangeOrName}`,
  };
}

registry.register(meta, handler as any);
export { meta };
