/**
 * groupRows – Group or ungroup rows/columns for outline collapsing.
 *
 * Office.js API: Range.group() / Range.ungroup() (ExcelApi 1.10+)
 */

import { CapabilityMeta, GroupRowsParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";

const meta: CapabilityMeta = {
  action: "groupRows",
  description: "Group or ungroup rows/columns for outline collapsing",
  mutates: false,
  affectsFormatting: true,
  requiresApiSet: "ExcelApi 1.10",
};

async function handler(
  context: Excel.RequestContext,
  params: GroupRowsParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const verb = params.operation === "group" ? "Group" : "Ungroup";

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would ${verb.toLowerCase()} ${params.range}` };
  }

  options.onProgress?.(`${verb}ing ${params.range}...`);

  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange(params.range);

  // Determine if this is a row or column range
  // "3:8" → rows, "B:E" → columns
  const ref = params.range.includes("!") ? params.range.split("!")[1] : params.range;
  const isRowRange = /^\d+:\d+$/.test(ref);
  const groupOption = isRowRange
    ? ("ByRows" as Excel.GroupOption)
    : ("ByColumns" as Excel.GroupOption);

  if (params.operation === "group") {
    range.group(groupOption);
  } else {
    range.ungroup(groupOption);
  }

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `${verb}ed ${params.range}`,
    outputs: { range: params.range },
  };
}

registry.register(meta, handler as any);
export { meta };
