/**
 * setRowColSize – Set row height or column width manually.
 *
 * rowHeight is in points (default ~15).
 * columnWidth is in character widths (default ~8.43).
 */

import { CapabilityMeta, SetRowColSizeParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";

const meta: CapabilityMeta = {
  action: "setRowColSize",
  description: "Set row height or column width manually",
  mutates: false,
  affectsFormatting: true,
};

async function handler(
  context: Excel.RequestContext,
  params: SetRowColSizeParams,
  options: ExecutionOptions
): Promise<StepResult> {
  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would set ${params.dimension} to ${params.size} on ${params.range}` };
  }

  options.onProgress?.(`Setting ${params.dimension}...`);

  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange(params.range);

  if (params.dimension === "rowHeight") {
    range.format.rowHeight = params.size;
  } else {
    range.format.columnWidth = params.size;
  }

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Set ${params.dimension} to ${params.size} on ${params.range}`,
    outputs: { range: params.range },
  };
}

registry.register(meta, handler as any);
export { meta };
