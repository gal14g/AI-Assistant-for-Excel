/**
 * autoFitColumns – Auto-fit column widths to their content.
 *
 * When a range is provided, only those columns are fitted.
 * Otherwise the entire used range on the sheet is fitted.
 *
 * Office.js notes:
 * - RangeFormat.autofitColumns() is available from ExcelApi 1.2.
 * - The call is synchronous within a batch — just queue and sync.
 */

import { CapabilityMeta, AutoFitColumnsParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "autoFitColumns",
  description: "Auto-fit column widths to their content",
  mutates: false,
  affectsFormatting: true,
};

async function handler(
  context: Excel.RequestContext,
  params: AutoFitColumnsParams,
  options: ExecutionOptions
): Promise<StepResult> {
  if (options.dryRun) {
    return { stepId: "", status: "success", message: "Would auto-fit column widths" };
  }

  options.onProgress?.("Auto-fitting column widths...");

  if (params.range) {
    resolveRange(context, params.range).format.autofitColumns();
  } else {
    const sheet = params.sheetName
      ? context.workbook.worksheets.getItem(params.sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    sheet.getUsedRange().format.autofitColumns();
  }

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: params.range
      ? `Auto-fitted columns in ${params.range}`
      : `Auto-fitted all used columns${params.sheetName ? ` on ${params.sheetName}` : ""}`,
  };
}

registry.register(meta, handler as any);
export { meta };
