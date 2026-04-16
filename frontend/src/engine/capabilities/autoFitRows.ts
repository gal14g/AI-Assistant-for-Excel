/**
 * autoFitRows — row-axis pair of autoFitColumns.
 *
 * Office.js: Range.getEntireRow().format.autoFitRows() (ExcelApi 1.2+).
 */

import { CapabilityMeta, AutoFitRowsParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "autoFitRows",
  description: "Auto-fit row heights — row-axis pair of autoFitColumns",
  mutates: false,
  affectsFormatting: true,
  requiresApiSet: "ExcelApi 1.2",
};

async function handler(
  context: Excel.RequestContext,
  params: AutoFitRowsParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { range: address, sheetName } = params;

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would auto-fit rows in ${address ?? "active sheet used range"}` };
  }

  options.onProgress?.("Auto-fitting row heights...");

  let target: Excel.Range;
  if (address) {
    target = resolveRange(context, address);
  } else {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    target = sheet.getUsedRange(true);
  }
  target.getEntireRow().format.autofitRows();
  target.load("address");
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Auto-fit rows on ${target.address}.`,
    outputs: { range: target.address },
  };
}

registry.register(meta, handler as any);
export { meta };
