/**
 * pivotCalculatedField – Add a calculated field to an existing PivotTable.
 *
 * Uses the PivotTable calculatedFields API to add a new field with a custom
 * formula. This requires ExcelApi 1.8+ and may not be available in all
 * Office.js environments.
 *
 * Note: The calculatedFields API has limited availability. If it is not
 * supported, a descriptive error message is returned.
 */

import { CapabilityMeta, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";

const meta: CapabilityMeta = {
  action: "pivotCalculatedField",
  description: "Add a calculated field to an existing PivotTable",
  mutates: true,
  affectsFormatting: false,
  requiresApiSet: "ExcelApi 1.8",
};

async function handler(
  context: Excel.RequestContext,
  params: any,
  options: ExecutionOptions
): Promise<StepResult> {
  const { pivotName, sheetName, fieldName, formula } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would add calculated field "${fieldName}" to pivot "${pivotName}"`,
    };
  }

  options.onProgress?.(`Adding calculated field "${fieldName}" to "${pivotName}"...`);

  try {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();

    const pivotTable = sheet.pivotTables.getItem(pivotName);

    // Attempt to add the calculated field
    // Note: calculatedFields may not be available in all API sets
    (pivotTable as any).calculatedFields.add(fieldName, formula);
    await context.sync();

    return {
      stepId: "",
      status: "success",
      message: `Added calculated field "${fieldName}" to pivot "${pivotName}"`,
    };
  } catch (err: any) {
    const message = err?.message ?? String(err);

    // Provide a helpful message if the API is not available
    if (
      message.includes("calculatedFields") ||
      message.includes("not a function") ||
      message.includes("undefined")
    ) {
      return {
        stepId: "",
        status: "error",
        message: `Calculated fields API is not available in this Office.js version. ` +
          `You may need to add the calculated field manually via the PivotTable Fields pane.`,
        error: message,
      };
    }

    return {
      stepId: "",
      status: "error",
      message: `Failed to add calculated field "${fieldName}": ${message}`,
      error: message,
    };
  }
}

registry.register(meta, handler as any);
export { meta };
