/**
 * cloneSheetStructure – Copy a sheet's structure (headers, formatting, column
 * widths) but no data rows.
 *
 * Uses the native Worksheet.copy() method to duplicate the sheet, then clears
 * all content except the first (header) row.
 */

import { CapabilityMeta, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";

const meta: CapabilityMeta = {
  action: "cloneSheetStructure",
  description: "Copy a sheet's structure (headers + formatting) without data",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: any,
  options: ExecutionOptions
): Promise<StepResult> {
  const { sourceSheet, newSheetName } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would clone structure of "${sourceSheet}" → "${newSheetName}"`,
    };
  }

  options.onProgress?.(`Copying sheet "${sourceSheet}"...`);

  const source = context.workbook.worksheets.getItem(sourceSheet);
  const copy = source.copy("End");
  copy.load("name");
  await context.sync();

  // Rename the copy
  copy.name = newSheetName;
  await context.sync();

  // Load the used range so we can clear data but keep headers
  const usedRange = copy.getUsedRangeOrNullObject(true);
  usedRange.load(["isNullObject", "values", "rowCount", "columnCount"]);
  await context.sync();

  if (!usedRange.isNullObject && usedRange.rowCount > 1) {
    // Save the header row
    const headers = (usedRange.values as any[][])[0];

    // Clear all contents (formatting is preserved)
    usedRange.clear(Excel.ClearApplyTo.contents);
    await context.sync();

    // Write headers back to row 1
    const headerRange = copy
      .getRange("A1")
      .getResizedRange(0, headers.length - 1);
    headerRange.values = [headers];
    await context.sync();
  }

  return {
    stepId: "",
    status: "success",
    message: `Cloned structure of "${sourceSheet}" → "${newSheetName}" (headers + formatting, no data)`,
  };
}

registry.register(meta, handler as any);
export { meta };
