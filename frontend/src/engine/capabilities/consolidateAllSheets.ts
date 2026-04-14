/**
 * consolidateAllSheets – Merge data from ALL worksheets into one combined sheet.
 *
 * Iterates every worksheet in the workbook, reads used-range data from each,
 * and writes the combined result into a single output sheet.
 *
 * Options:
 *   outputSheetName  → name of the destination sheet (default "Combined")
 *   hasHeaders       → if true, only include headers from the first sheet
 *   excludeSheets    → sheet names to skip
 */

import { CapabilityMeta, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";

const meta: CapabilityMeta = {
  action: "consolidateAllSheets",
  description: "Merge data from all worksheets into one combined sheet",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: any,
  options: ExecutionOptions
): Promise<StepResult> {
  const {
    outputSheetName = "Combined",
    hasHeaders = true,
    excludeSheets = [],
  } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would consolidate all sheets into "${outputSheetName}"`,
    };
  }

  options.onProgress?.("Loading worksheets...");

  // Load all worksheet names
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();

  const excludeSet = new Set<string>([
    outputSheetName,
    ...(excludeSheets as string[]),
  ]);

  // Collect qualifying sheets and load their used ranges
  const qualifying: { name: string; range: Excel.Range }[] = [];
  for (const sheet of sheets.items) {
    if (excludeSet.has(sheet.name)) continue;
    const usedRange = sheet.getUsedRangeOrNullObject(true);
    usedRange.load(["values", "isNullObject"]);
    qualifying.push({ name: sheet.name, range: usedRange });
  }
  await context.sync();

  type Row = (string | number | boolean | null)[];
  const combined: Row[] = [];
  let sheetCount = 0;

  for (let i = 0; i < qualifying.length; i++) {
    const { name, range } = qualifying[i];
    if (range.isNullObject) continue;

    const values = (range.values ?? []) as Row[];
    if (values.length === 0) continue;

    options.onProgress?.(`Reading sheet "${name}" (${values.length} rows)...`);
    sheetCount++;

    if (combined.length === 0) {
      // First sheet: include everything (headers + data)
      combined.push(...values);
    } else {
      // Subsequent sheets: skip header row if hasHeaders
      const startRow = hasHeaders ? 1 : 0;
      for (let r = startRow; r < values.length; r++) {
        combined.push(values[r]);
      }
    }
  }

  if (combined.length === 0) {
    return {
      stepId: "",
      status: "success",
      message: "No data found in any qualifying sheets.",
    };
  }

  options.onProgress?.(`Writing ${combined.length} rows to "${outputSheetName}"...`);

  // Create or reuse the output sheet
  const existing = context.workbook.worksheets.getItemOrNullObject(outputSheetName);
  existing.load("isNullObject");
  await context.sync();

  const outSheet = existing.isNullObject
    ? context.workbook.worksheets.add(outputSheetName)
    : existing;

  // Write combined data
  const outRange = outSheet
    .getRange("A1")
    .getResizedRange(combined.length - 1, combined[0].length - 1);
  outRange.values = combined as any;
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Combined ${sheetCount} sheets (${combined.length} total rows) into "${outputSheetName}"`,
    outputs: { outputRange: `${outputSheetName}!A1` },
  };
}

registry.register(meta, handler as any);
export { meta };
