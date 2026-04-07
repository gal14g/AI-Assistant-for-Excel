/**
 * addReportHeader – Insert a formatted report title row above data.
 *
 * Inserts a new row at the top (or above a specified range), writes the title
 * text, merges across the full width, and applies formatting (font size, fill
 * color, font color, bold, center alignment).
 */

import { CapabilityMeta, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange, resolveSheet } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "addReportHeader",
  description: "Insert a formatted report title row above data",
  mutates: true,
  affectsFormatting: true,
};

async function handler(
  context: Excel.RequestContext,
  params: any,
  options: ExecutionOptions
): Promise<StepResult> {
  const {
    title,
    sheetName,
    range,
    fontSize = 16,
    fillColor = "#4472C4",
    fontColor = "#FFFFFF",
    bold = true,
  } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would add report header "${title}"`,
    };
  }

  options.onProgress?.(`Adding report header "${title}"...`);

  // Determine the sheet
  const sheet = range
    ? resolveSheet(context, range)
    : sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();

  // Determine the column span from the used range or the provided range
  let columnCount: number;
  if (range) {
    const targetRange = resolveRange(context, range);
    targetRange.load("columnCount");
    await context.sync();
    columnCount = targetRange.columnCount;
  } else {
    const usedRange = sheet.getUsedRangeOrNullObject(true);
    usedRange.load(["isNullObject", "columnCount"]);
    await context.sync();
    columnCount = usedRange.isNullObject ? 5 : usedRange.columnCount;
  }

  // Insert a row at the top (shift existing data down)
  const insertRange = sheet.getRange("1:1");
  insertRange.insert(Excel.InsertShiftDirection.down);
  await context.sync();

  // Write the title to A1 and merge across the full width
  const headerRange = sheet
    .getRange("A1")
    .getResizedRange(0, columnCount - 1);
  headerRange.merge(false);

  // Set the title text
  const titleCell = sheet.getRange("A1");
  titleCell.values = [[title]];

  // Apply formatting
  const fmt = headerRange.format;
  fmt.font.size = fontSize;
  fmt.font.bold = bold;
  fmt.font.color = fontColor;
  fmt.fill.color = fillColor;
  fmt.horizontalAlignment = "Center" as Excel.HorizontalAlignment;
  fmt.verticalAlignment = "Center" as Excel.VerticalAlignment;
  fmt.rowHeight = fontSize * 2.5;

  await context.sync();

  const sheetLabel = sheetName ?? "active sheet";
  return {
    stepId: "",
    status: "success",
    message: `Added report header "${title}" to ${sheetLabel}`,
  };
}

registry.register(meta, handler as any);
export { meta };
