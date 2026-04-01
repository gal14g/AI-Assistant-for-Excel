/**
 * insertTextBox – Insert a text box into a worksheet.
 *
 * Office.js notes:
 * - worksheet.shapes.addTextBox(text) creates a text box shape.
 * - Position and size via shape.left/top/width/height in pixels.
 * - Font properties via shape.textFrame.textRange.font.
 * - Fill via shape.fill.
 */

import { CapabilityMeta, InsertTextBoxParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";

const meta: CapabilityMeta = {
  action: "insertTextBox",
  description: "Insert a text box with custom text and formatting",
  mutates: true,
  affectsFormatting: false,
  requiresApiSet: "ExcelApi 1.9",
};

async function handler(
  context: Excel.RequestContext,
  params: InsertTextBoxParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const { sheetName, text, left, top, width, height, fontSize, fontFamily, fontColor, fillColor, horizontalAlignment } = params;

  if (options.dryRun) {
    const preview = text.length > 30 ? text.substring(0, 30) + "..." : text;
    return {
      stepId: "",
      status: "success",
      message: `Would insert text box "${preview}" at (${left}, ${top})`,
    };
  }

  options.onProgress?.("Inserting text box...");

  // Resolve target sheet
  let sheet: Excel.Worksheet;
  if (sheetName) {
    const ws = context.workbook.worksheets.getItemOrNullObject(sheetName);
    ws.load("isNullObject");
    await context.sync();
    if (ws.isNullObject) {
      return {
        stepId: "",
        status: "error",
        message: `Sheet "${sheetName}" not found. Please check the sheet name.`,
      };
    }
    sheet = ws;
  } else {
    sheet = context.workbook.worksheets.getActiveWorksheet();
  }

  const shape = sheet.shapes.addTextBox(text);

  shape.left = left;
  shape.top = top;
  shape.width = width;
  shape.height = height;

  // Font properties
  const font = shape.textFrame.textRange.font;
  if (fontSize !== undefined) font.size = fontSize;
  if (fontFamily !== undefined) font.name = fontFamily;
  if (fontColor !== undefined) font.color = fontColor;

  // Fill
  if (fillColor) {
    shape.fill.setSolidColor(fillColor);
  }

  // Horizontal alignment
  if (horizontalAlignment) {
    const alignMap: Record<string, Excel.ShapeTextHorizontalAlignment> = {
      left: Excel.ShapeTextHorizontalAlignment.left,
      center: Excel.ShapeTextHorizontalAlignment.center,
      right: Excel.ShapeTextHorizontalAlignment.right,
    };
    const align = alignMap[horizontalAlignment];
    if (align !== undefined) {
      shape.textFrame.horizontalAlignment = align;
    }
  }

  await context.sync();

  const preview = text.length > 30 ? text.substring(0, 30) + "..." : text;
  return {
    stepId: "",
    status: "success",
    message: `Inserted text box "${preview}" at (${left}, ${top})${sheetName ? ` on "${sheetName}"` : ""}`,
  };
}

registry.register(meta, handler as any);
export { meta };
