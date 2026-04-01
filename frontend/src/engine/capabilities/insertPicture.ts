/**
 * insertPicture – Insert an image into a worksheet.
 *
 * Office.js notes:
 * - worksheet.shapes.addImage(base64) adds an image from a base64-encoded string.
 * - Position and size are set via shape.left/top/width/height in pixels.
 * - The base64 string should not include the data URI prefix.
 */

import { CapabilityMeta, InsertPictureParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";

const meta: CapabilityMeta = {
  action: "insertPicture",
  description: "Insert an image into a worksheet from base64 data",
  mutates: true,
  affectsFormatting: false,
  requiresApiSet: "ExcelApi 1.9",
};

async function handler(
  context: Excel.RequestContext,
  params: InsertPictureParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const { sheetName, imageBase64, left, top, width, height, altText } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would insert picture${sheetName ? ` on "${sheetName}"` : ""}`,
    };
  }

  options.onProgress?.("Inserting picture...");

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

  const shape = sheet.shapes.addImage(imageBase64);

  if (left !== undefined) shape.left = left;
  if (top !== undefined) shape.top = top;
  if (width !== undefined) shape.width = width;
  if (height !== undefined) shape.height = height;
  if (altText !== undefined) shape.altTextDescription = altText;

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Inserted picture${sheetName ? ` on "${sheetName}"` : ""}`,
  };
}

registry.register(meta, handler as any);
export { meta };
