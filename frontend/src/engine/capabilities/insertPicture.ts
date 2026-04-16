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

// ── Legacy-Excel fallback (ExcelApi < 1.9) ────────────────────────────────────
// worksheet.shapes.addImage requires 1.9 and Office.js on Excel 2016 RTM has
// no other programmatic path to embed an image (ActiveSheet.Pictures.Insert
// is VBA-only and unreachable from Office.js sandboxed add-ins). We emit a
// merged-cell placeholder labeled with the alt text so the user knows where
// the picture *would* go, and can insert it manually. Success (not error) —
// the plan shouldn't abort over a cosmetic asset that can be placed by hand.
async function fallback(
  context: Excel.RequestContext,
  params: InsertPictureParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { sheetName, left, top, width, height, altText } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would emit picture placeholder (legacy fallback; images not renderable on this Excel).`,
    };
  }

  options.onProgress?.("Legacy-Excel mode: images not renderable, writing placeholder cell...");

  const sheet = sheetName
    ? context.workbook.worksheets.getItem(sheetName)
    : context.workbook.worksheets.getActiveWorksheet();

  const COL_PT = 64;
  const ROW_PT = 15;
  const startCol = Math.max(0, Math.round((left ?? 0) / COL_PT));
  const startRow = Math.max(0, Math.round((top ?? 0) / ROW_PT));
  const colSpan = Math.max(2, Math.round((width ?? 128) / COL_PT));
  const rowSpan = Math.max(2, Math.round((height ?? 60) / ROW_PT));

  const block = sheet.getRangeByIndexes(startRow, startCol, rowSpan, colSpan);
  try { block.merge(true); } catch { /* may already be merged */ }

  const label = `[Image placeholder${altText ? ` — ${altText}` : ""}]`;
  // After merge we write only the top-left cell; grid shape for block.values:
  const grid: (string | null)[][] = [];
  for (let r = 0; r < rowSpan; r++) {
    const row: (string | null)[] = [];
    for (let c = 0; c < colSpan; c++) row.push(r === 0 && c === 0 ? label : null);
    grid.push(row);
  }
  try {
    block.values = grid as unknown as (string | number | boolean)[][];
  } catch {
    sheet.getRangeByIndexes(startRow, startCol, 1, 1).values = [[label]];
  }

  block.format.horizontalAlignment = Excel.HorizontalAlignment.center;
  block.format.verticalAlignment = Excel.VerticalAlignment.center;
  block.format.fill.color = "#F3F3F3";
  block.format.font.italic = true;
  block.format.font.color = "#606060";
  block.format.wrapText = true;

  const sides: Excel.BorderIndex[] = [
    Excel.BorderIndex.edgeTop,
    Excel.BorderIndex.edgeBottom,
    Excel.BorderIndex.edgeLeft,
    Excel.BorderIndex.edgeRight,
  ];
  for (const side of sides) {
    const b = block.format.borders.getItem(side);
    b.style = Excel.BorderLineStyle.dash;
    b.weight = Excel.BorderWeight.thin;
    b.color = "#A0A0A0";
  }
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message:
      `Emitted image placeholder at row ${startRow + 1}, col ${startCol + 1} — images require ` +
      `ExcelApi 1.9+ (worksheet.shapes.addImage). Office.js on Excel 2016 RTM has no supported ` +
      `path to embed pixels; insert the picture manually at this cell (legacy-Excel fallback).`,
  };
}

registry.register(meta, handler as any, fallback as any);
export { meta };
