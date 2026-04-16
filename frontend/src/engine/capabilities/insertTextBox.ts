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

// ── Legacy-Excel fallback (ExcelApi < 1.9) ────────────────────────────────────
// TextBox shapes require 1.9. We substitute a merged cell block with centered
// wrapped text, optional fill color, and a light border. Fidelity cost: no
// free positioning (snaps to the grid), no auto-size-to-content, no overlay
// over other cells. Most annotation/label use cases survive this intact.
async function fallback(
  context: Excel.RequestContext,
  params: InsertTextBoxParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const {
    sheetName, text, left, top, width, height,
    fontSize, fontFamily, fontColor, fillColor, horizontalAlignment,
  } = params;

  if (options.dryRun) {
    const preview = text.length > 30 ? text.substring(0, 30) + "..." : text;
    return {
      stepId: "",
      status: "success",
      message: `Would emit merged-cell textbox "${preview}" (legacy fallback).`,
    };
  }

  options.onProgress?.("Legacy-Excel mode: faking textbox via merged cells...");

  const sheet = sheetName
    ? context.workbook.worksheets.getItem(sheetName)
    : context.workbook.worksheets.getActiveWorksheet();

  // Approximate column/row size (Excel defaults).
  const COL_PT = 64;
  const ROW_PT = 15;
  const startCol = Math.max(0, Math.round(left / COL_PT));
  const startRow = Math.max(0, Math.round(top / ROW_PT));
  const colSpan = Math.max(1, Math.round(width / COL_PT));
  const rowSpan = Math.max(1, Math.round(height / ROW_PT));

  const block = sheet.getRangeByIndexes(startRow, startCol, rowSpan, colSpan);
  if (rowSpan > 1 || colSpan > 1) {
    try { block.merge(true); } catch { /* may already be merged */ }
  }

  // Write the text into the (now merged) top-left cell. values on a merged
  // range is a single-cell grid, but Office.js expects a rowSpan×colSpan 2D
  // for range.values. After merge the range reports as the merged block; feed
  // a 2D array of matching shape where only [0][0] has the text.
  const grid: (string | null)[][] = [];
  for (let r = 0; r < rowSpan; r++) {
    const row: (string | null)[] = [];
    for (let c = 0; c < colSpan; c++) row.push(r === 0 && c === 0 ? text : null);
    grid.push(row);
  }
  try {
    block.values = grid as unknown as (string | number | boolean)[][];
  } catch {
    // After merge Office.js may prefer a 1×1 write to the merged cell.
    sheet.getRangeByIndexes(startRow, startCol, 1, 1).values = [[text]];
  }

  // Font
  if (fontSize !== undefined) block.format.font.size = fontSize;
  if (fontFamily !== undefined) block.format.font.name = fontFamily;
  if (fontColor !== undefined) block.format.font.color = fontColor;

  // Fill
  if (fillColor) block.format.fill.color = fillColor;

  // Alignment — default to center so the merged block looks like a textbox.
  const alignMap: Record<string, Excel.HorizontalAlignment> = {
    left:   Excel.HorizontalAlignment.left,
    center: Excel.HorizontalAlignment.center,
    right:  Excel.HorizontalAlignment.right,
  };
  block.format.horizontalAlignment = alignMap[horizontalAlignment ?? "center"] ?? Excel.HorizontalAlignment.center;
  block.format.verticalAlignment = Excel.VerticalAlignment.center;
  block.format.wrapText = true;

  // Border to visually separate the "textbox" from surrounding cells.
  const sides: Excel.BorderIndex[] = [
    Excel.BorderIndex.edgeTop,
    Excel.BorderIndex.edgeBottom,
    Excel.BorderIndex.edgeLeft,
    Excel.BorderIndex.edgeRight,
  ];
  for (const side of sides) {
    const b = block.format.borders.getItem(side);
    b.style = Excel.BorderLineStyle.continuous;
    b.weight = Excel.BorderWeight.thin;
    b.color = "#808080";
  }

  await context.sync();

  const preview = text.length > 30 ? text.substring(0, 30) + "..." : text;
  return {
    stepId: "",
    status: "success",
    message:
      `Inserted textbox approximation "${preview}" at row ${startRow + 1}, col ${startCol + 1} ` +
      `(${rowSpan}×${colSpan} merged block) — textbox shapes require ExcelApi 1.9+ ` +
      `(legacy-Excel fallback; snaps to cell grid, no free positioning).`,
  };
}

registry.register(meta, handler as any, fallback as any);
export { meta };
