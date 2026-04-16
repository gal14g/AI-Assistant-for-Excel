/**
 * insertShape – Insert a geometric shape into a worksheet.
 *
 * Office.js notes:
 * - worksheet.shapes.addGeometricShape(type) creates a shape.
 * - Position and size via shape.left/top/width/height in pixels.
 * - Fill and line properties are set via shape.fill and shape.lineFormat.
 * - Text is added via shape.textFrame.textRange.
 */

import { CapabilityMeta, InsertShapeParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";

const meta: CapabilityMeta = {
  action: "insertShape",
  description: "Insert a geometric shape (rectangle, oval, arrow, star, etc.)",
  mutates: true,
  affectsFormatting: false,
  requiresApiSet: "ExcelApi 1.9",
};

async function handler(
  context: Excel.RequestContext,
  params: InsertShapeParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const { sheetName, shapeType, left, top, width, height, fillColor, lineColor, lineWeight, textContent } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would insert ${shapeType} shape at (${left}, ${top})`,
    };
  }

  options.onProgress?.(`Inserting ${shapeType} shape...`);

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

  // Map string shape type to Excel.GeometricShapeType
  const shapeTypeMap: Record<string, Excel.GeometricShapeType> = {
    rectangle: Excel.GeometricShapeType.rectangle,
    oval: Excel.GeometricShapeType.ellipse,
    diamond: Excel.GeometricShapeType.diamond,
    rightTriangle: Excel.GeometricShapeType.rightTriangle,
    rightArrow: Excel.GeometricShapeType.rightArrow,
    leftArrow: Excel.GeometricShapeType.leftArrow,
    upArrow: Excel.GeometricShapeType.upArrow,
    downArrow: Excel.GeometricShapeType.downArrow,
    star5: Excel.GeometricShapeType.star5,
    heart: Excel.GeometricShapeType.heart,
  };

  const excelShapeType = shapeTypeMap[shapeType];
  if (excelShapeType === undefined) {
    return {
      stepId: "",
      status: "error",
      message: `Unknown shape type "${shapeType}". Supported: ${Object.keys(shapeTypeMap).join(", ")}`,
    };
  }

  const shape = sheet.shapes.addGeometricShape(excelShapeType);

  shape.left = left;
  shape.top = top;
  shape.width = width;
  shape.height = height;

  if (fillColor) {
    shape.fill.setSolidColor(fillColor);
  }
  if (lineColor) {
    shape.lineFormat.color = lineColor;
  }
  if (lineWeight !== undefined) {
    shape.lineFormat.weight = lineWeight;
  }
  if (textContent) {
    shape.textFrame.textRange.text = textContent;
  }

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Inserted ${shapeType} shape at (${left}, ${top})${sheetName ? ` on "${sheetName}"` : ""}`,
  };
}

// ── Legacy-Excel fallback (ExcelApi < 1.9) ────────────────────────────────────
// The Shapes API doesn't exist on Excel 2016. We simulate rectangles (and
// rectangle-shaped substitutes for diamond/oval/triangle) as merged-cell
// blocks with thick borders and an optional fill color + caption. Arrows and
// lines cannot be drawn convincingly on a cell grid — for those we fall back
// to a single cell with an arrow glyph (▶ ◀ ▲ ▼) as display text so the
// *intent* survives even if the visual doesn't.
//
// Position mapping: the source `left/top/width/height` are in points (Office.js
// convention). We divide by approximate defaults (column ≈ 64pt, row ≈ 15pt)
// to pick a landing cell block. This is lossy — the fallback message records
// the approximation so the user can nudge it.
async function fallback(
  context: Excel.RequestContext,
  params: InsertShapeParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const {
    sheetName, shapeType, left, top, width, height,
    fillColor, lineColor, lineWeight, textContent,
  } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would emit cell-grid approximation of ${shapeType} shape (legacy fallback).`,
    };
  }

  options.onProgress?.(`Legacy-Excel mode: faking ${shapeType} via merged cells + borders...`);

  // Resolve target sheet.
  const sheet = sheetName
    ? context.workbook.worksheets.getItem(sheetName)
    : context.workbook.worksheets.getActiveWorksheet();

  // Approximate column/row size (Excel defaults — columns ≈ 64pt wide, rows ≈ 15pt tall).
  const COL_PT = 64;
  const ROW_PT = 15;
  const startCol = Math.max(0, Math.round(left / COL_PT));
  const startRow = Math.max(0, Math.round(top / ROW_PT));
  const colSpan = Math.max(1, Math.round(width / COL_PT));
  const rowSpan = Math.max(1, Math.round(height / ROW_PT));

  // Arrows/lines: cell-grid can't draw diagonals. Place a glyph in a single
  // cell and skip the merged-block treatment.
  const arrowGlyphs: Record<string, string> = {
    rightArrow: "▶",
    leftArrow:  "◀",
    upArrow:    "▲",
    downArrow:  "▼",
  };
  if (arrowGlyphs[shapeType]) {
    const cell = sheet.getRangeByIndexes(startRow, startCol, 1, 1);
    cell.values = [[textContent ?? arrowGlyphs[shapeType]]];
    cell.format.font.bold = true;
    cell.format.font.size = Math.max(10, Math.round(height / 3));
    if (fillColor) cell.format.fill.color = fillColor;
    cell.format.horizontalAlignment = Excel.HorizontalAlignment.center;
    cell.format.verticalAlignment = Excel.VerticalAlignment.center;
    await context.sync();
    return {
      stepId: "",
      status: "success",
      message:
        `Wrote ${shapeType} glyph "${arrowGlyphs[shapeType]}" at approximately ` +
        `row ${startRow + 1}, col ${startCol + 1} — shapes require ExcelApi 1.9+, ` +
        `arrows cannot be drawn on the cell grid (legacy-Excel fallback).`,
    };
  }

  // Rectangle / oval / diamond / triangle / star / heart: merged-cell block
  // with borders. Oval/diamond etc. are approximated as rectangles — callers
  // get a warning in the audit line.
  const block = sheet.getRangeByIndexes(startRow, startCol, rowSpan, colSpan);
  if (rowSpan > 1 || colSpan > 1) {
    try { block.merge(true); } catch { /* range may already be merged */ }
  }
  if (fillColor) block.format.fill.color = fillColor;

  // Thick border to hint at shape outline.
  const borderColor = lineColor ?? "#000000";
  const borderWeight = (lineWeight !== undefined && lineWeight >= 2)
    ? Excel.BorderWeight.thick
    : Excel.BorderWeight.medium;
  const sides: Excel.BorderIndex[] = [
    Excel.BorderIndex.edgeTop,
    Excel.BorderIndex.edgeBottom,
    Excel.BorderIndex.edgeLeft,
    Excel.BorderIndex.edgeRight,
  ];
  for (const side of sides) {
    const b = block.format.borders.getItem(side);
    b.style = Excel.BorderLineStyle.continuous;
    b.color = borderColor;
    b.weight = borderWeight;
  }

  // Caption text (optional)
  if (textContent) {
    block.values = [[textContent, ...Array(colSpan - 1).fill("")], ...Array(Math.max(0, rowSpan - 1)).fill(Array(colSpan).fill(""))].slice(0, rowSpan);
    block.format.horizontalAlignment = Excel.HorizontalAlignment.center;
    block.format.verticalAlignment = Excel.VerticalAlignment.center;
  }
  await context.sync();

  const approxNote = (shapeType === "rectangle")
    ? ""
    : ` — ${shapeType} approximated as a rectangle on the cell grid (no curved/diagonal primitives on ExcelApi 1.3)`;

  return {
    stepId: "",
    status: "success",
    message:
      `Inserted ${shapeType} approximation at row ${startRow + 1}, col ${startCol + 1} ` +
      `(${rowSpan}×${colSpan} merged block)${approxNote} — shapes require ExcelApi 1.9+ ` +
      `(legacy-Excel fallback).`,
  };
}

registry.register(meta, handler as any, fallback as any);
export { meta };
