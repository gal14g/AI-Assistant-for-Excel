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

registry.register(meta, handler as any);
export { meta };
