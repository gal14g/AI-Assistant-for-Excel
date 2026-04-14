/**
 * addSlicer – Add a slicer to filter a PivotTable or Table.
 *
 * Office.js notes:
 * - workbook.slicers.add(source, sourceField, destinationSheet?) creates a slicer.
 * - The source can be a PivotTable name or Table name.
 * - Position and size are set via slicer.left/top/width/height.
 */

import { CapabilityMeta, AddSlicerParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";

const meta: CapabilityMeta = {
  action: "addSlicer",
  description: "Add a slicer to filter a PivotTable or Table",
  mutates: true,
  affectsFormatting: false,
  requiresApiSet: "ExcelApi 1.10",
};

async function handler(
  context: Excel.RequestContext,
  params: AddSlicerParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const { sheetName, sourceName, sourceField, left, top, width, height } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would add slicer for field "${sourceField}" on "${sourceName}"`,
    };
  }

  options.onProgress?.(`Adding slicer for "${sourceField}"...`);

  // Resolve destination sheet if provided
  let destinationSheet: Excel.Worksheet | undefined;
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
    destinationSheet = ws;
  }

  const slicer = destinationSheet
    ? context.workbook.slicers.add(sourceName, sourceField, destinationSheet)
    : context.workbook.slicers.add(sourceName, sourceField);

  if (left !== undefined) slicer.left = left;
  if (top !== undefined) slicer.top = top;
  if (width !== undefined) slicer.width = width;
  if (height !== undefined) slicer.height = height;

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Added slicer for "${sourceField}" on "${sourceName}"${sheetName ? ` on sheet "${sheetName}"` : ""}`,
    outputs: { slicerName: sourceField },
  };
}

registry.register(meta, handler as any);
export { meta };
