/**
 * writeFormula – Write a formula to a cell, optionally fill down.
 *
 * This is the preferred action when the planner can express the operation
 * as a native Excel formula. Native formulas are better because:
 * - They recalculate automatically when source data changes
 * - They are visible and auditable by the user
 * - They leverage Excel's optimized calculation engine
 *
 * Office.js notes:
 * - range.formulas accepts a 2D string array where each string starts with "=".
 * - fillDown() copies the formula and adjusts relative references automatically.
 * - For XLOOKUP, FILTER, and other dynamic array formulas, the formula
 *   "spills" into adjacent cells automatically (Excel 365+).
 */

import { CapabilityMeta, WriteFormulaParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "writeFormula",
  description: "Write a formula to a cell, optionally fill down",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: WriteFormulaParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const { cell, formula, fillDown } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would write formula ${formula} to ${cell}${fillDown ? ` and fill down ${fillDown} rows` : ""}`,
    };
  }

  options.onProgress?.(`Writing formula to ${cell}...`);

  const range = resolveRange(context, cell);

  // Write the formula to the single cell
  range.formulas = [[formula]];
  await context.sync();

  // Optionally fill down
  if (fillDown && fillDown > 1) {
    options.onProgress?.(`Filling formula down ${fillDown} rows...`);

    // Get the column letter and row from the cell reference
    const sheet = getSheet(context, cell);
    const cellRef = cell.includes("!") ? cell.split("!")[1] : cell;

    // Expand range to include fill-down rows
    const startCell = sheet.getRange(cellRef);
    const fillRange = startCell.getResizedRange(fillDown - 1, 0);
    fillRange.load("address");
    await context.sync();

    // Use fill down
    startCell.copyFrom(startCell, Excel.RangeCopyType.formulas);
    const fullRange = sheet.getRange(fillRange.address);
    fullRange.load("address");
    await context.sync();

    // Actually do the fill: set formula on first cell, then auto-fill
    const targetRange = sheet.getRange(fillRange.address);
    startCell.autoFill(targetRange, Excel.AutoFillType.fillDefault);
    await context.sync();

    return {
      stepId: "",
      status: "success",
      message: `Wrote formula to ${cell} and filled down ${fillDown} rows`,
    };
  }

  return {
    stepId: "",
    status: "success",
    message: `Wrote formula ${formula} to ${cell}`,
  };
}


function getSheet(context: Excel.RequestContext, address: string): Excel.Worksheet {
  if (address.includes("!")) {
    const idx = address.lastIndexOf("!");
    let sheet = address.substring(0, idx);
    if (sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1);
    return context.workbook.worksheets.getItem(sheet);
  }
  return context.workbook.worksheets.getActiveWorksheet();
}

registry.register(meta, handler as any);
export { meta };
