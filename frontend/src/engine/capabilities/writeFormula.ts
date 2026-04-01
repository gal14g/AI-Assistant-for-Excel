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

  // Optionally fill down — autoFill adjusts relative references automatically
  if (fillDown && fillDown > 1) {
    options.onProgress?.(`Filling formula down ${fillDown} rows...`);
    const fillRange = range.getResizedRange(fillDown - 1, 0);
    range.autoFill(fillRange, Excel.AutoFillType.fillDefault);
    await context.sync();
  }

  // Check for formula errors after writing (#SPILL!, #REF!, #VALUE!, #NAME?, etc.)
  const checkRange = fillDown && fillDown > 1
    ? range.getResizedRange(Math.min(fillDown - 1, 5), 0) // check first few rows
    : range;
  checkRange.load("values");
  await context.sync();

  const errorTypes = ["#SPILL!", "#REF!", "#VALUE!", "#NAME?", "#NULL!", "#N/A", "#DIV/0!", "#CALC!"];
  const firstVal = String(checkRange.values?.[0]?.[0] ?? "");
  const hasError = errorTypes.some((e) => firstVal.includes(e));

  const rowInfo = fillDown && fillDown > 1 ? ` and filled down ${fillDown} rows` : "";
  if (hasError) {
    return {
      stepId: "",
      status: "error",
      message: `Formula wrote to ${cell}${rowInfo} but produced ${firstVal}. The formula may need to be corrected.`,
      error: `Formula error: ${firstVal}`,
    };
  }

  return {
    stepId: "",
    status: "success",
    message: `Wrote formula to ${cell}${rowInfo}`,
  };
}



registry.register(meta, handler as any);
export { meta };
