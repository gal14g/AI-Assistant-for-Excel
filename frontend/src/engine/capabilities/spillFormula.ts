/**
 * spillFormula – Write a dynamic array formula that spills automatically.
 *
 * Dynamic array formulas (FILTER, SORT, UNIQUE, SEQUENCE, etc.) write a
 * single formula to one cell and Excel spills the results into adjacent
 * cells. This capability writes the formula and reports the spill size.
 *
 * Office.js notes:
 * - Dynamic arrays are supported in Excel 365+.
 * - getSpillingToRangeOrNullObject() returns the spill range after sync.
 * - The formula must start with "=".
 */

import { CapabilityMeta, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "spillFormula",
  description: "Write a dynamic array formula (FILTER, SORT, UNIQUE, etc.) that spills automatically",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: any,
  options: ExecutionOptions
): Promise<StepResult> {
  const { cell, formula, sheetName } = params;

  if (!cell || !formula) {
    return {
      stepId: "",
      status: "error",
      message: "cell and formula are required",
    };
  }

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would write dynamic array formula to ${cell}: ${formula}`,
    };
  }

  options.onProgress?.(`Writing dynamic array formula to ${cell}...`);

  // Resolve target cell — prepend sheetName if provided
  const cellAddress = sheetName && !cell.includes("!") ? `${sheetName}!${cell}` : cell;
  const range = resolveRange(context, cellAddress);

  // Write the formula to the single cell
  range.formulas = [[formula]];
  await context.sync();

  // Check for formula errors
  range.load("values");
  await context.sync();

  const firstVal = String(range.values?.[0]?.[0] ?? "");
  const errorTypes = ["#SPILL!", "#REF!", "#VALUE!", "#NAME?", "#NULL!", "#N/A", "#DIV/0!", "#CALC!"];
  const hasError = errorTypes.some((e) => firstVal.includes(e));

  if (hasError) {
    return {
      stepId: "",
      status: "error",
      message: `Formula wrote to ${cell} but produced ${firstVal}. The formula may need to be corrected.`,
      error: `Formula error: ${firstVal}`,
    };
  }

  // Try to determine the spill range size
  let spillInfo = "spill size unknown";
  try {
    const spillRange = range.getSpillingToRangeOrNullObject();
    spillRange.load(["isNullObject", "rowCount", "columnCount", "address"]);
    await context.sync();

    if (!spillRange.isNullObject) {
      const totalCells = spillRange.rowCount * spillRange.columnCount;
      spillInfo = `spilled to ${totalCells} cells (${spillRange.address})`;
    }
  } catch {
    // getSpillingToRangeOrNullObject may not be available in all API sets
    spillInfo = "spill size unknown";
  }

  return {
    stepId: "",
    status: "success",
    message: `Wrote dynamic array formula to ${cell}, ${spillInfo}`,
    outputs: { cell: cellAddress },
  };
}


registry.register(meta, handler as any);
export { meta };
