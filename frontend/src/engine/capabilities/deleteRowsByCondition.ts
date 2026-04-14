/**
 * deleteRowsByCondition – Delete rows where a column meets a condition.
 *
 * Iterates bottom-to-top to avoid index shifting issues when deleting
 * rows. Supports conditions: blank, notBlank, equals, notEquals,
 * contains, greaterThan, lessThan.
 */

import { CapabilityMeta, DeleteRowsByConditionParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "deleteRowsByCondition",
  description: "Delete rows where a column meets a specified condition",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: DeleteRowsByConditionParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { range: address, column, condition, value, hasHeaders = true } = params;

  if (options.dryRun) {
    const valStr = value !== undefined ? ` "${value}"` : "";
    return {
      stepId: "",
      status: "success",
      message: `Would delete rows in ${address} where column ${column} is ${condition}${valStr}`,
    };
  }

  options.onProgress?.("Reading range...");

  const rng = resolveRange(context, address);
  const used = rng.getUsedRange(false);
  used.load("values, rowCount, columnCount, address");
  await context.sync();

  const vals = (used.values ?? []) as (string | number | boolean | null)[][];
  if (!vals.length) {
    return { stepId: "", status: "success", message: "No data found." };
  }

  const colIdx = column - 1; // convert 1-based to 0-based
  const startDataRow = hasHeaders ? 1 : 0;

  // Find matching row indices (in original array)
  const matchingRows: number[] = [];
  for (let i = startDataRow; i < vals.length; i++) {
    const cellVal = vals[i][colIdx];
    if (meetsCondition(cellVal, condition, value)) {
      matchingRows.push(i);
    }
  }

  if (matchingRows.length === 0) {
    return { stepId: "", status: "success", message: `No rows matched condition "${condition}" in column ${column}.` };
  }

  options.onProgress?.(`Deleting ${matchingRows.length} rows...`);

  // Delete bottom-to-top to avoid index shifting
  for (let i = matchingRows.length - 1; i >= 0; i--) {
    const rowIdx = matchingRows[i];
    const rowRange = used.getRow(rowIdx);
    rowRange.delete("Up" as any);
  }

  await context.sync();

  const condDesc = value !== undefined ? `${condition} "${value}"` : condition;
  return {
    stepId: "",
    status: "success",
    message: `Deleted ${matchingRows.length} rows where column ${column} ${condDesc}`,
    outputs: { range: address, deletedCount: matchingRows.length },
  };
}

function meetsCondition(
  cellVal: string | number | boolean | null,
  condition: DeleteRowsByConditionParams["condition"],
  value?: string | number,
): boolean {
  const strVal = String(cellVal ?? "").trim();
  const isEmpty = cellVal === null || cellVal === "" || strVal === "";

  switch (condition) {
    case "blank":
      return isEmpty;
    case "notBlank":
      return !isEmpty;
    case "equals":
      if (value === undefined) return false;
      return strVal.toLowerCase() === String(value).toLowerCase();
    case "notEquals":
      if (value === undefined) return false;
      return strVal.toLowerCase() !== String(value).toLowerCase();
    case "contains":
      if (value === undefined) return false;
      return strVal.toLowerCase().includes(String(value).toLowerCase());
    case "greaterThan": {
      if (value === undefined) return false;
      const numCell = Number(cellVal);
      const numVal = Number(value);
      if (isNaN(numCell) || isNaN(numVal)) return strVal > String(value);
      return numCell > numVal;
    }
    case "lessThan": {
      if (value === undefined) return false;
      const numCell = Number(cellVal);
      const numVal = Number(value);
      if (isNaN(numCell) || isNaN(numVal)) return strVal < String(value);
      return numCell < numVal;
    }
    default:
      return false;
  }
}

registry.register(meta, handler as any);
export { meta };
