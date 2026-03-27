/**
 * insertDeleteRows – Insert or delete rows/columns at a range.
 *
 * The range determines which rows or columns are affected (how many and where).
 * The shiftDirection controls whether to insert or delete, and the axis:
 *
 *   "down"  → insert blank rows above the range (existing rows shift down)
 *   "up"    → delete the range's rows (remaining rows shift up)
 *   "right" → insert blank columns to the left of the range (existing shift right)
 *   "left"  → delete the range's columns (remaining columns shift left)
 *
 * Examples:
 *   Insert 2 rows above row 5:  range="Sheet1!A5:A6", shiftDirection="down"
 *   Delete rows 3–5:            range="Sheet1!A3:A5", shiftDirection="up"
 *   Insert 1 column before B:   range="Sheet1!B1:B1", shiftDirection="right"
 *
 * Office.js notes:
 * - range.insert() / range.delete() are available from ExcelApi 1.1.
 * - These operations shift all data in the affected direction, preserving
 *   relative references in formulas where possible.
 */

import { CapabilityMeta, InsertDeleteRowsParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "insertDeleteRows",
  description: "Insert or delete rows/columns at a specified range",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: InsertDeleteRowsParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const { range: address, shiftDirection } = params;
  const isInsert = shiftDirection === "down" || shiftDirection === "right";

  if (options.dryRun) {
    const op = isInsert ? "insert" : "delete";
    const axis = shiftDirection === "down" || shiftDirection === "up" ? "rows" : "columns";
    return { stepId: "", status: "success", message: `Would ${op} ${axis} at ${address}` };
  }

  const range = resolveRange(context, address);

  if (isInsert) {
    const direction =
      shiftDirection === "down"
        ? Excel.InsertShiftDirection.down
        : Excel.InsertShiftDirection.right;
    range.insert(direction);
  } else {
    const direction =
      shiftDirection === "up"
        ? Excel.DeleteShiftDirection.up
        : Excel.DeleteShiftDirection.left;
    range.delete(direction);
  }

  await context.sync();

  const opLabel = isInsert ? "Inserted" : "Deleted";
  const axisLabel = shiftDirection === "down" || shiftDirection === "up" ? "rows" : "columns";
  return {
    stepId: "",
    status: "success",
    message: `${opLabel} ${axisLabel} at ${address}`,
  };
}

registry.register(meta, handler as any);
export { meta };
