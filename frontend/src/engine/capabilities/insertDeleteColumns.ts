/**
 * insertDeleteColumns — parity with insertDeleteRows.
 *
 * Accepts a column-letter range like "C:E" (or a cell-address range whose
 * column span we infer) and either inserts blank columns (shifting existing
 * ones) or deletes columns (shifting neighbors).
 */

import { CapabilityMeta, InsertDeleteColumnsParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";
import { registerInverseOp } from "../snapshot";

const meta: CapabilityMeta = {
  action: "insertDeleteColumns",
  description: "Insert or delete worksheet columns — column-axis pair of insertDeleteRows",
  mutates: true,
  affectsFormatting: false,
  requiresApiSet: "ExcelApi 1.1",
};

async function handler(
  context: Excel.RequestContext,
  params: InsertDeleteColumnsParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { range: address, action, shiftDirection = "right" } = params;

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would ${action} columns ${address}` };
  }

  options.onProgress?.(`${action === "insert" ? "Inserting" : "Deleting"} columns ${address}...`);

  // If the user gave a cell-address range like "C1:E10", getEntireColumn() on
  // that range resolves to the columns C:E. For a plain "C:E" the range is
  // already full-column.
  const range = resolveRange(context, address);
  const colRange = range.getEntireColumn();
  colRange.load(["columnCount", "address"]);
  await context.sync();

  try {
    if (action === "insert") {
      const dir = shiftDirection === "left"
        ? Excel.InsertShiftDirection.right // "shift existing cells right" = inserting to the LEFT of them
        : Excel.InsertShiftDirection.right;
      // For whole-column inserts the direction is moot (you're inserting full columns);
      // we keep the param for API symmetry with insertDeleteRows.
      colRange.insert(dir);
      await context.sync();
      // Undo an insert = delete the columns we just added.
      colRange.load(["address", "worksheet/name"]);
      await context.sync();
      const sheetName = colRange.worksheet.name;
      const rangeAddr = colRange.address.includes("!") ? colRange.address.split("!").pop()! : colRange.address;
      registerInverseOp({ kind: "deleteColumns", sheetName, rangeAddress: rangeAddr });
    } else {
      colRange.delete(Excel.DeleteShiftDirection.left);
      await context.sync();
    }
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return {
      stepId: "",
      status: "error",
      message: `Failed to ${action} columns ${address}: ${msg}`,
      error: msg,
    };
  }

  return {
    stepId: "",
    status: "success",
    message: `${action === "insert" ? "Inserted" : "Deleted"} columns ${address}.`,
    outputs: { range: address, columnCount: colRange.columnCount },
  };
}

registry.register(meta, handler as any);
export { meta };
