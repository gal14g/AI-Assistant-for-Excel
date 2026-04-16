/**
 * sheetPosition — move a sheet to a specific position in the tab order.
 *
 * Office.js: worksheet.position (ExcelApi 1.7+).
 */

import { CapabilityMeta, SheetPositionParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { registerInverseOp } from "../snapshot";

const meta: CapabilityMeta = {
  action: "sheetPosition",
  description: "Move a sheet to a specific position (0-based) in the tab order",
  mutates: false,
  affectsFormatting: false,
  requiresApiSet: "ExcelApi 1.7",
};

async function handler(
  context: Excel.RequestContext,
  params: SheetPositionParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { position, sheetName } = params;

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would move sheet to position ${position}` };
  }

  const ws = sheetName
    ? context.workbook.worksheets.getItem(sheetName)
    : context.workbook.worksheets.getActiveWorksheet();
  ws.load(["name", "position"]);
  await context.sync();
  const previousPosition = ws.position;
  ws.position = position;
  await context.sync();
  // Undo = move back to the original tab position.
  registerInverseOp({ kind: "restoreSheetPosition", sheetName: ws.name, position: previousPosition });

  return {
    stepId: "",
    status: "success",
    message: `Moved sheet "${ws.name}" to position ${position}.`,
    outputs: { sheetName: ws.name },
  };
}

// ── Legacy fallback (< 1.7): can't move tabs programmatically ───────────────
async function fallback(
  _context: Excel.RequestContext,
  params: SheetPositionParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  if (options.dryRun) {
    return { stepId: "", status: "success", message: "Would skip sheet position (legacy fallback)." };
  }
  return {
    stepId: "",
    status: "success",
    message:
      `Sheet reposition skipped — Worksheet.position requires ExcelApi 1.7+. ` +
      `Drag the tab manually to position ${params.position + 1} (legacy-Excel fallback).`,
  };
}

registry.register(meta, handler as any, fallback as any);
export { meta };
