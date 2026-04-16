/**
 * freezePanes – Freeze rows and/or columns.
 *
 * Office.js notes:
 * - WorksheetFreezePanes API is in ExcelApi 1.7+.
 * - freezeAt(range) freezes rows above and columns left of the given range.
 * - e.g. freezeAt("B2") freezes row 1 and column A.
 */

import { CapabilityMeta, FreezePanesParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";

const meta: CapabilityMeta = {
  action: "freezePanes",
  description: "Freeze rows and/or columns at a specified cell",
  mutates: false,
  affectsFormatting: false,
  requiresApiSet: "ExcelApi 1.7",
};

async function handler(
  context: Excel.RequestContext,
  params: FreezePanesParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const { cell, sheetName } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would freeze panes at ${cell}`,
    };
  }

  options.onProgress?.(`Freezing panes at ${cell}...`);

  // Validate and resolve target sheet
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

  const range = sheet.getRange(cell);
  sheet.freezePanes.freezeAt(range);
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Froze panes at ${cell}${sheetName ? ` on "${sheetName}"` : ""}`,
  };
}

// ── Legacy-Excel fallback (ExcelApi < 1.7) ────────────────────────────────────
// WorksheetFreezePanes requires 1.7. Before that, Office.js offers no
// programmatic freeze. We can't simulate the split-scrolling behavior
// through any other primitive, so we gracefully skip with a warning — the
// user can freeze manually via View › Freeze Panes. Status=success because
// this is cosmetic and shouldn't abort a data plan.
async function fallback(
  _context: Excel.RequestContext,
  params: FreezePanesParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would skip freeze at ${params.cell} (legacy fallback; requires ExcelApi 1.7+).`,
    };
  }

  options.onProgress?.("Legacy-Excel mode: freeze panes unavailable, skipping...");

  return {
    stepId: "",
    status: "success",
    message:
      `Freeze panes at ${params.cell} skipped — WorksheetFreezePanes requires ExcelApi 1.7+ ` +
      `(Excel 2019 / 2021 / Microsoft 365). Use View › Freeze Panes manually ` +
      `(legacy-Excel fallback).`,
  };
}

registry.register(meta, handler as any, fallback as any);
export { meta };
