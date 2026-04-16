/**
 * tabColor — set the color of a worksheet's tab.
 *
 * Office.js: worksheet.tabColor (ExcelApi 1.7+). Accepts a hex string or
 * "none" to clear.
 */

import { CapabilityMeta, TabColorParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { registerInverseOp } from "../snapshot";

const meta: CapabilityMeta = {
  action: "tabColor",
  description: "Set or clear the color of a worksheet's tab",
  mutates: false,
  affectsFormatting: true,
  requiresApiSet: "ExcelApi 1.7",
};

async function handler(
  context: Excel.RequestContext,
  params: TabColorParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { color, sheetName } = params;

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would set tab color to ${color}` };
  }

  const ws = sheetName
    ? context.workbook.worksheets.getItem(sheetName)
    : context.workbook.worksheets.getActiveWorksheet();

  // Capture the previous tab color BEFORE we overwrite it, so undo can
  // restore the exact prior state (including "cleared/no color").
  ws.load(["name", "tabColor"]);
  await context.sync();
  const previousColor = (ws.tabColor ?? "") as string;

  // Office.js accepts hex strings; "" clears the color. Map "none" → "".
  const normalized = color.trim().toLowerCase() === "none" ? "" : color;
  ws.tabColor = normalized;
  await context.sync();

  // Register undo: set the tab color back to whatever it was.
  registerInverseOp({ kind: "restoreTabColor", sheetName: ws.name, color: previousColor });

  return {
    stepId: "",
    status: "success",
    message:
      normalized === ""
        ? `Cleared tab color on "${ws.name}".`
        : `Set tab color on "${ws.name}" to ${normalized}.`,
    outputs: { sheetName: ws.name },
  };
}

// ── Legacy-Excel fallback (ExcelApi < 1.7) ──────────────────────────────────
async function fallback(
  _context: Excel.RequestContext,
  params: TabColorParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  if (options.dryRun) {
    return { stepId: "", status: "success", message: "Would skip tab color (legacy fallback)." };
  }
  return {
    stepId: "",
    status: "success",
    message:
      `Tab color skipped — Worksheet.tabColor requires ExcelApi 1.7+ (Excel 2019+). ` +
      `Right-click the sheet tab → Tab Color to set manually (legacy-Excel fallback).`,
  };
}

registry.register(meta, handler as any, fallback as any);
export { meta };
