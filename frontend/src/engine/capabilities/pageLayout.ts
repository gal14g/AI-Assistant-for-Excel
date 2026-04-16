/**
 * pageLayout – Configure page layout settings for printing.
 *
 * Office.js notes:
 * - WorksheetPageLayout API is in ExcelApi 1.9+.
 * - Margins are set in inches.
 * - Orientation and paper size are enums.
 * - setPrintArea() defines the range that will be printed.
 */

import { CapabilityMeta, PageLayoutParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";

const meta: CapabilityMeta = {
  action: "pageLayout",
  description: "Configure page layout — margins, orientation, paper size, print area, gridlines",
  mutates: false,
  affectsFormatting: true,
  requiresApiSet: "ExcelApi 1.9",
};

async function handler(
  context: Excel.RequestContext,
  params: PageLayoutParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const { sheetName, margins, orientation, paperSize, printArea, showGridlines, printGridlines } = params;

  if (options.dryRun) {
    const parts: string[] = [];
    if (margins) parts.push("margins");
    if (orientation) parts.push(`orientation=${orientation}`);
    if (paperSize) parts.push(`paperSize=${paperSize}`);
    if (printArea) parts.push(`printArea=${printArea}`);
    if (showGridlines !== undefined) parts.push(`showGridlines=${showGridlines}`);
    if (printGridlines !== undefined) parts.push(`printGridlines=${printGridlines}`);
    return {
      stepId: "",
      status: "success",
      message: `Would set page layout: ${parts.join(", ")}`,
    };
  }

  options.onProgress?.("Configuring page layout...");

  // Resolve target sheet
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

  const layout = sheet.pageLayout;

  // Margins (in inches)
  if (margins) {
    if (margins.top !== undefined) layout.topMargin = margins.top * 72;
    if (margins.bottom !== undefined) layout.bottomMargin = margins.bottom * 72;
    if (margins.left !== undefined) layout.leftMargin = margins.left * 72;
    if (margins.right !== undefined) layout.rightMargin = margins.right * 72;
    if (margins.header !== undefined) layout.headerMargin = margins.header * 72;
    if (margins.footer !== undefined) layout.footerMargin = margins.footer * 72;
  }

  // Orientation
  if (orientation) {
    layout.orientation = orientation === "landscape"
      ? Excel.PageOrientation.landscape
      : Excel.PageOrientation.portrait;
  }

  // Paper size
  if (paperSize) {
    const sizeMap: Record<string, Excel.PaperType> = {
      letter: Excel.PaperType.letter,
      legal: Excel.PaperType.legal,
      a3: Excel.PaperType.a3,
      a4: Excel.PaperType.a4,
      a5: Excel.PaperType.a5,
      b4: Excel.PaperType.b4,
      b5: Excel.PaperType.b5,
      tabloid: Excel.PaperType.tabloid,
    };
    const paperType = sizeMap[paperSize.toLowerCase()];
    if (paperType !== undefined) {
      layout.paperSize = paperType;
    }
  }

  // Print area
  if (printArea) {
    layout.setPrintArea(printArea);
  }

  // Gridlines
  if (showGridlines !== undefined) {
    sheet.showGridlines = showGridlines;
  }
  if (printGridlines !== undefined) {
    layout.printGridlines = printGridlines;
  }

  await context.sync();

  const changes: string[] = [];
  if (margins) changes.push("margins");
  if (orientation) changes.push(orientation);
  if (paperSize) changes.push(paperSize);
  if (printArea) changes.push(`print area ${printArea}`);
  if (showGridlines !== undefined) changes.push(`gridlines ${showGridlines ? "on" : "off"}`);
  if (printGridlines !== undefined) changes.push(`print gridlines ${printGridlines ? "on" : "off"}`);

  return {
    stepId: "",
    status: "success",
    message: `Page layout updated: ${changes.join(", ")}${sheetName ? ` on "${sheetName}"` : ""}`,
  };
}

// ── Legacy-Excel fallback (ExcelApi < 1.9) ────────────────────────────────────
// WorksheetPageLayout is 1.9+. Below that, Office.js offers no programmatic
// handle on margins, orientation, paper size, or print area. `showGridlines`
// and `sheet.getRange(...).printArea` don't exist on 1.3 either. The correct
// behavior is a graceful skip — page layout is print-only and doesn't affect
// on-screen data; the plan should keep running. We emit a descriptive
// success with a warning so the audit trail shows what was skipped.
async function fallback(
  _context: Excel.RequestContext,
  params: PageLayoutParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: "Would skip page-layout (legacy fallback; requires ExcelApi 1.9+).",
    };
  }

  options.onProgress?.("Legacy-Excel mode: page layout unavailable, skipping...");

  const requested: string[] = [];
  if (params.margins) requested.push("margins");
  if (params.orientation) requested.push(`orientation=${params.orientation}`);
  if (params.paperSize) requested.push(`paperSize=${params.paperSize}`);
  if (params.printArea) requested.push(`printArea=${params.printArea}`);
  if (params.showGridlines !== undefined) requested.push(`showGridlines=${params.showGridlines}`);
  if (params.printGridlines !== undefined) requested.push(`printGridlines=${params.printGridlines}`);

  return {
    stepId: "",
    status: "success",
    message:
      `Page layout update skipped — WorksheetPageLayout requires ExcelApi 1.9+. ` +
      `Requested: ${requested.join(", ") || "(nothing)"}. Configure manually via File › Page Setup ` +
      `(legacy-Excel fallback).`,
  };
}

registry.register(meta, handler as any, fallback as any);
export { meta };
