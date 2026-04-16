/**
 * setSheetDirection — request RTL / LTR sheet display.
 *
 * Office.js does NOT expose a Worksheet.rightToLeft (or equivalent)
 * property in any ExcelApi version up through 1.18+. This is a confirmed
 * Office.js API gap — the COM property Worksheet.DisplayRightToLeft exists
 * and works via VBA / xlwings / AppleScript, but is not reachable from a
 * sandboxed Office.js add-in.
 *
 * In add-in mode this handler returns a success-with-warning outcome so the
 * plan doesn't abort over something the user can toggle manually in one
 * click (View > Sheet Right-to-Left). When the MCP desktop bridge lands
 * (Item 4 of the productionization plan), the xlwings-side executor will
 * set the COM property directly and this handler becomes a pure pass-through.
 */

import { CapabilityMeta, SetSheetDirectionParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";

const meta: CapabilityMeta = {
  action: "setSheetDirection",
  description: "Request sheet right-to-left / left-to-right display (add-in: manual-toggle warning; MCP: sets COM property)",
  mutates: false,
  affectsFormatting: true,
};

async function handler(
  context: Excel.RequestContext,
  params: SetSheetDirectionParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { direction, sheetName } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would request ${direction.toUpperCase()} sheet direction (Office.js has no API; user must toggle manually).`,
    };
  }

  // Validate that the sheet exists, so we at least fail cleanly on typos.
  let resolvedName = sheetName;
  try {
    if (sheetName) {
      const ws = context.workbook.worksheets.getItemOrNullObject(sheetName);
      ws.load(["isNullObject", "name"]);
      await context.sync();
      if (ws.isNullObject) {
        return {
          stepId: "",
          status: "error",
          message: `Sheet "${sheetName}" not found.`,
        };
      }
      resolvedName = ws.name;
    } else {
      const active = context.workbook.worksheets.getActiveWorksheet();
      active.load("name");
      await context.sync();
      resolvedName = active.name;
    }
  } catch {
    // If sheet lookup fails we still return a soft warning — the user can
    // still apply the toggle manually.
  }

  const instruction =
    direction === "rtl"
      ? "Toggle View > Sheet Right-to-Left (or in Hebrew: תצוגה > גיליון מימין לשמאל)"
      : "Toggle View > Sheet Left-to-Right";

  return {
    stepId: "",
    status: "success",
    message:
      `Sheet direction change to ${direction.toUpperCase()} on "${resolvedName ?? "(active sheet)"}" is not supported by Office.js — ` +
      `the add-in cannot set this programmatically. ${instruction} to apply it manually. ` +
      `(This handler will set the property automatically when running in MCP/xlwings desktop mode.)`,
  };
}

registry.register(meta, handler as any);
export { meta };
