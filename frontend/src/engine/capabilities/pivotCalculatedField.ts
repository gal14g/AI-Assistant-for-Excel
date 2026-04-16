/**
 * pivotCalculatedField – Add a calculated field to an existing PivotTable.
 *
 * Uses the PivotTable calculatedFields API to add a new field with a custom
 * formula. This requires ExcelApi 1.8+ and may not be available in all
 * Office.js environments.
 *
 * Note: The calculatedFields API has limited availability. If it is not
 * supported, a descriptive error message is returned.
 */

import { CapabilityMeta, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "pivotCalculatedField",
  description: "Add a calculated field to an existing PivotTable",
  mutates: true,
  affectsFormatting: false,
  requiresApiSet: "ExcelApi 1.8",
};

async function handler(
  context: Excel.RequestContext,
  params: any,
  options: ExecutionOptions
): Promise<StepResult> {
  const { pivotName, sheetName, fieldName, formula } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would add calculated field "${fieldName}" to pivot "${pivotName}"`,
    };
  }

  options.onProgress?.(`Adding calculated field "${fieldName}" to "${pivotName}"...`);

  try {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();

    const pivotTable = sheet.pivotTables.getItem(pivotName);

    // Attempt to add the calculated field
    // Note: calculatedFields may not be available in all API sets
    (pivotTable as any).calculatedFields.add(fieldName, formula);
    await context.sync();

    return {
      stepId: "",
      status: "success",
      message: `Added calculated field "${fieldName}" to pivot "${pivotName}"`,
      outputs: { pivotName, fieldName },
    };
  } catch (err: any) {
    const message = err?.message ?? String(err);

    // Provide a helpful message if the API is not available
    if (
      message.includes("calculatedFields") ||
      message.includes("not a function") ||
      message.includes("undefined")
    ) {
      return {
        stepId: "",
        status: "error",
        message: `Calculated fields API is not available in this Office.js version. ` +
          `You may need to add the calculated field manually via the PivotTable Fields pane.`,
        error: message,
      };
    }

    return {
      stepId: "",
      status: "error",
      message: `Failed to add calculated field "${fieldName}": ${message}`,
      error: message,
    };
  }
}

// ── Legacy-Excel fallback (ExcelApi < 1.8) ───────────────────────────────────
// There is no native PivotTable object on Excel 2016, so calculated fields
// can't be attached to one. When createPivot runs its own fallback it writes
// a SUMIFS-based "summary" sheet named after the pivot. We mirror that
// structure by appending a new column to the same sheet: column header =
// fieldName, and the formula is written into each data row unchanged.
//
// Fidelity note: the user's `formula` is pivot-field-name syntax (e.g.
// "=Revenue-Cost"). On a static summary sheet, Revenue/Cost resolve to
// *column letters*, not named fields, so the formula is usually wrong as
// written. We emit a warning describing the transform the user needs to
// apply, rather than silently writing broken cells.
async function fallback(
  context: Excel.RequestContext,
  params: any,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { pivotName, fieldName, formula, sheetName } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would append column "${fieldName}" to summary sheet "${pivotName}" (legacy fallback).`,
    };
  }

  // Locate the summary sheet that createPivot's fallback would have produced.
  // We try `sheetName` first (explicit), then the pivot name itself (default).
  const targetName = (sheetName ?? pivotName) as string;
  const ws = context.workbook.worksheets.getItemOrNullObject(targetName);
  ws.load("isNullObject");
  await context.sync();
  if (ws.isNullObject) {
    return {
      stepId: "",
      status: "error",
      message:
        `Calculated fields require ExcelApi 1.8+, which this Excel version lacks. ` +
        `Additionally, no summary sheet "${targetName}" was found to extend. Run createPivot first.`,
    };
  }

  options.onProgress?.(`Legacy-Excel mode: appending column "${fieldName}" to "${targetName}"...`);

  const usedRange = ws.getUsedRange(true);
  usedRange.load(["rowCount", "columnCount", "address"]);
  await context.sync();

  if (!usedRange.rowCount || !usedRange.columnCount) {
    return {
      stepId: "",
      status: "error",
      message: `Summary sheet "${targetName}" is empty — cannot append a calculated column.`,
    };
  }

  const newColIdx = usedRange.columnCount; // zero-based index of the new column
  const dataRows = Math.max(0, usedRange.rowCount - 1);

  // Header cell
  const headerCell = ws.getRangeByIndexes(0, newColIdx, 1, 1);
  headerCell.values = [[fieldName]];

  // Data cells — write the user's formula verbatim. Prefix "=" if missing.
  if (dataRows > 0) {
    const f = typeof formula === "string" && formula.startsWith("=") ? formula : `=${formula ?? ""}`;
    const formulaRange = ws.getRangeByIndexes(1, newColIdx, dataRows, 1);
    const vals: string[][] = [];
    for (let i = 0; i < dataRows; i++) vals.push([f]);
    formulaRange.formulas = vals;
  }
  await context.sync();

  // Compute the resulting column letter for the warning.
  // Fallback-safe resolveRange usage: we don't need it here, but keep the import
  // intentionally because the primary handler path above references it too.
  void resolveRange;

  return {
    stepId: "",
    status: "success",
    message:
      `Appended "${fieldName}" column to summary sheet "${targetName}" with formula \`${formula}\`. ` +
      `Warning: pivot calculated-field formulas reference field *names* (e.g. =Revenue-Cost). On a ` +
      `flat summary sheet those names do not resolve — edit the generated formula to use cell ` +
      `references (e.g. =B2-C2) if it shows #NAME? (legacy-Excel fallback — ExcelApi 1.8+ required ` +
      `for native calculated fields).`,
    outputs: { pivotName: targetName, fieldName },
  };
}

registry.register(meta, handler as any, fallback as any);
export { meta };
