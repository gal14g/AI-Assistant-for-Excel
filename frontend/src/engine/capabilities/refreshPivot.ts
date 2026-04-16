/**
 * refreshPivot – Refresh a PivotTable or all PivotTables on a sheet.
 *
 * If pivotName is specified, refreshes that single pivot. Otherwise refreshes
 * all pivots on the specified sheet (or the active sheet).
 *
 * This is a non-mutating read-refresh operation — it recalculates from
 * existing source data without changing the workbook structure.
 */

import { CapabilityMeta, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";

const meta: CapabilityMeta = {
  action: "refreshPivot",
  description: "Refresh a PivotTable or all PivotTables on a sheet",
  mutates: false,
  affectsFormatting: false,
  // Worksheet.pivotTables + PivotTable.refresh() are ExcelApi 1.3+.
  // (Programmatic pivot creation/editing needs 1.8 — see createPivot.)
  requiresApiSet: "ExcelApi 1.3",
};

async function handler(
  context: Excel.RequestContext,
  params: any,
  options: ExecutionOptions
): Promise<StepResult> {
  const { pivotName, sheetName } = params;

  if (options.dryRun) {
    const target = pivotName
      ? `pivot "${pivotName}"`
      : sheetName
        ? `all pivots on "${sheetName}"`
        : "all pivots on active sheet";
    return {
      stepId: "",
      status: "success",
      message: `Would refresh ${target}`,
    };
  }

  if (pivotName) {
    // Refresh a specific pivot table by name
    options.onProgress?.(`Refreshing pivot "${pivotName}"...`);
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const pivot = sheet.pivotTables.getItem(pivotName);
    pivot.refresh();
    await context.sync();

    return {
      stepId: "",
      status: "success",
      message: `Refreshed pivot table "${pivotName}"`,
      outputs: { pivotName },
    };
  }

  // Refresh all pivots on the target sheet
  const sheet = sheetName
    ? context.workbook.worksheets.getItem(sheetName)
    : context.workbook.worksheets.getActiveWorksheet();

  const pivots = sheet.pivotTables;
  pivots.load("items/name");
  await context.sync();

  options.onProgress?.(`Refreshing ${pivots.items.length} pivot table(s)...`);

  for (const pivot of pivots.items) {
    pivot.refresh();
  }
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Refreshed ${pivots.items.length} pivot table(s)`,
    outputs: { pivotName: pivots.items.map(p => p.name).join(", ") },
  };
}

registry.register(meta, handler as any);
export { meta };
