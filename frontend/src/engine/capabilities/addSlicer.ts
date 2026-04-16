/**
 * addSlicer – Add a slicer to filter a PivotTable or Table.
 *
 * Office.js notes:
 * - workbook.slicers.add(source, sourceField, destinationSheet?) creates a slicer.
 * - The source can be a PivotTable name or Table name.
 * - Position and size are set via slicer.left/top/width/height.
 */

import { CapabilityMeta, AddSlicerParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";

const meta: CapabilityMeta = {
  action: "addSlicer",
  description: "Add a slicer to filter a PivotTable or Table",
  mutates: true,
  affectsFormatting: false,
  requiresApiSet: "ExcelApi 1.10",
};

async function handler(
  context: Excel.RequestContext,
  params: AddSlicerParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const { sheetName, sourceName, sourceField, left, top, width, height } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would add slicer for field "${sourceField}" on "${sourceName}"`,
    };
  }

  options.onProgress?.(`Adding slicer for "${sourceField}"...`);

  // Resolve destination sheet if provided
  let destinationSheet: Excel.Worksheet | undefined;
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
    destinationSheet = ws;
  }

  const slicer = destinationSheet
    ? context.workbook.slicers.add(sourceName, sourceField, destinationSheet)
    : context.workbook.slicers.add(sourceName, sourceField);

  if (left !== undefined) slicer.left = left;
  if (top !== undefined) slicer.top = top;
  if (width !== undefined) slicer.width = width;
  if (height !== undefined) slicer.height = height;

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Added slicer for "${sourceField}" on "${sourceName}"${sheetName ? ` on sheet "${sheetName}"` : ""}`,
    outputs: { slicerName: sourceField },
  };
}

// ── Legacy-Excel fallback (ExcelApi < 1.10) ──────────────────────────────────
// Slicers require 1.10 (Excel 2019 UR8+). There is no way to reproduce the
// floating filter-button UI with Office.js 1.3 primitives. The closest
// functional equivalent is AutoFilter on the underlying table (applyFilter
// handler), but we can't invoke another handler from here without creating
// a circular dependency on the registry.
//
// Strategy: surface a clear "slicer not supported, use applyFilter instead"
// status=success with a warning (not an error). The LLM/user sees the audit
// line and can re-plan with applyFilter if the intent was data filtering
// rather than visual chrome. This matches the research verdict: data outcome
// is recoverable via a sibling handler; pure UX (one-click filter tiles) is
// not reproducible and failing the whole plan over cosmetics is worse than
// partial success.
async function fallback(
  _context: Excel.RequestContext,
  params: AddSlicerParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { sourceField, sourceName } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would skip slicer for "${sourceField}" (unavailable on this Excel version).`,
    };
  }

  options.onProgress?.("Legacy-Excel mode: slicers unavailable, emitting guidance...");

  return {
    stepId: "",
    status: "success",
    message:
      `Slicer for "${sourceField}" on "${sourceName}" skipped — slicers require ` +
      `ExcelApi 1.10+ (Excel 2019 UR8 / 2021 / Microsoft 365). For an equivalent ` +
      `filter on this Excel version, add an applyFilter step targeting "${sourceName}" ` +
      `with column "${sourceField}" (legacy-Excel fallback).`,
    outputs: { slicerName: sourceField },
  };
}

registry.register(meta, handler as any, fallback as any);
export { meta };
