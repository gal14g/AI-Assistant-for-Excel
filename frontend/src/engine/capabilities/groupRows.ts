/**
 * groupRows – Group or ungroup rows/columns for outline collapsing.
 *
 * Office.js API: Range.group() / Range.ungroup() (ExcelApi 1.10+)
 */

import { CapabilityMeta, GroupRowsParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";

const meta: CapabilityMeta = {
  action: "groupRows",
  description: "Group or ungroup rows/columns for outline collapsing",
  mutates: false,
  affectsFormatting: true,
  requiresApiSet: "ExcelApi 1.10",
};

async function handler(
  context: Excel.RequestContext,
  params: GroupRowsParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const verb = params.operation === "group" ? "Group" : "Ungroup";

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would ${verb.toLowerCase()} ${params.range}` };
  }

  options.onProgress?.(`${verb}ing ${params.range}...`);

  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange(params.range);

  // Determine if this is a row or column range
  // "3:8" → rows, "B:E" → columns
  const ref = params.range.includes("!") ? params.range.split("!")[1] : params.range;
  const isRowRange = /^\d+:\d+$/.test(ref);
  const groupOption = isRowRange
    ? ("ByRows" as Excel.GroupOption)
    : ("ByColumns" as Excel.GroupOption);

  if (params.operation === "group") {
    range.group(groupOption);
  } else {
    range.ungroup(groupOption);
  }

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `${verb}ed ${params.range}`,
    outputs: { range: params.range },
  };
}

// ── Legacy-Excel fallback (ExcelApi < 1.10) ──────────────────────────────────
// Range.group / Range.ungroup require 1.10. On older builds we degrade to the
// closest visible equivalent — hiding or unhiding the rows/columns directly.
// We lose the outline's collapse/expand chrome (the +/− buttons in the
// gutter), but the data-visibility effect is identical and users can still
// toggle visibility via the plan.
async function fallback(
  context: Excel.RequestContext,
  params: GroupRowsParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would ${params.operation === "group" ? "hide" : "unhide"} ${params.range} (legacy fallback for group outline).`,
    };
  }

  options.onProgress?.(
    `Legacy-Excel mode: ${params.operation === "group" ? "hiding" : "unhiding"} ${params.range} ` +
    `(outline grouping unavailable)...`,
  );

  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange(params.range);

  const ref = params.range.includes("!") ? params.range.split("!")[1] : params.range;
  const isRowRange = /^\d+:\d+$/.test(ref);

  const shouldHide = params.operation === "group";
  if (isRowRange) {
    range.rowHidden = shouldHide;
  } else {
    range.columnHidden = shouldHide;
  }
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message:
      `${shouldHide ? "Hid" : "Unhid"} ${params.range} — outline grouping (+/− gutter buttons) ` +
      `requires ExcelApi 1.10+, so hide/unhide was used instead (legacy-Excel fallback).`,
    outputs: { range: params.range },
  };
}

registry.register(meta, handler as any, fallback as any);
export { meta };
