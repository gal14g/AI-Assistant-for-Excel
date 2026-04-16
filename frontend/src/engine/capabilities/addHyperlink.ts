/**
 * addHyperlink – Insert a hyperlink in a cell.
 */

import { CapabilityMeta, AddHyperlinkParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "addHyperlink",
  description: "Insert a hyperlink in a cell",
  mutates: true,
  affectsFormatting: false,
  // Range.hyperlink was introduced in ExcelApi 1.7 (Excel 2019+).
  requiresApiSet: "ExcelApi 1.7",
};

async function handler(
  context: Excel.RequestContext,
  params: AddHyperlinkParams,
  options: ExecutionOptions
): Promise<StepResult> {
  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would add hyperlink to ${params.cell}` };
  }

  options.onProgress?.("Inserting hyperlink...");

  const range = resolveRange(context, params.cell);
  range.hyperlink = {
    address: params.url,
    textToDisplay: params.displayText ?? params.url,
  };
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Added hyperlink to ${params.cell}: ${params.url}`,
  };
}

// ── Legacy-Excel fallback (ExcelApi < 1.7) ────────────────────────────────────
// Range.hyperlink requires 1.7. The `=HYPERLINK(url, display)` worksheet function
// exists in every Excel version back to 97 and produces a clickable cell that
// navigates to the URL — visually and functionally identical to what the user
// asked for. Escape any embedded double-quotes in the url/display text so the
// formula stays well-formed even for odd URLs.
async function fallback(
  context: Excel.RequestContext,
  params: AddHyperlinkParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would add HYPERLINK() formula to ${params.cell}` };
  }

  options.onProgress?.("Legacy-Excel mode: writing HYPERLINK() formula instead of Range.hyperlink...");

  const escape = (s: string) => s.replace(/"/g, '""');
  const display = params.displayText ?? params.url;
  const formula = `=HYPERLINK("${escape(params.url)}","${escape(display)}")`;

  const range = resolveRange(context, params.cell);
  range.formulas = [[formula]];
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message:
      `Added HYPERLINK() formula to ${params.cell}: ${params.url} ` +
      `(legacy-Excel fallback — Range.hyperlink requires ExcelApi 1.7+).`,
  };
}

registry.register(meta, handler as any, fallback as any);
export { meta };
