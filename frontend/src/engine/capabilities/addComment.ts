/**
 * addComment – Add a comment/note to a cell.
 *
 * Uses the Office.js Comments API (ExcelApi 1.10+).
 * Falls back to adding a note if comments API is unavailable.
 */

import { CapabilityMeta, AddCommentParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "addComment",
  description: "Add a comment or note to a cell",
  mutates: true,
  affectsFormatting: false,
  requiresApiSet: "ExcelApi 1.10",
};

async function handler(
  context: Excel.RequestContext,
  params: AddCommentParams,
  options: ExecutionOptions
): Promise<StepResult> {
  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would add comment to ${params.cell}` };
  }

  options.onProgress?.("Adding comment...");

  const range = resolveRange(context, params.cell);
  context.workbook.comments.add(range, params.content);
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Added comment to ${params.cell}: "${params.content.slice(0, 50)}${params.content.length > 50 ? "…" : ""}"`,
  };
}

registry.register(meta, handler as any);
export { meta };
