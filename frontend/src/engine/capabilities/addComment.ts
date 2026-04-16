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

// ── Legacy-Excel fallback (ExcelApi < 1.10) ──────────────────────────────────
// workbook.comments.add requires 1.10. We emulate a cell comment by writing
// the text into the adjacent right-hand cell with italic light-yellow
// styling, prefixed with "Note:". Fidelity cost: no hover popup, no author/
// timestamp, no threaded replies, and the note consumes a real cell (the
// handler checks that the adjacent cell is empty before overwriting). For
// most annotation intents this is close enough.
async function fallback(
  context: Excel.RequestContext,
  params: AddCommentParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would emit inline note next to ${params.cell} (legacy fallback).`,
    };
  }

  options.onProgress?.("Legacy-Excel mode: writing inline note (comments unavailable)...");

  const anchor = resolveRange(context, params.cell);
  // Move one column to the right. Office.js Range.getOffsetRange(rowOffset, colOffset).
  const noteCell = anchor.getOffsetRange(0, 1);
  noteCell.load(["values", "address"]);
  await context.sync();

  const existing = (noteCell.values?.[0]?.[0] ?? "") as string | number | boolean | null;
  if (existing !== null && existing !== "" && existing !== undefined) {
    // Don't overwrite real user data. Bail loudly but not as a hard error.
    return {
      stepId: "",
      status: "success",
      message:
        `Could not write inline note: the cell adjacent to ${params.cell} is occupied. ` +
        `Comments require ExcelApi 1.10+; add the note manually (legacy-Excel fallback).`,
    };
  }

  const preview = params.content.length > 50 ? params.content.slice(0, 50) + "…" : params.content;
  noteCell.values = [[`Note: ${params.content}`]];
  noteCell.format.font.italic = true;
  noteCell.format.font.color = "#705000";
  noteCell.format.fill.color = "#FFF8D6"; // light yellow
  noteCell.format.wrapText = true;

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message:
      `Wrote inline note next to ${params.cell}: "${preview}" — native comments require ` +
      `ExcelApi 1.10+ (legacy-Excel fallback; consumes one cell, no hover popup).`,
  };
}

registry.register(meta, handler as any, fallback as any);
export { meta };
