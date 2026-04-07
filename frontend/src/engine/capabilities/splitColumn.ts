/**
 * splitColumn – Split a text column into multiple columns by a delimiter.
 *
 * Common uses:
 *   - "First Last" → First Name / Last Name
 *   - "City, State" → City / State
 *   - "2024-01-15" → Year / Month / Day
 */

import { CapabilityMeta, SplitColumnParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange, resolveSheet, stripWorkbookQualifier } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "splitColumn",
  description: "Split a text column into multiple columns using a delimiter",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: SplitColumnParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { sourceRange, delimiter, outputStartColumn, parts = 2 } = params;

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would split ${sourceRange} by "${delimiter}" into ${parts} columns starting at ${outputStartColumn}` };
  }

  options.onProgress?.("Reading source column...");

  const ws = resolveSheet(context, sourceRange);
  const srcRng = resolveRange(context, sourceRange);
  const used = srcRng.getUsedRange(false);
  used.load("values, address");
  await context.sync();

  const vals = (used.values ?? []) as (string | number | boolean | null)[][];
  if (!vals.length) return { stepId: "", status: "success", message: "No data to split." };

  // Detect start row from address (e.g. "$A$2:$A$50" → 2)
  const cellPart = used.address.includes("!") ? used.address.split("!").pop()! : used.address;
  const startRow = parseInt((cellPart.replace(/\$/g, "").match(/[A-Z]+(\d+)/) ?? ["", "1"])[1], 10);

  options.onProgress?.(`Splitting ${vals.length} rows...`);

  const colToIndex = (col: string): number => {
    let n = 0;
    for (const c of col.toUpperCase()) n = n * 26 + (c.charCodeAt(0) - 64);
    return n - 1;
  };
  const indexToCol = (idx: number): string => {
    let s = "";
    idx++;
    while (idx > 0) { const r = (idx - 1) % 26; s = String.fromCharCode(65 + r) + s; idx = Math.floor((idx - 1) / 26); }
    return s;
  };

  const startColIdx = colToIndex(outputStartColumn);

  // Build the entire output grid in memory, then write as a single batch.
  // This avoids per-cell writes which are slow and can fail on merged cells.
  const outGrid: (string | number | boolean | null)[][] = [];

  for (let i = 0; i < vals.length; i++) {
    const cell = String(vals[i][0] ?? "");
    const splitParts = cell.split(delimiter).slice(0, parts);
    while (splitParts.length < parts) splitParts.push("");
    outGrid.push(splitParts.map((p) => p.trim()));
  }

  // Overwrite with headers if provided
  if (params.outputHeaders?.length && outGrid.length > 0) {
    for (let c = 0; c < Math.min(params.outputHeaders.length, parts); c++) {
      outGrid[0][c] = params.outputHeaders[c];
    }
  }

  // Single batch write — much faster and safer than per-cell writes
  const outAddr = `${indexToCol(startColIdx)}${startRow}:${indexToCol(startColIdx + parts - 1)}${startRow + outGrid.length - 1}`;
  try {
    ws.getRange(outAddr).values = outGrid as any;
    await context.sync();
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return { stepId: "", status: "error", message: `Failed to write split results: ${msg}. Range may contain merged or protected cells.` };
  }

  // Auto-fit the new columns
  const addr = stripWorkbookQualifier(sourceRange);
  const sheetPrefix = addr.includes("!") ? addr.split("!")[0] + "!" : "";
  try {
    const outRange = resolveRange(context, `${sheetPrefix}${outputStartColumn}${startRow}:${indexToCol(startColIdx + parts - 1)}${startRow + vals.length - 1}`);
    outRange.format.autofitColumns();
    await context.sync();
  } catch { /* non-fatal */ }

  return {
    stepId: "",
    status: "success",
    message: `Split ${vals.length} rows from ${sourceRange} into ${parts} columns starting at ${outputStartColumn}`,
  };
}

registry.register(meta, handler as any);
export { meta };
