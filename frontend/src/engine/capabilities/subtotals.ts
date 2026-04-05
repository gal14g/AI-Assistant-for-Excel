/**
 * subtotals – Insert subtotal rows at each group boundary.
 *
 * Sorts data by the group column, then inserts a subtotal row after each group.
 * Works on the existing data in-place by re-writing the range with subtotal rows inserted.
 */

import { CapabilityMeta, SubtotalsParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "subtotals",
  description: "Insert subtotal rows at each group boundary in a data range",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: SubtotalsParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { dataRange, groupByColumn, subtotalColumns, aggregation = "sum", subtotalLabel = "Total" } = params;

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would add ${aggregation} subtotals to ${dataRange} grouped by column ${groupByColumn}` };
  }

  options.onProgress?.("Reading data...");
  const rng = resolveRange(context, dataRange);
  rng.load("values, address");
  await context.sync();

  const vals = (rng.values ?? []) as (string | number | boolean | null)[][];
  if (vals.length < 2) return { stepId: "", status: "success", message: "Not enough rows." };

  const headerRow = vals[0];
  const dataRows = vals.slice(1);
  const grpIdx = groupByColumn - 1;
  const subCols = subtotalColumns.map((c) => c - 1);

  options.onProgress?.("Sorting by group column...");
  dataRows.sort((a, b) => String(a[grpIdx] ?? "").localeCompare(String(b[grpIdx] ?? "")));

  // Build output with subtotal rows inserted
  const out: (string | number | boolean | null)[][] = [headerRow];
  let currentGroup = String(dataRows[0]?.[grpIdx] ?? "");
  let groupRows: (string | number | boolean | null)[][] = [];

  const flush = () => {
    out.push(...groupRows);
    // Build subtotal row
    const subRow: (string | number | boolean | null)[] = headerRow.map(() => null);
    subRow[grpIdx] = `${currentGroup} ${subtotalLabel}`;
    for (const ci of subCols) {
      const nums = groupRows.map((r) => Number(r[ci]) || 0);
      if (aggregation === "sum") subRow[ci] = nums.reduce((a, b) => a + b, 0);
      else if (aggregation === "count") subRow[ci] = nums.length;
      else subRow[ci] = nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
    }
    out.push(subRow);
    groupRows = [];
  };

  for (const row of dataRows) {
    const grp = String(row[grpIdx] ?? "");
    if (grp !== currentGroup) {
      flush();
      currentGroup = grp;
    }
    groupRows.push(row);
  }
  flush(); // last group

  // Write back — may need to expand the range
  const cellPart = rng.address.includes("!") ? rng.address.split("!").pop()! : rng.address;
  const startRow = parseInt((cellPart.replace(/\$/g, "").match(/[A-Z]+(\d+)/) ?? ["", "1"])[1], 10);
  const startColMatch = cellPart.replace(/\$/g, "").match(/([A-Z]+)\d+/);
  const startCol = startColMatch ? startColMatch[1] : "A";
  const ws = rng.worksheet;

  // Insert blank rows to make room if needed (out.length > vals.length)
  const extraRows = out.length - vals.length;
  if (extraRows > 0) {
    ws.getRange(`${startRow + vals.length}:${startRow + vals.length + extraRows - 1}`).insert(Excel.InsertShiftDirection.down);
  }

  ws.getRange(`${startCol}${startRow}`).getResizedRange(out.length - 1, (out[0]?.length ?? 1) - 1).values = out as any;
  await context.sync();

  const subtotalCount = out.filter((r) => String(r[grpIdx]).endsWith(subtotalLabel)).length;
  return {
    stepId: "",
    status: "success",
    message: `Added ${subtotalCount} subtotal rows to ${dataRange}`,
  };
}

registry.register(meta, handler as any);
export { meta };
