/**
 * splitByGroup – Split a data range into separate sheets by unique values
 * in a column.
 *
 * For each unique value in the groupBy column, creates a new worksheet and
 * writes matching rows (optionally with headers) to it.
 */

import { CapabilityMeta, SplitByGroupParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "splitByGroup",
  description: "Split a data range into separate sheets by unique values in a column",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: SplitByGroupParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { dataRange, groupByColumn, keepHeaders = true } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would split ${dataRange} into separate sheets by column ${groupByColumn}`,
    };
  }

  options.onProgress?.("Reading data range...");

  const rng = resolveRange(context, dataRange);
  const used = rng.getUsedRange(false);
  used.load("values");
  await context.sync();

  const vals = (used.values ?? []) as (string | number | boolean | null)[][];
  if (!vals.length) {
    return { stepId: "", status: "success", message: "No data to split." };
  }

  const colIdx = groupByColumn - 1; // convert 1-based to 0-based
  const headerRow = vals[0];
  const dataRows = vals.slice(1);

  // Group rows by unique values in the target column
  const groups = new Map<string, (string | number | boolean | null)[][]>();
  for (const row of dataRows) {
    const key = String(row[colIdx] ?? "").trim();
    if (!key) continue;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key)!.push(row);
  }

  options.onProgress?.(`Creating ${groups.size} sheets...`);

  for (const [groupName, rows] of groups) {
    // Clean sheet name: remove invalid chars, truncate to 31 chars
    const cleanName = groupName
      .replace(/[:\\/?*[\]]/g, "_")
      .substring(0, 31);

    const newSheet = context.workbook.worksheets.add(cleanName);

    const writeRows: (string | number | boolean | null)[][] = [];
    if (keepHeaders) {
      writeRows.push(headerRow);
    }
    writeRows.push(...rows);

    if (writeRows.length > 0 && writeRows[0].length > 0) {
      const startCell = newSheet.getRange("A1");
      const targetRange = startCell.getResizedRange(writeRows.length - 1, writeRows[0].length - 1);
      targetRange.values = writeRows;
    }
  }

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Split into ${groups.size} sheets by column ${groupByColumn}`,
  };
}

registry.register(meta, handler as any);
export { meta };
