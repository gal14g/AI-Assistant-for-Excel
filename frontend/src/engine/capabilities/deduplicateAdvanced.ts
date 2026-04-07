/**
 * deduplicateAdvanced – Remove duplicates with a strategy for which row to keep.
 *
 * Unlike the basic removeDuplicates (Office.js built-in), this handler supports
 * choosing WHICH duplicate to keep:
 *   "first"        – keep the first occurrence
 *   "last"         – keep the last occurrence
 *   "mostComplete" – keep the row with fewest blank cells
 *   "newest"       – keep the row with the latest value in dateColumn
 */

import { CapabilityMeta, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "deduplicateAdvanced",
  description: "Remove duplicate rows with a strategy for which row to keep (first, last, mostComplete, newest)",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: any,
  options: ExecutionOptions,
): Promise<StepResult> {
  const {
    range,
    keyColumns,
    keepStrategy = "first",
    dateColumn,
    hasHeaders = true,
  } = params as {
    range: string;
    keyColumns: number[];
    keepStrategy: "first" | "last" | "mostComplete" | "newest";
    dateColumn?: number;
    hasHeaders?: boolean;
  };

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would deduplicate ${range} on columns [${keyColumns.join(", ")}] keeping "${keepStrategy}"`,
    };
  }

  options.onProgress?.("Reading range...");
  const rng = resolveRange(context, range);
  rng.load("values");
  await context.sync();

  const allVals = (rng.values ?? []) as (string | number | boolean | null)[][];
  if (!allVals.length) return { stepId: "", status: "success", message: "No data found." };

  const headerRow = hasHeaders ? allVals[0] : null;
  const dataRows = hasHeaders ? allVals.slice(1) : allVals;

  options.onProgress?.("Grouping rows by key...");

  // Group rows by composite key
  const groups = new Map<string, number[]>();
  for (let i = 0; i < dataRows.length; i++) {
    const key = keyColumns.map((col) => String(dataRows[i][col - 1] ?? "").toLowerCase()).join("\x00");
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key)!.push(i);
  }

  options.onProgress?.("Selecting rows to keep...");

  // For each group, pick which row to keep
  const keepIndices = new Set<number>();

  for (const [, indices] of groups) {
    let pickIdx: number;
    switch (keepStrategy) {
      case "first":
        pickIdx = indices[0];
        break;
      case "last":
        pickIdx = indices[indices.length - 1];
        break;
      case "mostComplete": {
        let bestIdx = indices[0];
        let fewestBlanks = Infinity;
        for (const idx of indices) {
          const blanks = dataRows[idx].filter((v) => v === null || v === "").length;
          if (blanks < fewestBlanks) {
            fewestBlanks = blanks;
            bestIdx = idx;
          }
        }
        pickIdx = bestIdx;
        break;
      }
      case "newest": {
        const dc = (dateColumn ?? 1) - 1;
        let bestIdx = indices[0];
        let latestTime = -Infinity;
        for (const idx of indices) {
          const val = dataRows[idx][dc];
          const t = typeof val === "number"
            ? (val - 25569) * 86400000 // Excel serial
            : Date.parse(String(val ?? ""));
          if (!isNaN(t) && t > latestTime) {
            latestTime = t;
            bestIdx = idx;
          }
        }
        pickIdx = bestIdx;
        break;
      }
      default:
        pickIdx = indices[0];
    }
    keepIndices.add(pickIdx);
  }

  // Build output
  const keptRows = dataRows.filter((_, i) => keepIndices.has(i));
  const output: (string | number | boolean | null)[][] = headerRow
    ? [headerRow, ...keptRows]
    : keptRows;

  // Pad remaining rows with nulls to clear old data
  const totalRows = allVals.length;
  const cols = allVals[0].length;
  while (output.length < totalRows) {
    output.push(new Array(cols).fill(null));
  }

  options.onProgress?.("Writing deduplicated data...");
  rng.values = output as any;
  await context.sync();

  const removed = dataRows.length - keptRows.length;
  return {
    stepId: "",
    status: "success",
    message: `Removed ${removed} duplicate rows, kept ${keptRows.length} unique rows (strategy: ${keepStrategy})`,
  };
}

registry.register(meta, handler as any);
export { meta };
