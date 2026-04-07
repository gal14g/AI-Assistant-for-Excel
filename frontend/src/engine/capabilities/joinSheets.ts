/**
 * joinSheets – SQL-style join between two ranges on key columns.
 *
 * Supports inner, left, right, and full outer joins.
 * Both ranges are expected to include header rows.
 * The result (combined headers + data) is written to outputRange.
 */

import { CapabilityMeta, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "joinSheets",
  description: "SQL-style join between two ranges on key columns (inner, left, right, full)",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: any,
  options: ExecutionOptions,
): Promise<StepResult> {
  const {
    leftRange,
    rightRange,
    leftKeyColumn,
    rightKeyColumn,
    joinType = "inner",
    outputRange,
  } = params as {
    leftRange: string;
    rightRange: string;
    leftKeyColumn: number;
    rightKeyColumn: number;
    joinType: "inner" | "left" | "right" | "full";
    outputRange: string;
  };

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would ${joinType} join ${leftRange} with ${rightRange} → ${outputRange}`,
    };
  }

  options.onProgress?.("Reading left range...");
  const leftRng = resolveRange(context, leftRange);
  leftRng.load("values");
  const rightRng = resolveRange(context, rightRange);
  rightRng.load("values");
  await context.sync();

  const leftData = (leftRng.values ?? []) as (string | number | boolean | null)[][];
  const rightData = (rightRng.values ?? []) as (string | number | boolean | null)[][];

  if (leftData.length < 1 || rightData.length < 1) {
    return { stepId: "", status: "success", message: "One or both ranges are empty." };
  }

  const lKeyIdx = leftKeyColumn - 1;
  const rKeyIdx = rightKeyColumn - 1;

  // Headers
  const leftHeaders = leftData[0];
  const rightHeaders = rightData[0];
  // Right headers excluding the key column to avoid duplication
  const rightHeadersFiltered = rightHeaders.filter((_, i) => i !== rKeyIdx);

  const combinedHeaders = [...leftHeaders, ...rightHeadersFiltered];

  options.onProgress?.("Building index on right range...");

  // Index right rows by key (allow multiple matches per key)
  const rightIndex = new Map<string, number[]>();
  for (let r = 1; r < rightData.length; r++) {
    const key = String(rightData[r][rKeyIdx] ?? "").toLowerCase();
    if (!rightIndex.has(key)) rightIndex.set(key, []);
    rightIndex.get(key)!.push(r);
  }

  options.onProgress?.("Joining rows...");

  const nullRight = new Array(rightHeadersFiltered.length).fill(null);
  const nullLeft = new Array(leftHeaders.length).fill(null);

  const resultRows: (string | number | boolean | null)[][] = [combinedHeaders];
  const matchedRightKeys = new Set<string>();

  // Process left rows
  for (let lr = 1; lr < leftData.length; lr++) {
    const key = String(leftData[lr][lKeyIdx] ?? "").toLowerCase();
    const rightMatches = rightIndex.get(key);

    if (rightMatches && rightMatches.length > 0) {
      matchedRightKeys.add(key);
      for (const rr of rightMatches) {
        const rightVals = rightData[rr].filter((_, i) => i !== rKeyIdx);
        resultRows.push([...leftData[lr], ...rightVals]);
      }
    } else if (joinType === "left" || joinType === "full") {
      resultRows.push([...leftData[lr], ...nullRight]);
    }
    // inner: skip unmatched left rows
  }

  // For right and full joins, add unmatched right rows
  if (joinType === "right" || joinType === "full") {
    for (let rr = 1; rr < rightData.length; rr++) {
      const key = String(rightData[rr][rKeyIdx] ?? "").toLowerCase();
      if (!matchedRightKeys.has(key)) {
        const rightVals = rightData[rr].filter((_, i) => i !== rKeyIdx);
        resultRows.push([...nullLeft, ...rightVals]);
      }
    }
  }

  options.onProgress?.("Writing results...");
  const outRng = resolveRange(context, outputRange);
  outRng.getResizedRange(resultRows.length - 1, resultRows[0].length - 1).values = resultRows as any;
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Joined ${leftData.length - 1} left rows with ${rightData.length - 1} right rows → ${resultRows.length - 1} result rows (${joinType})`,
  };
}

registry.register(meta, handler as any);
export { meta };
