/**
 * fuzzyMatch – Fuzzy string matching between two columns using Levenshtein distance.
 *
 * For each value in the lookup range, finds the best match in the source range
 * using similarity scoring. Writes the matched source value (or a constant
 * writeValue) to the output column when the similarity exceeds the threshold.
 */

import { CapabilityMeta, FuzzyMatchParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "fuzzyMatch",
  description: "Fuzzy string matching between two columns using Levenshtein distance",
  mutates: true,
  affectsFormatting: false,
};

function levenshtein(a: string, b: string): number {
  const m = a.length, n = b.length;
  const dp: number[][] = Array.from({ length: m + 1 }, (_, i) =>
    Array.from({ length: n + 1 }, (_, j) => i === 0 ? j : j === 0 ? i : 0),
  );
  for (let i = 1; i <= m; i++)
    for (let j = 1; j <= n; j++)
      dp[i][j] = a[i - 1] === b[j - 1] ? dp[i - 1][j - 1] : 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]);
  return dp[m][n];
}

function similarity(a: string, b: string): number {
  if (a === b) return 1;
  const maxLen = Math.max(a.length, b.length);
  if (maxLen === 0) return 1;
  return 1 - levenshtein(a, b) / maxLen;
}

async function handler(
  context: Excel.RequestContext,
  params: FuzzyMatchParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { lookupRange, sourceRange, outputRange, threshold = 0.7, writeValue, returnBestMatch = false } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would fuzzy match ${lookupRange} against ${sourceRange} (threshold ${Math.round(threshold * 100)}%), output to ${outputRange}`,
    };
  }

  options.onProgress?.("Reading lookup and source ranges...");

  const lookupRng = resolveRange(context, lookupRange).getUsedRange(false);
  const sourceRng = resolveRange(context, sourceRange).getUsedRange(false);
  lookupRng.load("values");
  sourceRng.load("values");
  await context.sync();

  const lookupVals = (lookupRng.values ?? []) as (string | number | boolean | null)[][];
  const sourceVals = (sourceRng.values ?? []) as (string | number | boolean | null)[][];

  if (!lookupVals.length || !sourceVals.length) {
    return { stepId: "", status: "success", message: "No data to match." };
  }

  // Flatten source values (first column)
  const sourceStrings = sourceVals.map((row) => String(row[0] ?? "").trim().toLowerCase());

  options.onProgress?.(`Fuzzy matching ${lookupVals.length} rows against ${sourceStrings.length} source values...`);

  const results: (string | null)[][] = [];
  let matchCount = 0;

  for (const row of lookupVals) {
    const lookupStr = String(row[0] ?? "").trim().toLowerCase();
    if (!lookupStr) {
      results.push([null]);
      continue;
    }

    let bestScore = 0;
    let bestIdx = -1;

    for (let j = 0; j < sourceStrings.length; j++) {
      if (!sourceStrings[j]) continue;
      const score = similarity(lookupStr, sourceStrings[j]);
      if (score > bestScore) {
        bestScore = score;
        bestIdx = j;
      }
    }

    if (bestScore >= threshold && bestIdx >= 0) {
      const output = writeValue !== undefined
        ? writeValue
        : returnBestMatch
          ? String(sourceVals[bestIdx][0] ?? "")
          : String(sourceVals[bestIdx][0] ?? "");
      results.push([output]);
      matchCount++;
    } else {
      results.push([null]);
    }
  }

  // Write results to output range
  const outRng = resolveRange(context, outputRange);
  outRng.getResizedRange(results.length - 1, 0).values = results as any;
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Fuzzy matched ${matchCount}/${lookupVals.length} rows (threshold ${Math.round(threshold * 100)}%)`,
  };
}

registry.register(meta, handler as any);
export { meta };
