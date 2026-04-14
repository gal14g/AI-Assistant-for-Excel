/**
 * lookupAll – Find ALL matching rows between two ranges, not just the first.
 *
 * For each lookup value, finds every matching row in the source range and
 * collects values from the specified return column. Results are joined with
 * a delimiter and written to the output range.
 */

import { CapabilityMeta, LookupAllParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "lookupAll",
  description: "Find all matching rows between two ranges and return joined values",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: LookupAllParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { lookupRange, sourceRange, returnColumn, outputRange, delimiter = ", ", matchType = "exact" } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would look up all matches from ${lookupRange} in ${sourceRange}, return column ${returnColumn} to ${outputRange}`,
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
    return { stepId: "", status: "success", message: "No data to look up." };
  }

  const retColIdx = returnColumn - 1; // convert 1-based to 0-based

  options.onProgress?.(`Looking up ${lookupVals.length} values against ${sourceVals.length} source rows...`);

  // Build an index: source key (first column) → array of return column values
  const index = new Map<string, string[]>();
  for (const row of sourceVals) {
    const key = String(row[0] ?? "").trim().toLowerCase();
    if (!key) continue;
    const retVal = String(row[retColIdx] ?? "");
    if (!index.has(key)) index.set(key, []);
    index.get(key)!.push(retVal);
  }

  const results: (string | null)[][] = [];
  let matchedLookups = 0;
  let totalMatches = 0;

  for (const row of lookupVals) {
    const lookupStr = String(row[0] ?? "").trim().toLowerCase();
    if (!lookupStr) {
      results.push([null]);
      continue;
    }

    let matched: string[] = [];

    if (matchType === "exact") {
      matched = index.get(lookupStr) ?? [];
    } else {
      // contains: find all source keys that contain or are contained by the lookup value
      for (const [key, vals] of index) {
        if (key.includes(lookupStr) || lookupStr.includes(key)) {
          matched.push(...vals);
        }
      }
    }

    if (matched.length > 0) {
      results.push([matched.join(delimiter)]);
      matchedLookups++;
      totalMatches += matched.length;
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
    message: `Found matches for ${matchedLookups}/${lookupVals.length} lookup values (${totalMatches} total matches)`,
    outputs: { outputRange },
  };
}

registry.register(meta, handler as any);
export { meta };
