/**
 * groupSum – Aggregate values by group using SUMIF or computed values.
 *
 * Strategy (preferFormula = true, default):
 *   Writes SUMIF/SUMIFS formulas so the results auto-update.
 *
 * Strategy (preferFormula = false):
 *   Reads data, computes group sums in JS, writes result values.
 */

import { CapabilityMeta, GroupSumParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "groupSum",
  description: "Sum values grouped by a column using SUMIF or computed aggregation",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: GroupSumParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const { dataRange, groupByColumn, sumColumn, outputRange, preferFormula = true, includeHeaders } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would compute grouped sums from ${dataRange}, output to ${outputRange}`,
    };
  }

  if (preferFormula) {
    return formulaGroupSum(context, params, options);
  }
  return computedGroupSum(context, params, options);
}

async function formulaGroupSum(
  context: Excel.RequestContext,
  params: GroupSumParams,
  options: ExecutionOptions
): Promise<StepResult> {
  options.onProgress?.("Reading data for group keys...");

  const dataRng = resolveRange(context, params.dataRange);
  dataRng.load("values");
  await context.sync();

  const data = dataRng.values;
  const startRow = params.includeHeaders ? 1 : 0;

  // Extract unique group keys
  const groupCol = params.groupByColumn - 1;
  const uniqueKeys: (string | number)[] = [];
  const seen = new Set<string>();
  for (let i = startRow; i < data.length; i++) {
    const key = String(data[i][groupCol]);
    if (!seen.has(key)) {
      seen.add(key);
      uniqueKeys.push(data[i][groupCol] as string | number);
    }
  }

  options.onProgress?.(`Found ${uniqueKeys.length} unique groups`);

  // Build output: group key column + SUMIF formula column
  const output: (string | number)[][] = [];
  if (params.includeHeaders) {
    output.push([String(data[0][groupCol]), `Sum of ${String(data[0][params.sumColumn - 1])}`]);
  }

  // Build column references for SUMIF
  const criteriaCol = getColumnFromRange(params.dataRange, groupCol);
  const sumCol = getColumnFromRange(params.dataRange, params.sumColumn - 1);

  for (const key of uniqueKeys) {
    const criteriaValue = typeof key === "string" ? `"${key}"` : key;
    output.push([key, `=SUMIF(${criteriaCol},${criteriaValue},${sumCol})` as any]);
  }

  const outRng = resolveRange(context, params.outputRange);
  // Write the group keys as values, formulas separately
  const values: (string | number | null)[][] = output.map((row) => [row[0], null]);
  const formulas: (string | null)[][] = output.map((row) => [null, String(row[1])]);

  // Write group keys
  const keyRange = outRng.getColumn(0).getResizedRange(output.length - 1, 0);
  keyRange.values = values.map((r) => [r[0]]);

  // Write SUMIF formulas
  const formulaRange = outRng.getColumn(1).getResizedRange(output.length - 1, 0);
  formulaRange.formulas = formulas.map((r) => [r[1] ?? ""]);

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Created ${uniqueKeys.length} SUMIF formulas in ${params.outputRange}`,
  };
}

async function computedGroupSum(
  context: Excel.RequestContext,
  params: GroupSumParams,
  options: ExecutionOptions
): Promise<StepResult> {
  options.onProgress?.("Reading data...");

  const dataRng = resolveRange(context, params.dataRange);
  dataRng.load("values");
  await context.sync();

  const data = dataRng.values;
  const startRow = params.includeHeaders ? 1 : 0;
  const groupCol = params.groupByColumn - 1;
  const sumCol = params.sumColumn - 1;

  options.onProgress?.("Computing group sums...");

  const groups = new Map<string, number>();
  for (let i = startRow; i < data.length; i++) {
    const key = String(data[i][groupCol]);
    const val = Number(data[i][sumCol]) || 0;
    groups.set(key, (groups.get(key) ?? 0) + val);
  }

  const output: (string | number)[][] = [];
  if (params.includeHeaders) {
    output.push([String(data[0][groupCol]), `Sum of ${String(data[0][sumCol])}`]);
  }
  for (const [key, sum] of groups) {
    output.push([key, sum]);
  }

  const outRng = resolveRange(context, params.outputRange);
  outRng.values = output;
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Computed ${groups.size} group sums, wrote to ${params.outputRange}`,
  };
}


function getColumnFromRange(rangeAddress: string, colOffset: number): string {
  const parts = rangeAddress.split("!");
  const ref = parts.length > 1 ? parts[1] : parts[0];
  const col = ref.match(/[A-Z]+/)?.[0] ?? "A";
  const offsetCol = offsetColumnLetter(col, colOffset);
  const prefix = parts.length > 1 ? parts[0] + "!" : "";
  return `${prefix}${offsetCol}:${offsetCol}`;
}

function offsetColumnLetter(col: string, offset: number): string {
  let num = 0;
  for (let i = 0; i < col.length; i++) {
    num = num * 26 + (col.charCodeAt(i) - 64);
  }
  num += offset;
  let result = "";
  while (num > 0) {
    const rem = (num - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    num = Math.floor((num - 1) / 26);
  }
  return result || "A";
}

registry.register(meta, handler as any);
export { meta };
