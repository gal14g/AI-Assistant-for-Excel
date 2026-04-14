/**
 * categorize – Label rows based on ordered if/else rules.
 *
 * Rules are evaluated top-to-bottom; the first matching rule's label is written.
 * Supports: contains, equals, startsWith, endsWith, greaterThan, lessThan, regex
 */

import { CapabilityMeta, CategorizeParams, CategorizeRule, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "categorize",
  description: "Label cells based on ordered if/else rules (contains, equals, regex, numeric comparisons)",
  mutates: true,
  affectsFormatting: false,
};

function applyRule(cell: string | number | boolean | null, rule: CategorizeRule): boolean {
  const strCell = String(cell ?? "").toLowerCase();
  const strVal  = String(rule.value).toLowerCase();
  const numCell = Number(cell);
  const numVal  = Number(rule.value);

  switch (rule.operator) {
    case "contains":    return strCell.includes(strVal);
    case "equals":      return strCell === strVal;
    case "startsWith":  return strCell.startsWith(strVal);
    case "endsWith":    return strCell.endsWith(strVal);
    case "greaterThan": return !isNaN(numCell) && !isNaN(numVal) && numCell > numVal;
    case "lessThan":    return !isNaN(numCell) && !isNaN(numVal) && numCell < numVal;
    case "regex": {
      try { return new RegExp(String(rule.value), "i").test(String(cell ?? "")); }
      catch { return false; }
    }
    default: return false;
  }
}

async function handler(
  context: Excel.RequestContext,
  params: CategorizeParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { sourceRange, outputRange, rules, defaultValue = "" } = params;

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would categorize ${sourceRange} with ${rules.length} rules` };
  }

  options.onProgress?.("Reading source range...");
  const srcRng = resolveRange(context, sourceRange);
  const used = srcRng.getUsedRange(false);
  used.load("values");
  await context.sync();

  const vals = (used.values ?? []) as (string | number | boolean | null)[][];
  options.onProgress?.(`Applying ${rules.length} rules to ${vals.length} rows...`);

  const labelCounts: Record<string, number> = {};
  const out: (string | number)[][] = vals.map((row) =>
    row.map((cell) => {
      for (const rule of rules) {
        if (applyRule(cell, rule)) {
          labelCounts[rule.label] = (labelCounts[rule.label] ?? 0) + 1;
          return rule.label;
        }
      }
      return defaultValue;
    }),
  );

  const outRng = resolveRange(context, outputRange);
  outRng.getResizedRange(out.length - 1, out[0].length - 1).values = out as any;
  await context.sync();

  const summary = Object.entries(labelCounts).map(([k, v]) => `${k}:${v}`).join(", ");
  return {
    stepId: "",
    status: "success",
    message: `Categorized ${vals.flat().length} cells — ${summary || "no matches"}`,
    outputs: { outputRange },
  };
}

registry.register(meta, handler as any);
export { meta };
