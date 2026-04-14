/**
 * regexReplace – Apply a regex find-and-replace across a range.
 *
 * Supports capture group references ($1, $2, etc.) in the replacement string.
 * Only modifies cells containing string values.
 */

import { CapabilityMeta, RegexReplaceParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "regexReplace",
  description: "Apply a regex find-and-replace across a range of cells",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: RegexReplaceParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { range: address, pattern, replacement, flags = "gi" } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would apply regex /${pattern}/${flags} replacement on ${address}`,
    };
  }

  options.onProgress?.("Reading range...");

  const rng = resolveRange(context, address);
  const used = rng.getUsedRange(false);
  used.load("values");
  await context.sync();

  const vals = (used.values ?? []) as (string | number | boolean | null)[][];
  if (!vals.length) {
    return { stepId: "", status: "success", message: "No data to process." };
  }

  let regex: RegExp;
  try {
    regex = new RegExp(pattern, flags);
  } catch (e) {
    return {
      stepId: "",
      status: "error",
      message: `Invalid regex pattern "/${pattern}/${flags}": ${e instanceof Error ? e.message : String(e)}`,
    };
  }

  options.onProgress?.(`Applying regex replacement across ${vals.length} rows...`);

  let replacedCells = 0;
  const newVals = vals.map((row) =>
    row.map((cell) => {
      if (typeof cell !== "string") return cell;
      // Reset lastIndex for global regexes
      regex.lastIndex = 0;
      if (!regex.test(cell)) return cell;
      regex.lastIndex = 0;
      replacedCells++;
      return cell.replace(regex, replacement);
    }),
  );

  // Write back modified values
  used.values = newVals;
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Replaced ${replacedCells} cells across ${vals.length} rows`,
    outputs: { range: address, replacementCount: replacedCells },
  };
}

registry.register(meta, handler as any);
export { meta };
