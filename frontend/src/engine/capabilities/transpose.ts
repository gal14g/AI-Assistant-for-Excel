/**
 * transpose – Flip rows and columns of a range.
 *
 * Reads the source range, transposes the 2D array, writes to outputRange.
 * Values-only by default (formatting not copied).
 */

import { CapabilityMeta, TransposeParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "transpose",
  description: "Transpose a range — flip rows and columns",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: TransposeParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { sourceRange, outputRange } = params;

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would transpose ${sourceRange} → ${outputRange}` };
  }

  options.onProgress?.("Reading source range...");
  const srcRng = resolveRange(context, sourceRange);
  srcRng.load("values, rowCount, columnCount");
  await context.sync();

  const vals = (srcRng.values ?? []) as (string | number | boolean | null)[][];
  if (!vals.length) return { stepId: "", status: "success", message: "No data to transpose." };

  const rows = vals.length;
  const cols = vals[0]?.length ?? 0;
  options.onProgress?.(`Transposing ${rows} × ${cols}...`);

  // Transpose: out[c][r] = vals[r][c]
  const transposed: (string | number | boolean | null)[][] = Array.from({ length: cols }, (_, c) =>
    Array.from({ length: rows }, (_, r) => vals[r]?.[c] ?? null),
  );

  const outRng = resolveRange(context, outputRange);
  outRng.getResizedRange(transposed.length - 1, (transposed[0]?.length ?? 1) - 1).values = transposed as any;
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Transposed ${rows}×${cols} → ${cols}×${rows} written to ${outputRange}`,
    outputs: { outputRange },
  };
}

registry.register(meta, handler as any);
export { meta };
