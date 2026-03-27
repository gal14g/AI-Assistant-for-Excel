/**
 * writeValues – Write a 2D array of values to a range.
 *
 * CRITICAL: By default, this writes ONLY values. It does NOT copy or alter
 * any formatting (fill, font, borders, number formats). This is the primary
 * mechanism for preserving existing workbook formatting.
 *
 * Office.js notes:
 * - Setting range.values writes raw values; formulas in the values array
 *   will be written as literal strings unless they start with "=".
 * - The values array dimensions must match the target range dimensions.
 * - Null values in the array preserve the existing cell value.
 */

import { CapabilityMeta, WriteValuesParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "writeValues",
  description: "Write values to a cell range (preserves formatting)",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: WriteValuesParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const { range: address, values, valuesOnly = true } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would write ${values.length} rows to ${address}`,
    };
  }

  options.onProgress?.(`Writing ${values.length} rows to ${address}...`);

  const range = resolveRange(context, address);

  // Only set values — never touch formatting
  range.values = values;

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Wrote ${values.length} rows to ${address}`,
  };
}


registry.register(meta, handler as any);
export { meta };
