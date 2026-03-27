/**
 * setNumberFormat – Apply a number format string to a range.
 *
 * Common format strings:
 *   "#,##0.00"       → 1,234.56  (number with thousands separator)
 *   "0%"             → 45%       (percentage)
 *   "0.00%"          → 45.50%
 *   "dd/mm/yyyy"     → 01/01/2024
 *   "yyyy-mm-dd"     → 2024-01-01
 *   "$#,##0.00"      → $1,234.56 (USD currency)
 *   "€#,##0.00"      → €1,234.56 (Euro)
 *   "0.00E+00"       → 1.23E+03  (scientific notation)
 *   "@"              → text (force Excel to treat cell as text)
 *   "General"        → Excel default (removes custom format)
 *
 * Office.js notes:
 * - range.numberFormat requires a 2D array exactly matching range dimensions.
 * - We load rowCount/columnCount first, then broadcast the single format string.
 */

import { CapabilityMeta, SetNumberFormatParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "setNumberFormat",
  description: "Apply a number format (currency, percentage, date, etc.) to a range",
  mutates: false,
  affectsFormatting: true,
};

async function handler(
  context: Excel.RequestContext,
  params: SetNumberFormatParams,
  options: ExecutionOptions
): Promise<StepResult> {
  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would apply format "${params.format}" to ${params.range}`,
    };
  }

  options.onProgress?.(`Applying number format "${params.format}"...`);

  const range = resolveRange(context, params.range);
  range.load(["rowCount", "columnCount"]);
  await context.sync();

  // numberFormat must be a 2D array matching the range dimensions
  const formatGrid = Array.from({ length: range.rowCount }, () =>
    Array(range.columnCount).fill(params.format)
  );
  range.numberFormat = formatGrid;
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Applied format "${params.format}" to ${params.range}`,
  };
}

registry.register(meta, handler as any);
export { meta };
