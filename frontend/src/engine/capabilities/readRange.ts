/**
 * readRange – Read cell values from a range.
 *
 * Office.js notes:
 * - range.load("values") reads only calculated values (not formulas).
 * - For ranges with headers, the first row of values[] is the header row.
 * - Large ranges (>10k cells) may need chunking in production.
 */

import { CapabilityMeta, ReadRangeParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "readRange",
  description: "Read values from a cell range",
  mutates: false,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: ReadRangeParams,
  options: ExecutionOptions
): Promise<StepResult> {
  options.onProgress?.("Reading range...");

  const range = resolveRange(context, params.range);
  range.load(["values", "address", "rowCount", "columnCount"]);
  await context.sync();

  const values = range.values;
  const rowCount = range.rowCount;
  const colCount = range.columnCount;

  options.onProgress?.(`Read ${rowCount} rows x ${colCount} columns`);

  return {
    stepId: "",
    status: "success",
    message: `Read ${rowCount}x${colCount} from ${range.address}`,
    data: {
      values,
      address: range.address,
      rowCount,
      columnCount: colCount,
      headers: params.includeHeaders ? (values ?? [])[0] : undefined,
    },
  };
}


registry.register(meta, handler as any);
export { meta };
