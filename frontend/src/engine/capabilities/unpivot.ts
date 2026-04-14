/**
 * unpivot – Convert a wide table into a tall (tidy) format.
 *
 * Example:
 *   Name | Jan | Feb | Mar          Name | Month | Value
 *   Alice|  10 |  20 |  30    →     Alice| Jan   | 10
 *   Bob  |   5 |  15 |  25          Alice| Feb   | 20
 *                                   ...
 *
 * The first `idColumns` columns are kept as-is on every output row.
 * The remaining columns become rows with two new columns: variable + value.
 */

import { CapabilityMeta, UnpivotParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "unpivot",
  description: "Unpivot (melt) a wide table to tall/tidy format",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: UnpivotParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const {
    sourceRange,
    idColumns,
    outputRange,
    variableColumnName = "Attribute",
    valueColumnName = "Value",
  } = params;

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would unpivot ${sourceRange} (${idColumns} id col(s)) to ${outputRange}` };
  }

  options.onProgress?.("Reading source table...");
  const srcRng = resolveRange(context, sourceRange);
  srcRng.load("values");
  await context.sync();

  const data = (srcRng.values ?? []) as (string | number | boolean | null)[][];
  if (data.length < 2) return { stepId: "", status: "success", message: "Not enough rows to unpivot." };

  const headers = data[0];
  const idHeaders = headers.slice(0, idColumns);
  const valueHeaders = headers.slice(idColumns);

  options.onProgress?.(`Unpivoting ${data.length - 1} rows × ${valueHeaders.length} value columns...`);

  // Build output rows: [id1, id2, ..., variableCol, valueCol]
  const outHeaders = [...idHeaders.map(String), variableColumnName, valueColumnName];
  const outRows: (string | number | boolean | null)[][] = [outHeaders];

  for (let r = 1; r < data.length; r++) {
    const idVals = data[r].slice(0, idColumns);
    for (let c = 0; c < valueHeaders.length; c++) {
      outRows.push([...idVals, String(valueHeaders[c]), data[r][idColumns + c]]);
    }
  }

  const outRng = resolveRange(context, outputRange);
  outRng.getResizedRange(outRows.length - 1, outHeaders.length - 1).values = outRows as any;
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Unpivoted ${data.length - 1} rows into ${outRows.length - 1} rows (${valueHeaders.length} value columns → rows)`,
    outputs: { outputRange },
  };
}

registry.register(meta, handler as any);
export { meta };
