/**
 * sortRange – Sort a range by one or more columns.
 *
 * Office.js notes:
 * - RangeSort.apply() takes an array of SortField objects.
 * - Sort modifies the range in-place (row order changes).
 * - hasHeaders tells Excel to skip the first row during sort.
 */

import { CapabilityMeta, SortRangeParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";
import { ensureUnmerged } from "../utils/mergedCells";

const meta: CapabilityMeta = {
  action: "sortRange",
  description: "Sort a range by one or more columns",
  mutates: true,
  affectsFormatting: false,
  requiresApiSet: "ExcelApi 1.2",
};

async function handler(
  context: Excel.RequestContext,
  params: SortRangeParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const { range: address, hasHeaders = true } = params;
  // Default: sort by first column ascending if no sortFields provided
  const sortFields = params.sortFields?.length ? params.sortFields : [{ columnIndex: 0, ascending: true }];

  if (options.dryRun) {
    const desc = sortFields
      .map((f) => `col ${f.columnIndex} ${f.ascending !== false ? "asc" : "desc"}`)
      .join(", ");
    return { stepId: "", status: "success", message: `Would sort ${address} by ${desc}` };
  }

  options.onProgress?.("Sorting range...");

  const range = resolveRange(context, address);

  // Office.js Range.sort.apply() errors outright if the range contains
  // merged cells. Auto-unmerge so the sort can proceed.
  const mergeReport = await ensureUnmerged(context, range, {
    operation: "sortRange",
    policy: "refuseWithError",
  });
  if (mergeReport.error) return mergeReport.error;

  const excelSortFields: Excel.SortField[] = sortFields.map((f) => ({
    key: f.columnIndex,
    ascending: f.ascending !== false,
  }));

  range.sort.apply(excelSortFields, false, hasHeaders);
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Sorted ${address} by ${sortFields.length} field(s)${mergeReport.warning ?? ""}`,
    outputs: { range: address },
  };
}


registry.register(meta, handler as any);
export { meta };
