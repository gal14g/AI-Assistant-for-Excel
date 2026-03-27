/**
 * createPivot – Create a PivotTable from source data.
 *
 * Office.js notes:
 * - PivotTable API requires ExcelApi 1.8+.
 * - Fields are added by name (must match header names in source data).
 * - We read the source headers first so we can validate and auto-detect fields.
 * - The pivot name must be non-empty and unique within the workbook.
 */

import { CapabilityMeta, CreatePivotParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "createPivot",
  description: "Create a PivotTable from source data",
  mutates: true,
  affectsFormatting: true,
  requiresApiSet: "ExcelApi 1.8",
};

async function handler(
  context: Excel.RequestContext,
  params: CreatePivotParams,
  options: ExecutionOptions
): Promise<StepResult> {
  let { sourceRange, destinationRange, pivotName, rows, columns, values, filters } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would create PivotTable from ${sourceRange}`,
    };
  }

  // ── 1. Read source headers so we can validate / auto-detect fields ──────
  const sourceRng = resolveRange(context, sourceRange);
  sourceRng.load("values");
  await context.sync();

  const firstRow = (sourceRng.values ?? [])[0] as (string | number | boolean)[] | undefined;
  const headers = (firstRow ?? []).map((h) => String(h)).filter(Boolean);

  if (headers.length === 0) {
    return { stepId: "", status: "error", message: "Source range has no headers." };
  }

  // ── 2. Defaults ─────────────────────────────────────────────────────────
  if (!pivotName) pivotName = `PivotTable_${Date.now()}`;

  // If the LLM guessed field names that don't exist, fall back to actual headers.
  const validField = (name: string) => headers.includes(name);

  if (!rows || rows.length === 0 || !rows.every(validField)) {
    rows = [headers[0]]; // first column as row grouping
  }
  if (!values || values.length === 0 || !values.every((v) => validField(v.field))) {
    // Use the first header not already used as a row
    const valueHeader = headers.find((h) => !rows.includes(h)) ?? headers[0];
    values = [{ field: valueHeader, summarizeBy: "sum" }];
  }
  if (columns && !columns.every(validField)) columns = undefined;
  if (filters && !filters.every(validField)) filters = undefined;

  // ── 3. Destination ───────────────────────────────────────────────────────
  // If destinationRange is missing, add a new sheet for the pivot.
  let destRng: Excel.Range;
  if (!destinationRange) {
    const pivotSheetName = pivotName.slice(0, 31); // Excel sheet name limit
    const newSheet = context.workbook.worksheets.add(pivotSheetName);
    destRng = newSheet.getRange("A1");
  } else {
    destRng = resolveRange(context, destinationRange);
  }

  options.onProgress?.(`Creating PivotTable "${pivotName}"...`);

  // ── 4. Create the PivotTable ─────────────────────────────────────────────
  const pivotTable = context.workbook.pivotTables.add(pivotName, sourceRng, destRng);
  await context.sync();

  // Row fields
  for (const fieldName of rows) {
    options.onProgress?.(`Adding row field: ${fieldName}`);
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(fieldName));
  }

  // Column fields
  if (columns) {
    for (const fieldName of columns) {
      options.onProgress?.(`Adding column field: ${fieldName}`);
      pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem(fieldName));
    }
  }

  // Value fields
  for (const val of values) {
    options.onProgress?.(`Adding value field: ${val.field} (${val.summarizeBy})`);
    const dataHierarchy = pivotTable.dataHierarchies.add(
      pivotTable.hierarchies.getItem(val.field)
    );
    dataHierarchy.summarizeBy = mapSummarizeBy(val.summarizeBy);
    if (val.displayName) {
      dataHierarchy.name = val.displayName;
    }
  }

  // Filter fields
  if (filters) {
    for (const fieldName of filters) {
      pivotTable.filterHierarchies.add(pivotTable.hierarchies.getItem(fieldName));
    }
  }

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Created PivotTable "${pivotName}" — rows: ${rows.join(", ")} | values: ${values.map((v) => v.field).join(", ")}`,
  };
}

function mapSummarizeBy(summarizeBy: string): Excel.AggregationFunction {
  switch (summarizeBy) {
    case "sum":     return Excel.AggregationFunction.sum;
    case "count":   return Excel.AggregationFunction.count;
    case "average": return Excel.AggregationFunction.average;
    case "max":     return Excel.AggregationFunction.max;
    case "min":     return Excel.AggregationFunction.min;
    default:        return Excel.AggregationFunction.sum;
  }
}

registry.register(meta, handler as any);
export { meta };
