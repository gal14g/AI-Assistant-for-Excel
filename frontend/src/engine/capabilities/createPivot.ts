/**
 * createPivot – Create a PivotTable from source data.
 *
 * Office.js notes:
 * - PivotTable API requires ExcelApi 1.8+.
 * - Fields are added by name (must match header names in source data).
 * - We read the source headers first so we can validate and auto-detect fields.
 * - The pivot name must be non-empty and unique within the workbook.
 *
 * LLM resilience:
 * - Small models (llama3, etc.) often output range addresses as field names,
 *   e.g. rows: ["Sheet2!A:A"] instead of rows: ["מספר עובד"].
 * - resolveFieldRef() converts those addresses to the matching header name
 *   by extracting the column letter and mapping it to the header at that index.
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
  const { sourceRange, destinationRange } = params;
  let { pivotName, rows, columns, values, filters } = params;

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

  // ── 2. Resolve field references ─────────────────────────────────────────
  // Small LLMs frequently output range addresses ("Sheet2!A:A") instead of
  // field names ("מספר עובד"). Convert those addresses to actual header names.
  const resolveRef = (ref: string): string => {
    if (!ref || typeof ref !== "string") return headers[0]; // fallback for undefined/null
    if (headers.includes(ref)) return ref; // already a valid header name
    const colIdx = columnLetterToIndex(extractColumnLetter(ref));
    if (colIdx !== null && colIdx < headers.length) return headers[colIdx];
    return ref; // keep as-is; will fail gracefully later
  };

  if (rows && rows.length > 0) {
    rows = rows.map(resolveRef);
  }
  if (values && values.length > 0) {
    values = values.map((v) => ({ ...v, field: resolveRef(v.field) }));
  }
  if (columns && columns.length > 0) {
    columns = columns.map(resolveRef);
  }
  if (filters && filters.length > 0) {
    filters = filters.map(resolveRef);
  }

  // ── 3. Defaults for any still-invalid fields ─────────────────────────────
  if (!pivotName) pivotName = `PivotTable_${Date.now()}`;

  const validField = (name: string) => headers.includes(name);

  if (!rows || rows.length === 0 || !rows.every(validField)) {
    rows = [headers[0]];
  }
  if (!values || values.length === 0 || !values.every((v) => validField(v.field))) {
    // Prefer a numeric column for the value field — dates/text are poor SUM candidates.
    // Check the second row (first data row) to find columns that contain numbers.
    const dataRows = (sourceRng.values ?? []).slice(1); // skip header row
    const candidateHeaders = headers.filter((h) => !(rows as string[]).includes(h));
    let valueHeader = candidateHeaders[0] ?? headers[0]; // default: first non-row header

    if (dataRows.length > 0) {
      const numericHeader = candidateHeaders.find((h) => {
        const colIdx = headers.indexOf(h);
        // Check if the majority of data values in this column are numbers
        const numericCount = dataRows.filter(
          (row) => typeof row[colIdx] === "number"
        ).length;
        return numericCount > dataRows.length / 2;
      });
      if (numericHeader) valueHeader = numericHeader;
    }

    values = [{ field: valueHeader, summarizeBy: "sum" }];
  }
  if (columns && !columns.every(validField)) columns = undefined;
  if (filters && !filters.every(validField)) filters = undefined;

  // ── 4. Destination ───────────────────────────────────────────────────────
  // If destinationRange is missing, place the pivot on a new sheet.
  let destRng: Excel.Range;
  if (!destinationRange) {
    const pivotSheetName = pivotName.slice(0, 31); // Excel sheet name max length
    // Reuse existing sheet if it already exists (idempotent)
    const existing = context.workbook.worksheets.getItemOrNullObject(pivotSheetName);
    existing.load("isNullObject");
    await context.sync();
    const pivotSheet = existing.isNullObject
      ? context.workbook.worksheets.add(pivotSheetName)
      : existing;
    destRng = pivotSheet.getRange("A1");
  } else {
    destRng = resolveRange(context, destinationRange);
  }

  options.onProgress?.(`Creating PivotTable "${pivotName}"...`);

  // ── 5. Create the PivotTable ─────────────────────────────────────────────
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
    outputs: { pivotName },
  };
}

// ── Helpers ──────────────────────────────────────────────────────────────────

/**
 * Extract the first column letter sequence from a range address.
 * "Sheet2!D:D"   → "D"
 * "[WB.xlsx]Sheet2!AB1:AB7" → "AB"
 * "D:D"          → "D"
 * "D1"           → "D"
 */
function extractColumnLetter(address: string | undefined | null): string | null {
  if (!address || typeof address !== "string") return null;
  // Strip workbook qualifier [...]
  const clean = address.replace(/^\[.*?\]/, "");
  // Strip sheet name (everything up to and including "!")
  const afterBang = clean.includes("!") ? clean.split("!").pop()! : clean;
  // Strip single quotes around sheet name if any leaked
  const ref = afterBang.replace(/^'/, "");
  const match = ref.match(/^([A-Za-z]+)/);
  return match ? match[1].toUpperCase() : null;
}

/**
 * Convert a column letter to a 0-based index.
 * "A" → 0, "B" → 1, "Z" → 25, "AA" → 26, "D" → 3
 */
function columnLetterToIndex(letter: string | null): number | null {
  if (!letter) return null;
  let index = 0;
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.charCodeAt(i) - 64);
  }
  return index - 1; // convert to 0-based
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
