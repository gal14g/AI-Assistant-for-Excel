/**
 * applyFilter – Apply a filter to a table or auto-filter range.
 *
 * Office.js notes:
 * - Filters work on Table objects or via Worksheet.autoFilter.
 * - Column index is 0-based within the table.
 * - Multiple filter types: values, topItems, custom.
 */

import { CapabilityMeta, ApplyFilterParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";

const meta: CapabilityMeta = {
  action: "applyFilter",
  description: "Apply filters to a table or range",
  mutates: false,
  affectsFormatting: false,
  requiresApiSet: "ExcelApi 1.2",
};

async function handler(
  context: Excel.RequestContext,
  params: ApplyFilterParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const { tableNameOrRange, columnIndex, criteria } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would apply filter on column ${columnIndex} of ${tableNameOrRange}`,
    };
  }

  options.onProgress?.("Applying filter...");

  // Try as table name first, fall back to range-based autoFilter
  try {
    const table = context.workbook.tables.getItem(tableNameOrRange);
    const column = table.columns.getItemAt(columnIndex);
    column.load("name");
    await context.sync();

    if (criteria.filterOn === "values" && criteria.values) {
      column.filter.applyValuesFilter(criteria.values);
    } else if (criteria.filterOn === "custom" && criteria.operator && criteria.value !== undefined) {
      const filterCriteria: Excel.FilterCriteria = {
        filterOn: Excel.FilterOn.custom,
        criterion1: `${mapOperator(criteria.operator)}${criteria.value}`,
      };
      column.filter.apply(filterCriteria);
    } else if (criteria.filterOn === "topItems" && criteria.value) {
      column.filter.applyTopItemsFilter(Number(criteria.value));
    }

    await context.sync();

    return {
      stepId: "",
      status: "success",
      message: `Applied ${criteria.filterOn} filter on column ${columnIndex}`,
    };
  } catch {
    // Fall back to autoFilter on range
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(tableNameOrRange);

    if (criteria.filterOn === "values" && criteria.values) {
      sheet.autoFilter.apply(range, columnIndex, {
        filterOn: Excel.FilterOn.values,
        values: criteria.values,
      });
    }

    await context.sync();

    return {
      stepId: "",
      status: "success",
      message: `Applied autoFilter on ${tableNameOrRange} column ${columnIndex}`,
    };
  }
}

function mapOperator(op: string): string {
  switch (op) {
    case "greaterThan": return ">";
    case "lessThan": return "<";
    case "equals": return "=";
    case "contains": return "*";
    default: return "=";
  }
}

registry.register(meta, handler as any);
export { meta };
