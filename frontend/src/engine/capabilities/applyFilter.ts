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
import { resolveRange, resolveSheet } from "./rangeUtils";

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

  // Probe whether tableNameOrRange is a named Table or a range address.
  // getItemOrNullObject avoids the crash that getItem() would cause.
  const tableObj = context.workbook.tables.getItemOrNullObject(tableNameOrRange);
  tableObj.load("isNullObject");
  await context.sync();

  if (!tableObj.isNullObject) {
    // ── Table path ─────────────────────────────────────────────────────────
    const column = tableObj.columns.getItemAt(columnIndex);
    column.load("name");
    await context.sync();

    if (criteria.filterOn === "values" && criteria.values) {
      column.filter.applyValuesFilter(criteria.values);
    } else if (criteria.filterOn === "custom" && criteria.operator && criteria.value !== undefined) {
      const criterion1 = criteria.operator === "contains"
        ? `*${criteria.value}*`
        : `${mapOperator(criteria.operator)}${criteria.value}`;
      const filterCriteria: Excel.FilterCriteria = {
        filterOn: Excel.FilterOn.custom,
        criterion1,
      };
      column.filter.apply(filterCriteria);
    } else if (criteria.filterOn === "topItems" && criteria.value) {
      column.filter.applyTopItemsFilter(Number(criteria.value));
    }

    await context.sync();

    return {
      stepId: "",
      status: "success",
      message: `Applied ${criteria.filterOn} filter on column ${columnIndex} of table "${tableNameOrRange}"`,
      outputs: { range: tableNameOrRange },
    };
  }

  // ── AutoFilter path (range address) ───────────────────────────────────
  // Use resolveRange/resolveSheet so the correct sheet is used regardless
  // of which sheet is currently active.
  const range = resolveRange(context, tableNameOrRange);
  const sheet = resolveSheet(context, tableNameOrRange);

  if (criteria.filterOn === "values" && criteria.values) {
    sheet.autoFilter.apply(range, columnIndex, {
      filterOn: Excel.FilterOn.values,
      values: criteria.values,
    });
  } else if (criteria.filterOn === "custom" && criteria.operator && criteria.value !== undefined) {
    const criterion1 = criteria.operator === "contains"
      ? `*${criteria.value}*`
      : `${mapOperator(criteria.operator)}${criteria.value}`;
    sheet.autoFilter.apply(range, columnIndex, {
      filterOn: Excel.FilterOn.custom,
      criterion1,
    });
  } else if (criteria.filterOn === "topItems" && criteria.value) {
    sheet.autoFilter.apply(range, columnIndex, {
      filterOn: Excel.FilterOn.topItems,
      criterion1: String(criteria.value),
    });
  }

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Applied autoFilter on ${tableNameOrRange} column ${columnIndex}`,
    outputs: { range: tableNameOrRange },
  };
}

function mapOperator(op: string): string {
  switch (op) {
    case "greaterThan": return ">";
    case "lessThan":    return "<";
    case "equals":      return "=";
    case "contains":    return "*";
    default:            return "=";
  }
}

registry.register(meta, handler as any);
export { meta };
