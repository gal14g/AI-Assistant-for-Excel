/**
 * createTable – Convert a range into a structured Excel Table.
 *
 * Office.js notes:
 * - Tables provide auto-filter, structured references, and styling.
 * - Table names must be unique in the workbook.
 * - hasHeaders indicates whether the first row is a header row.
 * - Table styles are predefined strings like "TableStyleMedium2".
 */

import { CapabilityMeta, CreateTableParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange, resolveSheet } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "createTable",
  description: "Convert a range into a structured Excel Table",
  mutates: true,
  affectsFormatting: true,
  requiresApiSet: "ExcelApi 1.1",
};

async function handler(
  context: Excel.RequestContext,
  params: CreateTableParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const { range: address, hasHeaders = true, style } = params;
  // Auto-generate a unique table name if not provided
  const tableName = params.tableName || `Table_${Date.now()}`;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would create table "${tableName}" from ${address}`,
    };
  }

  options.onProgress?.(`Creating table "${tableName}"...`);

  // Resolve sheet from the range address so this works on any sheet,
  // not just the currently active one.
  const range = resolveRange(context, address);
  const sheet = resolveSheet(context, address);
  const table = sheet.tables.add(range, hasHeaders);

  table.name = tableName;

  if (style) {
    table.style = style;
  }

  table.load(["name", "style"]);
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Created table "${table.name}" with style ${table.style}`,
    outputs: { tableName: table.name, tableRange: address },
  };
}


registry.register(meta, handler as any);
export { meta };
