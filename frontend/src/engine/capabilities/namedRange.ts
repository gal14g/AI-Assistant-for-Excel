/**
 * namedRange – Create, update, or delete named ranges.
 *
 * Named ranges can be scoped to the whole workbook (default) or a specific sheet.
 * They are used in formulas (=SUM(SalesData)), validation lists, and dashboards.
 */

import { CapabilityMeta, NamedRangeParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "namedRange",
  description: "Create, update, or delete a named range in the workbook",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: NamedRangeParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { operation, name, range, sheetName } = params;

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would ${operation} named range "${name}"${range ? ` → ${range}` : ""}` };
  }

  if (operation === "delete") {
    options.onProgress?.(`Deleting named range "${name}"...`);
    try {
      const namedItem = context.workbook.names.getItem(name);
      namedItem.delete();
      await context.sync();
    } catch {
      return { stepId: "", status: "success", message: `Named range "${name}" not found — nothing deleted.` };
    }
    return { stepId: "", status: "success", message: `Deleted named range "${name}"` };
  }

  if (!range) {
    return { stepId: "", status: "error", message: "range param is required for create/update." };
  }

  options.onProgress?.(`${operation === "create" ? "Creating" : "Updating"} named range "${name}" → ${range}...`);

  const rng = resolveRange(context, range);
  rng.load("address");
  await context.sync();

  // Build a proper absolute address for the named range formula
  const absAddr = rng.address; // e.g. "'Sheet1'!$A$1:$B$10"

  if (operation === "update") {
    try {
      const existing = context.workbook.names.getItem(name);
      existing.delete();
      await context.sync();
    } catch { /* may not exist yet */ }
  }

  if (sheetName) {
    const ws = context.workbook.worksheets.getItem(sheetName);
    ws.names.add(name, rng);
  } else {
    context.workbook.names.add(name, rng);
  }
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `${operation === "create" ? "Created" : "Updated"} named range "${name}" → ${absAddr}`,
    outputs: { name, range: absAddr },
  };
}

registry.register(meta, handler as any);
export { meta };
