/**
 * Sheet operations – add, rename, delete, copy, protect sheets.
 *
 * These operations map to Worksheet collection methods in Office.js.
 * Each is a distinct sub-action within the "sheetOps" capability.
 */

import { CapabilityMeta, SheetOpParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";

// Register each sheet operation as a separate capability
const sheetActions: {
  action: "addSheet" | "renameSheet" | "deleteSheet" | "copySheet" | "protectSheet";
  desc: string;
  mutates: boolean;
}[] = [
  { action: "addSheet", desc: "Add a new worksheet", mutates: true },
  { action: "renameSheet", desc: "Rename a worksheet", mutates: true },
  { action: "deleteSheet", desc: "Delete a worksheet", mutates: true },
  { action: "copySheet", desc: "Copy a worksheet", mutates: true },
  { action: "protectSheet", desc: "Protect a worksheet", mutates: false },
];

for (const sa of sheetActions) {
  const meta: CapabilityMeta = {
    action: sa.action,
    description: sa.desc,
    mutates: sa.mutates,
    affectsFormatting: false,
  };

  const handler = async (
    context: Excel.RequestContext,
    params: SheetOpParams,
    options: ExecutionOptions
  ): Promise<StepResult> => {
    if (options.dryRun) {
      return {
        stepId: "",
        status: "success",
        message: `Would ${sa.action} "${params.sheetName}"`,
      };
    }

    switch (sa.action) {
      case "addSheet": {
        options.onProgress?.(`Adding sheet "${params.sheetName}"...`);
        // Use getItemOrNullObject so we don't crash if the sheet already exists
        const existing = context.workbook.worksheets.getItemOrNullObject(params.sheetName);
        existing.load("isNullObject");
        await context.sync();
        if (!existing.isNullObject) {
          return {
            stepId: "",
            status: "success",
            message: `Sheet "${params.sheetName}" already exists — using it`,
            outputs: { sheetName: params.sheetName },
          };
        }
        const newSheet = context.workbook.worksheets.add(params.sheetName);
        newSheet.load("name");
        await context.sync();
        return {
          stepId: "",
          status: "success",
          message: `Added sheet "${newSheet.name}"`,
          outputs: { sheetName: newSheet.name },
        };
      }

      case "renameSheet": {
        options.onProgress?.(`Renaming "${params.sheetName}" to "${params.newName}"...`);
        const sheetToRename = context.workbook.worksheets.getItemOrNullObject(params.sheetName);
        sheetToRename.load("isNullObject");
        await context.sync();
        if (sheetToRename.isNullObject) {
          return { stepId: "", status: "error", message: `Sheet "${params.sheetName}" not found` };
        }
        sheetToRename.name = params.newName ?? params.sheetName;
        await context.sync();
        return {
          stepId: "",
          status: "success",
          message: `Renamed sheet to "${params.newName}"`,
        };
      }

      case "deleteSheet": {
        options.onProgress?.(`Deleting sheet "${params.sheetName}"...`);
        const sheetToDelete = context.workbook.worksheets.getItemOrNullObject(params.sheetName);
        sheetToDelete.load("isNullObject");
        await context.sync();
        if (sheetToDelete.isNullObject) {
          // Already gone — treat as success (idempotent)
          return { stepId: "", status: "success", message: `Sheet "${params.sheetName}" already deleted` };
        }
        sheetToDelete.delete();
        await context.sync();
        return {
          stepId: "",
          status: "success",
          message: `Deleted sheet "${params.sheetName}"`,
        };
      }

      case "copySheet": {
        options.onProgress?.(`Copying sheet "${params.sheetName}"...`);
        const source = context.workbook.worksheets.getItem(params.sheetName);
        const copy = source.copy("End");
        copy.load("name");
        await context.sync();
        if (params.newName) {
          copy.name = params.newName;
          await context.sync();
        }
        return {
          stepId: "",
          status: "success",
          message: `Copied sheet "${params.sheetName}" as "${copy.name}"`,
        };
      }

      case "protectSheet": {
        options.onProgress?.(`Protecting sheet "${params.sheetName}"...`);
        const sheet = context.workbook.worksheets.getItem(params.sheetName);
        sheet.protection.protect({
          allowAutoFilter: true,
          allowSort: true,
        }, params.password);
        await context.sync();
        return {
          stepId: "",
          status: "success",
          message: `Protected sheet "${params.sheetName}"`,
        };
      }

      default:
        return {
          stepId: "",
          status: "error",
          message: `Unknown sheet operation: ${sa.action}`,
        };
    }
  };

  registry.register(meta, handler as any);
}
