/**
 * addDropdownControl – Create a dropdown (data validation list) in a cell.
 *
 * The dropdown can be driven by a comma-separated list of values or by a
 * range reference (e.g. "Sheet2!A:A"). Useful as a filter control or
 * user-input mechanism.
 *
 * Office.js notes:
 * - DataValidation API is in ExcelApi 1.8+.
 * - For list validation, formula1 is either a comma-separated string or a
 *   range reference prefixed with "=".
 */

import { CapabilityMeta, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange, stripWorkbookQualifier } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "addDropdownControl",
  description: "Create a dropdown (data validation list) in a cell",
  mutates: true,
  affectsFormatting: false,
  requiresApiSet: "ExcelApi 1.8",
};

/**
 * Determine whether a listSource string looks like a range reference
 * rather than a comma-separated list of literal values.
 */
function isRangeReference(source: string): boolean {
  const trimmed = source.trim();
  // Contains "!" (sheet-qualified) or ":" (multi-cell range) or is a simple column/cell ref
  if (trimmed.includes("!") || trimmed.includes(":")) return true;
  // Matches a simple cell or column reference like "A1", "B", "AA10"
  if (/^[A-Z]{1,3}\d*$/i.test(trimmed)) return true;
  return false;
}

async function handler(
  context: Excel.RequestContext,
  params: any,
  options: ExecutionOptions
): Promise<StepResult> {
  const {
    cell,
    listSource,
    promptMessage,
    sheetName,
  } = params;

  if (!cell || !listSource) {
    return {
      stepId: "",
      status: "error",
      message: "cell and listSource are required",
    };
  }

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would add dropdown control to ${cell}`,
    };
  }

  options.onProgress?.(`Adding dropdown control to ${cell}...`);

  // Resolve target cell — prepend sheetName if provided
  const cellAddress = sheetName && !cell.includes("!") ? `${sheetName}!${cell}` : cell;
  const range = resolveRange(context, cellAddress);

  // Build the validation source
  let source: string;
  let optionCount: number | string;

  if (isRangeReference(listSource)) {
    // Range reference — strip workbook qualifier and prefix with "="
    source = "=" + stripWorkbookQualifier(listSource);
    optionCount = "range " + listSource;
  } else {
    // Comma-separated literal values
    source = listSource;
    optionCount = listSource.split(",").length + " options";
  }

  const validationRule: Excel.DataValidationRule = {
    list: {
      inCellDropDown: true,
      source,
    },
  };

  range.dataValidation.rule = validationRule;

  if (promptMessage) {
    range.dataValidation.prompt = {
      showPrompt: true,
      title: "Select a value",
      message: promptMessage,
    };
  }

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Added dropdown control to ${cell} with ${optionCount}`,
  };
}


registry.register(meta, handler as any);
export { meta };
