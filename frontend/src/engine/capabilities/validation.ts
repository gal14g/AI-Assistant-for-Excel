/**
 * addValidation – Apply data validation to a range.
 *
 * Office.js notes:
 * - DataValidation API is in ExcelApi 1.8+.
 * - Supports list, wholeNumber, decimal, date, textLength, custom.
 * - Error alerts can be configured with custom messages.
 */

import { CapabilityMeta, AddValidationParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "addValidation",
  description: "Apply data validation rules to a range",
  mutates: false,
  affectsFormatting: false,
  requiresApiSet: "ExcelApi 1.8",
};

async function handler(
  context: Excel.RequestContext,
  params: AddValidationParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const { range: address, validationType, listValues, operator, min, max, formula, showErrorAlert = true, errorMessage } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would add ${validationType} validation to ${address}`,
    };
  }

  options.onProgress?.(`Adding ${validationType} validation...`);

  const range = resolveRange(context, address);

  const validationRule: Excel.DataValidationRule = {};

  switch (validationType) {
    case "list":
      validationRule.list = {
        inCellDropDown: true,
        source: listValues?.join(",") ?? "",
      };
      break;

    case "wholeNumber":
      validationRule.wholeNumber = {
        formula1: min !== undefined ? String(min) : "0",
        formula2: max !== undefined ? String(max) : undefined,
        operator: mapValidationOperator(operator),
      };
      break;

    case "decimal":
      validationRule.decimal = {
        formula1: min !== undefined ? String(min) : "0",
        formula2: max !== undefined ? String(max) : undefined,
        operator: mapValidationOperator(operator),
      };
      break;

    case "date": {
      const dateRule: Partial<NonNullable<Excel.DataValidationRule["date"]>> = {
        operator: mapValidationOperator(operator),
      };
    
      if (min !== undefined) dateRule.formula1 = String(min);
      if (max !== undefined) dateRule.formula2 = String(max);
    
      validationRule.date = dateRule as NonNullable<Excel.DataValidationRule["date"]>;
      break;
    }

    case "textLength":
      validationRule.textLength = {
        formula1: min !== undefined ? String(min) : "0",
        formula2: max !== undefined ? String(max) : undefined,
        operator: mapValidationOperator(operator),
      };
      break;

    case "custom":
      validationRule.custom = {
        formula: formula ?? "=TRUE",
      };
      break;
  }

  range.dataValidation.rule = validationRule;

  if (showErrorAlert) {
    range.dataValidation.errorAlert = {
      showAlert: true,
      style: Excel.DataValidationAlertStyle.stop,
      title: "Invalid Input",
      message: errorMessage ?? "The value you entered is not valid.",
    };
  }

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Applied ${validationType} validation to ${address}`,
  };
}

function mapValidationOperator(operator?: string): Excel.DataValidationOperator {
  switch (operator) {
    case "between": return Excel.DataValidationOperator.between;
    case "greaterThan": return Excel.DataValidationOperator.greaterThan;
    case "lessThan": return Excel.DataValidationOperator.lessThan;
    default: return Excel.DataValidationOperator.between;
  }
}


registry.register(meta, handler as any);
export { meta };
