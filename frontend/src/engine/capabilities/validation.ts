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
      // Support both a comma-separated list (listValues) AND a range reference
      // (formula = "=Sheet2!A:A"). Range reference is preferred for dynamic dropdowns.
      validationRule.list = {
        inCellDropDown: true,
        source: formula ?? listValues?.join(",") ?? "",
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
    outputs: { range: address },
  };
}

function mapValidationOperator(operator?: string): Excel.DataValidationOperator {
  switch (operator) {
    case "between":             return Excel.DataValidationOperator.between;
    case "notBetween":          return Excel.DataValidationOperator.notBetween;
    case "equalTo":             return Excel.DataValidationOperator.equalTo;
    case "notEqualTo":          return Excel.DataValidationOperator.notEqualTo;
    case "greaterThan":         return Excel.DataValidationOperator.greaterThan;
    case "greaterThanOrEqualTo":return Excel.DataValidationOperator.greaterThanOrEqualTo;
    case "lessThan":            return Excel.DataValidationOperator.lessThan;
    case "lessThanOrEqualTo":   return Excel.DataValidationOperator.lessThanOrEqualTo;
    default:                    return Excel.DataValidationOperator.between;
  }
}


// ── Legacy-Excel fallback (ExcelApi < 1.8) ────────────────────────────────────
// Range.dataValidation requires 1.8. There is no Office.js primitive on 1.3
// that enforces cell-level input constraints. We cannot emulate "reject value
// N > 100" without the validation engine. Fallback strategy: annotate each
// cell in the target range with a grey italic caption describing the
// constraint, so the user/data-entry agent *sees* the rule even though it
// isn't enforced. This intentionally doesn't touch cell values.
async function fallback(
  context: Excel.RequestContext,
  params: AddValidationParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { range: address, validationType, listValues, min, max, formula } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would annotate ${address} with ${validationType} validation guidance (legacy fallback).`,
    };
  }

  options.onProgress?.(`Legacy-Excel mode: data validation unavailable, annotating ${address}...`);

  const range = resolveRange(context, address);
  range.load(["rowCount", "columnCount", "address"]);
  await context.sync();

  // Build a human-readable rule description.
  let rule = "";
  switch (validationType) {
    case "list":
      rule = formula
        ? `Allowed values from range ${formula}`
        : `Allowed: ${(listValues ?? []).join(", ")}`;
      break;
    case "wholeNumber":
      rule = `Whole number${min !== undefined ? ` ≥ ${min}` : ""}${max !== undefined ? ` and ≤ ${max}` : ""}`;
      break;
    case "decimal":
      rule = `Number${min !== undefined ? ` ≥ ${min}` : ""}${max !== undefined ? ` and ≤ ${max}` : ""}`;
      break;
    case "date":
      rule = `Date${min !== undefined ? ` on/after ${min}` : ""}${max !== undefined ? ` and on/before ${max}` : ""}`;
      break;
    case "textLength":
      rule = `Text length${min !== undefined ? ` ≥ ${min}` : ""}${max !== undefined ? ` and ≤ ${max}` : ""}`;
      break;
    case "custom":
      rule = `Must satisfy formula ${formula ?? "(none)"}`;
      break;
  }

  // Apply a light, borderless visual cue to the target range so the rule
  // is discoverable. This is purely cosmetic — no actual enforcement.
  try {
    range.format.fill.color = "#FFF8E1"; // pale gold — "soft warn"
  } catch { /* best-effort */ }

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message:
      `Data validation for ${address} (${validationType}) could not be installed — requires ExcelApi 1.8+. ` +
      `Cells tinted as a visual cue; rule is NOT enforced by Excel. ${rule} (legacy-Excel fallback).`,
    outputs: { range: address },
  };
}

registry.register(meta, handler as any, fallback as any);
export { meta };
