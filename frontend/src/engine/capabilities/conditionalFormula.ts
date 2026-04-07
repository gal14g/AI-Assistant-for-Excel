/**
 * conditionalFormula – Write IF-based formulas that apply different formulas
 * based on a condition.
 *
 * Builds per-row IF() formulas using a condition column and writes them to
 * an output range. Supports conditions: blank, notBlank, equals, notEquals,
 * contains, greaterThan, lessThan.
 *
 * The trueFormula and falseFormula params support a {row} placeholder that
 * is replaced with the actual Excel row number for each row.
 */

import { CapabilityMeta, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange, stripWorkbookQualifier } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "conditionalFormula",
  description: "Write IF-based conditional formulas to a range",
  mutates: true,
  affectsFormatting: false,
};

/**
 * Convert a 1-based column number to a column letter (1 → A, 2 → B, 27 → AA).
 */
function columnNumberToLetter(col: number): string {
  let letter = "";
  while (col > 0) {
    const mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

async function handler(
  context: Excel.RequestContext,
  params: any,
  options: ExecutionOptions
): Promise<StepResult> {
  const {
    range: address,
    conditionColumn,
    condition,
    conditionValue,
    trueFormula,
    falseFormula,
    outputRange: outputAddress,
    hasHeaders = true,
  } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would create conditional formulas based on "${condition}" in ${address}`,
    };
  }

  options.onProgress?.("Building conditional formulas...");

  // Get the source range to determine row count
  const sourceRange = resolveRange(context, address);
  sourceRange.load(["rowCount", "rowIndex"]);
  await context.sync();

  const totalRows = sourceRange.rowCount;
  const startExcelRow = sourceRange.rowIndex + 1; // 1-based Excel row
  const dataStartRow = hasHeaders ? startExcelRow + 1 : startExcelRow;
  const dataRowCount = hasHeaders ? totalRows - 1 : totalRows;

  if (dataRowCount <= 0) {
    return { stepId: "", status: "success", message: "No data rows to process." };
  }

  // Build the column letter for the condition column
  const colLetter = columnNumberToLetter(conditionColumn);

  // Build formulas for each data row
  const formulas: string[][] = [];
  for (let i = 0; i < dataRowCount; i++) {
    const rowNum = dataStartRow + i;
    const cellRef = `${colLetter}${rowNum}`;

    // Replace {row} placeholder in true/false formulas
    const trueExpr = (trueFormula as string).replace(/\{row\}/g, String(rowNum));
    const falseExpr = (falseFormula as string).replace(/\{row\}/g, String(rowNum));

    let conditionExpr: string;
    switch (condition) {
      case "blank":
        conditionExpr = `${cellRef}=""`;
        break;
      case "notBlank":
        conditionExpr = `${cellRef}<>""`;
        break;
      case "equals":
        conditionExpr = typeof conditionValue === "number"
          ? `${cellRef}=${conditionValue}`
          : `${cellRef}="${conditionValue}"`;
        break;
      case "notEquals":
        conditionExpr = typeof conditionValue === "number"
          ? `${cellRef}<>${conditionValue}`
          : `${cellRef}<>"${conditionValue}"`;
        break;
      case "contains":
        conditionExpr = `ISNUMBER(SEARCH("${conditionValue}",${cellRef}))`;
        break;
      case "greaterThan":
        conditionExpr = `${cellRef}>${conditionValue}`;
        break;
      case "lessThan":
        conditionExpr = `${cellRef}<${conditionValue}`;
        break;
      default:
        conditionExpr = `${cellRef}="${conditionValue}"`;
    }

    formulas.push([`=IF(${conditionExpr},${trueExpr},${falseExpr})`]);
  }

  // Write to the output range
  const cleanOutput = stripWorkbookQualifier(outputAddress);
  const outRange = resolveRange(context, cleanOutput);

  // Write header if applicable
  if (hasHeaders) {
    const headerRange = outRange.getRow(0);
    headerRange.values = [["Result"]];
    const dataRange = outRange.getOffsetRange(1, 0).getResizedRange(formulas.length - 1, 0);
    dataRange.formulas = formulas as any;
  } else {
    outRange.getResizedRange(formulas.length - 1, 0).formulas = formulas as any;
  }

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Created ${formulas.length} conditional formulas`,
  };
}

registry.register(meta, handler as any);
export { meta };
