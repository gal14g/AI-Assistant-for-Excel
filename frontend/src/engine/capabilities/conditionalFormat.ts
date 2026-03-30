/**
 * addConditionalFormat – Apply conditional formatting rules.
 *
 * Supported rule types:
 *   cellValue  – compare cell value (>, <, between, =, !=, >=, <=, notBetween)
 *   colorScale – green-yellow-red gradient (auto-configured)
 *   dataBar    – in-cell bar proportional to value
 *   iconSet    – icon set (3/4/5 icons, auto-thresholds)
 *   text       – contains text
 *   formula    – custom Excel formula (most powerful: highlight row, blank check, cross-col compare)
 *
 * Office.js notes:
 * - ConditionalFormat API is in ExcelApi 1.6+.
 * - Conditional formats stack; this adds a new rule without removing existing ones.
 */

import {
  CapabilityMeta,
  AddConditionalFormatParams,
  StepResult,
  ExecutionOptions,
} from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "addConditionalFormat",
  description: "Apply conditional formatting rules to a range",
  mutates: true,
  affectsFormatting: true,
  requiresApiSet: "ExcelApi 1.6",
};

async function handler(
  context: Excel.RequestContext,
  params: AddConditionalFormatParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const { range: address, ruleType, operator, values, format, text, formula } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would add ${ruleType} conditional format to ${address}`,
    };
  }

  options.onProgress?.(`Applying ${ruleType} conditional format...`);

  const range = resolveRange(context, address);

  switch (ruleType) {
    case "cellValue": {
      const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
      const rule = cf.cellValue;
      rule.rule = buildCellValueRule(operator, values);
      if (format) {
        if (format.fillColor) rule.format.fill.color = format.fillColor;
        if (format.fontColor) rule.format.font.color = format.fontColor;
        if (format.bold !== undefined) rule.format.font.bold = format.bold;
      }
      break;
    }

    case "formula": {
      // Formula-based rules are the most powerful — they can reference other cells,
      // check entire rows, handle blanks, etc. Formula must start with "=".
      const formulaStr = formula ?? text ?? "=TRUE";
      const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
      cf.custom.rule.formula = formulaStr;
      if (format) {
        if (format.fillColor) cf.custom.format.fill.color = format.fillColor;
        if (format.fontColor) cf.custom.format.font.color = format.fontColor;
        if (format.bold !== undefined) cf.custom.format.font.bold = format.bold;
      }
      break;
    }

    case "colorScale": {
      const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
      // Configure a sensible red (low) → yellow (mid) → green (high) scale
      cf.colorScale.criteria = {
        minimum: { color: "#f4cccc", type: Excel.ConditionalFormatColorCriterionType.lowestValue },
        midpoint: { color: "#fff2cc", type: Excel.ConditionalFormatColorCriterionType.percent, formula: "50" },
        maximum: { color: "#d9ead3", type: Excel.ConditionalFormatColorCriterionType.highestValue },
      };
      break;
    }

    case "dataBar": {
      range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
      break;
    }

    case "iconSet": {
      const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
      // Default to 3-traffic-light icons
      cf.iconSet.style = Excel.IconSet.threeTrafficLights1;
      break;
    }

    case "text": {
      if (text) {
        const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
        cf.textComparison.rule = {
          operator: Excel.ConditionalTextOperator.contains,
          text: text,
        };
        if (format) {
          if (format.fillColor) cf.textComparison.format.fill.color = format.fillColor;
          if (format.fontColor) cf.textComparison.format.font.color = format.fontColor;
        }
      }
      break;
    }
  }

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Applied ${ruleType} conditional format to ${address}`,
  };
}

function buildCellValueRule(
  operator?: string,
  values?: (string | number)[]
): Excel.ConditionalCellValueRule {
  const f1 = String(values?.[0] ?? 0);
  const f2 = values?.[1] !== undefined ? String(values[1]) : undefined;

  const operatorMap: Record<string, Excel.ConditionalCellValueOperator> = {
    greaterThan:           Excel.ConditionalCellValueOperator.greaterThan,
    greaterThanOrEqualTo:  Excel.ConditionalCellValueOperator.greaterThanOrEqual,
    lessThan:              Excel.ConditionalCellValueOperator.lessThan,
    lessThanOrEqualTo:     Excel.ConditionalCellValueOperator.lessThanOrEqual,
    between:               Excel.ConditionalCellValueOperator.between,
    notBetween:            Excel.ConditionalCellValueOperator.notBetween,
    equalTo:               Excel.ConditionalCellValueOperator.equalTo,
    notEqualTo:            Excel.ConditionalCellValueOperator.notEqualTo,
  };

  const rule: Excel.ConditionalCellValueRule = {
    formula1: f1,
    operator: operatorMap[operator ?? "greaterThan"] ?? Excel.ConditionalCellValueOperator.greaterThan,
  };
  if (f2 !== undefined) rule.formula2 = f2;

  return rule;
}


registry.register(meta, handler as any);
export { meta };
