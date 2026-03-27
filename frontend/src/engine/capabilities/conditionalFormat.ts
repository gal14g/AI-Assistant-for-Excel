/**
 * addConditionalFormat – Apply conditional formatting rules.
 *
 * Office.js notes:
 * - ConditionalFormat API is in ExcelApi 1.6+.
 * - Each rule type (cellValue, colorScale, dataBar, iconSet, text)
 *   has its own creation method and configuration.
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
  const { range: address, ruleType, operator, values, format, text } = params;

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
    case "colorScale": {
      range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
      // Color scale auto-configures with sensible defaults (green-yellow-red)
      break;
    }
    case "dataBar": {
      range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
      break;
    }
    case "iconSet": {
      range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
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
  const rule: Excel.ConditionalCellValueRule = {
    formula1: String(values?.[0] ?? 0),
    operator: Excel.ConditionalCellValueOperator.greaterThan,
  };

  switch (operator) {
    case "greaterThan":
      rule.operator = Excel.ConditionalCellValueOperator.greaterThan;
      break;
    case "lessThan":
      rule.operator = Excel.ConditionalCellValueOperator.lessThan;
      break;
    case "between":
      rule.operator = Excel.ConditionalCellValueOperator.between;
      rule.formula2 = String(values?.[1] ?? 0);
      break;
    case "equalTo":
      rule.operator = Excel.ConditionalCellValueOperator.equalTo;
      break;
  }

  return rule;
}


registry.register(meta, handler as any);
export { meta };
