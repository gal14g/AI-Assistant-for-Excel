/**
 * formatCells – Apply visual formatting to a range.
 *
 * Supports: bold, italic, underline, strikethrough, font size/family/color,
 * fill color, horizontal/vertical alignment, text wrap, and borders.
 * Only the properties that are set are changed — everything else is untouched.
 */

import { CapabilityMeta, FormatCellsParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "formatCells",
  description: "Format cell appearance — font, colors, borders, alignment",
  mutates: false,
  affectsFormatting: true,
};

async function handler(
  context: Excel.RequestContext,
  params: FormatCellsParams,
  options: ExecutionOptions
): Promise<StepResult> {
  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would format ${params.range}` };
  }

  options.onProgress?.("Applying cell formatting...");

  const range = resolveRange(context, params.range);
  const fmt = range.format;

  // Font properties
  if (params.bold !== undefined) fmt.font.bold = params.bold;
  if (params.italic !== undefined) fmt.font.italic = params.italic;
  if (params.underline !== undefined) {
    fmt.font.underline = params.underline ? "Single" : "None";
  }
  if (params.strikethrough !== undefined) fmt.font.strikethrough = params.strikethrough;
  if (params.fontSize !== undefined) fmt.font.size = params.fontSize;
  if (params.fontFamily !== undefined) fmt.font.name = params.fontFamily;
  if (params.fontColor !== undefined) fmt.font.color = params.fontColor;

  // Fill
  if (params.fillColor !== undefined) {
    if (params.fillColor === "" || params.fillColor.toLowerCase() === "none") {
      fmt.fill.clear();
    } else {
      fmt.fill.color = params.fillColor;
    }
  }

  // Alignment
  if (params.horizontalAlignment !== undefined) {
    const map: Record<string, Excel.HorizontalAlignment> = {
      left: "Left" as Excel.HorizontalAlignment,
      center: "Center" as Excel.HorizontalAlignment,
      right: "Right" as Excel.HorizontalAlignment,
      justify: "Justify" as Excel.HorizontalAlignment,
    };
    fmt.horizontalAlignment = map[params.horizontalAlignment] ?? ("General" as Excel.HorizontalAlignment);
  }
  if (params.verticalAlignment !== undefined) {
    const map: Record<string, Excel.VerticalAlignment> = {
      top: "Top" as Excel.VerticalAlignment,
      middle: "Center" as Excel.VerticalAlignment,
      bottom: "Bottom" as Excel.VerticalAlignment,
    };
    fmt.verticalAlignment = map[params.verticalAlignment] ?? ("Bottom" as Excel.VerticalAlignment);
  }
  if (params.wrapText !== undefined) fmt.wrapText = params.wrapText;

  // Borders
  if (params.borders) {
    const styleMap: Record<string, Excel.BorderLineStyle> = {
      thin: "Continuous" as Excel.BorderLineStyle,
      medium: "Continuous" as Excel.BorderLineStyle,
      thick: "Continuous" as Excel.BorderLineStyle,
      dashed: "Dash" as Excel.BorderLineStyle,
      dotted: "Dot" as Excel.BorderLineStyle,
      double: "Double" as Excel.BorderLineStyle,
      none: "None" as Excel.BorderLineStyle,
    };
    const weightMap: Record<string, Excel.BorderWeight> = {
      thin: "Thin" as Excel.BorderWeight,
      medium: "Medium" as Excel.BorderWeight,
      thick: "Thick" as Excel.BorderWeight,
      dashed: "Thin" as Excel.BorderWeight,
      dotted: "Thin" as Excel.BorderWeight,
      double: "Thin" as Excel.BorderWeight,
      none: "Thin" as Excel.BorderWeight,
    };
    const lineStyle = styleMap[params.borders.style] ?? ("Continuous" as Excel.BorderLineStyle);
    const weight = weightMap[params.borders.style] ?? ("Thin" as Excel.BorderWeight);
    const color = params.borders.color ?? "#000000";
    const edges = params.borders.edges ?? ["all"];

    const borderIndexMap: Record<string, Excel.BorderIndex[]> = {
      top: ["EdgeTop" as Excel.BorderIndex],
      bottom: ["EdgeBottom" as Excel.BorderIndex],
      left: ["EdgeLeft" as Excel.BorderIndex],
      right: ["EdgeRight" as Excel.BorderIndex],
      all: [
        "EdgeTop" as Excel.BorderIndex,
        "EdgeBottom" as Excel.BorderIndex,
        "EdgeLeft" as Excel.BorderIndex,
        "EdgeRight" as Excel.BorderIndex,
        "InsideHorizontal" as Excel.BorderIndex,
        "InsideVertical" as Excel.BorderIndex,
      ],
      outside: [
        "EdgeTop" as Excel.BorderIndex,
        "EdgeBottom" as Excel.BorderIndex,
        "EdgeLeft" as Excel.BorderIndex,
        "EdgeRight" as Excel.BorderIndex,
      ],
      inside: [
        "InsideHorizontal" as Excel.BorderIndex,
        "InsideVertical" as Excel.BorderIndex,
      ],
    };

    for (const edge of edges) {
      const indices = borderIndexMap[edge] ?? [];
      for (const idx of indices) {
        const border = fmt.borders.getItem(idx);
        border.style = lineStyle;
        border.weight = weight;
        border.color = color;
      }
    }
  }

  await context.sync();

  const changes: string[] = [];
  if (params.bold !== undefined) changes.push("bold");
  if (params.italic !== undefined) changes.push("italic");
  if (params.fillColor !== undefined) changes.push("fill");
  if (params.fontColor !== undefined) changes.push("font color");
  if (params.fontSize !== undefined) changes.push(`size ${params.fontSize}`);
  if (params.borders) changes.push("borders");
  if (params.horizontalAlignment !== undefined) changes.push(`align ${params.horizontalAlignment}`);

  return {
    stepId: "",
    status: "success",
    message: `Formatted ${params.range}: ${changes.join(", ") || "applied"}`,
  };
}

registry.register(meta, handler as any);
export { meta };
