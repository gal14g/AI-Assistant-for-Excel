/**
 * extractPattern – Extract emails, phone numbers, URLs, dates, numbers,
 * or any custom regex from a range of text cells.
 */

import { CapabilityMeta, ExtractPatternParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "extractPattern",
  description: "Extract emails, phone numbers, URLs, dates, or custom regex patterns from text cells",
  mutates: true,
  affectsFormatting: false,
};

const BUILT_IN: Record<string, RegExp> = {
  email:    /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g, // eslint-disable-line no-useless-escape
  phone:    /(?:\+?[\d\s\-().]{7,15})/g, // eslint-disable-line no-useless-escape
  url:      /https?:\/\/[^\s"'>]+/g, // eslint-disable-line no-useless-escape
  date:     /\b\d{1,4}[/\-.]\d{1,2}[/\-.]\d{1,4}\b/g,
  number:   /[-+]?\d+(?:[.,]\d+)*/g,
  currency: /[$€£₪¥]?\s*\d+(?:[.,]\d{2,})+/g,
};

async function handler(
  context: Excel.RequestContext,
  params: ExtractPatternParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { sourceRange, pattern, outputRange, allMatches = false } = params;

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would extract "${pattern}" from ${sourceRange}` };
  }

  options.onProgress?.("Reading source range...");
  const srcRng = resolveRange(context, sourceRange);
  const used = srcRng.getUsedRange(false);
  used.load("values");
  await context.sync();

  const vals = (used.values ?? []) as (string | number | boolean | null)[][];
  options.onProgress?.(`Extracting from ${vals.length} rows...`);

  const regex: RegExp = BUILT_IN[pattern]
    ? new RegExp(BUILT_IN[pattern].source, "g")
    : (() => { try { return new RegExp(pattern, "g"); } catch { return /(?:)/g; } })();

  const results: (string | null)[][] = vals.map((row) =>
    row.map((cell) => {
      const text = String(cell ?? "");
      const matches = [...text.matchAll(new RegExp(regex.source, "g"))].map((m) => m[0]);
      if (!matches.length) return null;
      return allMatches ? matches.join(", ") : matches[0];
    }),
  );

  const outRng = resolveRange(context, outputRange);
  outRng.getResizedRange(results.length - 1, results[0].length - 1).values = results as any;
  await context.sync();

  const found = results.flat().filter((v) => v !== null).length;
  return {
    stepId: "",
    status: "success",
    message: `Extracted ${found} match(es) using pattern "${pattern}" from ${vals.length} rows`,
  };
}

registry.register(meta, handler as any);
export { meta };
