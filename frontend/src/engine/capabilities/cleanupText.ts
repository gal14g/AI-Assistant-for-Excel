/**
 * cleanupText – Apply text cleanup operations to a range.
 *
 * Reads values, applies transformations in JS, writes back clean values.
 * This is a computed operation (not formula-based) because text cleanup
 * is typically a one-time operation.
 *
 * NOTE: Writes values only — formatting is fully preserved.
 */

import { CapabilityMeta, CleanupTextParams, CleanupOperation, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "cleanupText",
  description: "Clean up text values (trim, case, whitespace, etc.)",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: CleanupTextParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const { range: address, operations, outputRange } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would apply [${operations.join(", ")}] to ${address}`,
    };
  }

  options.onProgress?.("Reading values for cleanup...");

  // Use getUsedRange so full-column refs like "A:D" don't load 1M rows.
  // Office.js proxy errors only surface at context.sync() time, so the sync
  // must be inside the try block for the catch to work.
  const raw = resolveRange(context, address);
  let range: Excel.Range;
  try {
    const used = raw.getUsedRange(false);
    used.load("values");
    await context.sync();
    range = used;
  } catch {
    raw.load("values");
    await context.sync();
    range = raw;
  }

  const values = range.values ?? [];
  options.onProgress?.(`Cleaning ${values.length} rows...`);

  // Apply each operation sequentially
  // Guard each row with ?? [] in case Office.js returns a sparse/null row
  const cleaned = values.map((row) =>
    (row ?? []).map((cell) => {
      if (typeof cell !== "string") return cell;
      let val = cell;
      for (const op of operations) {
        val = applyOperation(val, op);
      }
      return val;
    })
  );

  // Write to output range or in-place
  const targetAddress = outputRange ?? address;
  const target = resolveRange(context, targetAddress);
  target.values = cleaned;
  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Applied [${operations.join(", ")}] to ${values.length} rows in ${targetAddress}`,
    outputs: { outputRange: targetAddress },
  };
}

function applyOperation(value: string, operation: CleanupOperation): string {
  switch (operation) {
    case "trim":
      return value.trim();
    case "lowercase":
      return value.toLowerCase();
    case "uppercase":
      return value.toUpperCase();
    case "properCase":
      // Use (^|\s)\S rather than \w — \w only matches ASCII [A-Za-z0-9_]
      // and silently skips Hebrew and other non-Latin characters.
      // This approach capitalises the first character of every whitespace-delimited
      // word regardless of script. For Hebrew, toUpperCase() is a safe no-op.
      return value.replace(/(^|\s)(\S)/g, (_, space, char) => space + char.toUpperCase());
    case "removeNonPrintable":
      // Remove only actual ASCII control characters (U+0000–U+001F except
      // tab U+0009, newline U+000A, carriage return U+000D, and DEL U+007F).
      // The old regex [^\x20-\x7E] also matched ALL non-ASCII including Hebrew,
      // Arabic, emoji, etc. — this version preserves all legitimate Unicode text.
      // eslint-disable-next-line no-control-regex
      return value.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, "");
    case "normalizeWhitespace":
      return value.replace(/\s+/g, " ").trim();
    default:
      return value;
  }
}


registry.register(meta, handler as any);
export { meta };
