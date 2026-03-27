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

  const range = resolveRange(context, address);
  range.load("values");
  await context.sync();

  const values = range.values;
  options.onProgress?.(`Cleaning ${values.length} rows...`);

  // Apply each operation sequentially
  const cleaned = values.map((row) =>
    row.map((cell) => {
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
      return value.replace(
        /\w\S*/g,
        (txt) => txt.charAt(0).toUpperCase() + txt.slice(1).toLowerCase()
      );
    case "removeNonPrintable":
      return value.replace(/[^\x20-\x7E]/g, "");
    case "normalizeWhitespace":
      return value.replace(/\s+/g, " ").trim();
    default:
      return value;
  }
}


registry.register(meta, handler as any);
export { meta };
