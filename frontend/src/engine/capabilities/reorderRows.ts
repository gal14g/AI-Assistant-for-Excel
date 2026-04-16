/**
 * reorderRows — rearrange rows in-place by criteria.
 *
 * Modes:
 *  - "moveMatching": rows whose conditionColumn satisfies condition go to
 *    destination ("top" or "bottom" of the range). Order within the moved
 *    group and within the kept group is preserved.
 *  - "reverse": flip row order in the range. Headers are respected.
 *  - "clusterByKey": group rows sharing the same value in conditionColumn.
 *    First-appearance order of each key is preserved, but all rows with
 *    a given key become contiguous.
 *
 * Single context.sync — reads values, computes new order in JS, writes back.
 */

import { CapabilityMeta, ReorderRowsParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";
import { normalizeForCompare } from "../utils/normalizeString";
import { parseNumberFlexible } from "../utils/parseNumberFlexible";
import { ensureUnmerged } from "../utils/mergedCells";

const meta: CapabilityMeta = {
  action: "reorderRows",
  description:
    "Rearrange rows in-place: move matching rows, reverse order, or cluster by key",
  mutates: true,
  affectsFormatting: false,
};

type Row = (string | number | boolean | null)[];

function testCondition(
  cell: unknown,
  condition: ReorderRowsParams["condition"],
  value: unknown,
): boolean {
  // Normalize both sides of string comparisons so Hebrew with trailing RTL
  // marks, NBSP, or stray whitespace still matches. Use parseNumberFlexible
  // for numeric comparisons so text-stored numbers ("1,234", "$100") work.
  switch (condition) {
    case "blank":    return normalizeForCompare(cell) === "";
    case "notBlank": return normalizeForCompare(cell) !== "";
    case "equals": {
      if (typeof cell === "number" && typeof value === "number") return cell === value;
      return normalizeForCompare(cell) === normalizeForCompare(value);
    }
    case "notEquals": {
      if (typeof cell === "number" && typeof value === "number") return cell !== value;
      return normalizeForCompare(cell) !== normalizeForCompare(value);
    }
    case "contains":
      return normalizeForCompare(cell).includes(normalizeForCompare(value));
    case "notContains":
      return !normalizeForCompare(cell).includes(normalizeForCompare(value));
    case "greaterThan": {
      const n = parseNumberFlexible(cell);
      const v = parseNumberFlexible(value);
      return n !== null && v !== null && n > v;
    }
    case "lessThan": {
      const n = parseNumberFlexible(cell);
      const v = parseNumberFlexible(value);
      return n !== null && v !== null && n < v;
    }
    default: return false;
  }
}

async function handler(
  context: Excel.RequestContext,
  params: ReorderRowsParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const { range: address, mode, conditionColumn, condition, conditionValue, destination = "top", hasHeaders = true } = params;

  if (options.dryRun) {
    return { stepId: "", status: "success", message: `Would reorderRows mode=${mode} on ${address}` };
  }

  options.onProgress?.(`Reordering rows (${mode})...`);

  const rawRange = resolveRange(context, address);
  const used = rawRange.getUsedRange(false);
  used.load(["values", "rowCount", "columnCount", "address"]);
  await context.sync();

  // Merged cells make reordering unsafe — non-anchor cells come through as
  // null, and the write-back would produce misaligned data. Auto-unmerge.
  const mergeReport = await ensureUnmerged(context, used, {
    operation: "reorderRows",
    policy: "refuseWithError",
  });
  if (mergeReport.error) return mergeReport.error;
  if (mergeReport.hadMerges) {
    used.load(["values"]);
    await context.sync();
  }

  const values = (used.values ?? []) as Row[];
  if (values.length === 0) {
    return { stepId: "", status: "success", message: "Range is empty." };
  }

  const headerRow = hasHeaders ? values[0] : null;
  const dataRows = hasHeaders ? values.slice(1) : values.slice();

  let newOrder: Row[];
  let movedCount = 0;

  switch (mode) {
    case "reverse": {
      newOrder = dataRows.slice().reverse();
      break;
    }
    case "moveMatching": {
      if (conditionColumn === undefined) {
        return { stepId: "", status: "error", message: "moveMatching requires conditionColumn." };
      }
      const matching: Row[] = [];
      const rest: Row[] = [];
      for (const row of dataRows) {
        if (testCondition(row[conditionColumn], condition, conditionValue)) matching.push(row);
        else rest.push(row);
      }
      movedCount = matching.length;
      newOrder = destination === "top" ? [...matching, ...rest] : [...rest, ...matching];
      break;
    }
    case "clusterByKey": {
      if (conditionColumn === undefined) {
        return { stepId: "", status: "error", message: "clusterByKey requires conditionColumn." };
      }
      // Preserve first-appearance order of each unique key. Keys are
      // normalized for bucketing so Hebrew with trailing RTL marks,
      // NBSP-padded, or NFC-variant strings cluster together.
      const buckets = new Map<string, Row[]>();
      const keyOrder: string[] = [];
      for (const row of dataRows) {
        const key = normalizeForCompare(row[conditionColumn]);
        if (!buckets.has(key)) { buckets.set(key, []); keyOrder.push(key); }
        buckets.get(key)!.push(row);
      }
      newOrder = [];
      for (const key of keyOrder) for (const r of buckets.get(key)!) newOrder.push(r);
      movedCount = Array.from(buckets.values()).filter((arr) => arr.length > 1)
        .reduce((sum, arr) => sum + arr.length, 0);
      break;
    }
    default:
      return { stepId: "", status: "error", message: `Unknown mode: ${mode}` };
  }

  // Compose output and write back.
  const output: Row[] = headerRow ? [headerRow, ...newOrder] : newOrder;

  // Anchor write at used-range top-left.
  const sheet = used.worksheet;
  const addrPart = used.address.includes("!") ? used.address.split("!").pop()! : used.address;
  const [topLeftRef] = addrPart.split(":");
  const m = topLeftRef.match(/^([A-Z]+)(\d+)$/);
  if (!m) return { stepId: "", status: "error", message: `Could not parse used range: ${used.address}` };
  let startCol = 0;
  for (let c = 0; c < m[1].length; c++) startCol = startCol * 26 + (m[1].charCodeAt(c) - 64);
  startCol -= 1;
  const startRow = Number(m[2]) - 1;

  try {
    const out = sheet.getRangeByIndexes(startRow, startCol, output.length, used.columnCount);
    out.values = output as unknown as (string | number | boolean)[][];
    await context.sync();
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return { stepId: "", status: "error", message: `Failed to write reordered rows: ${msg}`, error: msg };
  }

  const baseMessage =
    mode === "reverse"
      ? `Reversed ${dataRows.length} rows in ${address}.`
      : mode === "moveMatching"
      ? `Moved ${movedCount} matching row(s) to the ${destination} of ${address}.`
      : `Clustered rows in ${address} by column ${conditionColumn} (${movedCount} row(s) in multi-row groups).`;

  return {
    stepId: "",
    status: "success",
    message: `${baseMessage}${mergeReport.warning ?? ""}`,
    outputs: { range: used.address, movedRowCount: movedCount },
  };
}

registry.register(meta, handler as any);
export { meta };
