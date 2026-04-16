/**
 * extractMatchedToNewRow — within-row-duplicate extraction.
 *
 * When `row[keyColumnIndexA] === row[keyColumnIndexB]` (same value in two
 * designated columns of the SAME row), lift the cells at
 * `extractColumnIndexes` into a new row inserted IMMEDIATELY BELOW the
 * matched row. The shared key value is copied into column-A position of the
 * new row so the extracted record is identifiable. Original positions of
 * extracted cells on the matched row are blanked.
 *
 * Useful for normalizing side-by-side comparison data:
 *   interview first/second round, primary/secondary email, request/fulfilled
 *   pairs, old/new values, etc.
 *
 * Example: columns [name, number, price, name2, number2, price2]
 *   Input row:  alice, 1, 5, alice, 1, 6
 *   keyColumnIndexA=0, keyColumnIndexB=3, extractColumnIndexes=[4,5]
 *   Output rows:
 *     alice, 1, 5, alice, __, __    (matched row; extracted cells blanked)
 *     alice, __, __, __, 1, 6       (new row: key in col A, extracted values in their original columns)
 */

import {
  CapabilityMeta,
  ExtractMatchedToNewRowParams,
  StepResult,
  ExecutionOptions,
} from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";
import { normalizeString, normalizeForCompare } from "../utils/normalizeString";
import { ensureUnmerged } from "../utils/mergedCells";
import { expandSnapshotFootprint } from "../snapshot";

const meta: CapabilityMeta = {
  action: "extractMatchedToNewRow",
  description:
    "Split rows where two designated columns match: lift the chosen columns into a new row below, keeping the key value as identifier",
  mutates: true,
  affectsFormatting: false,
};

async function handler(
  context: Excel.RequestContext,
  params: ExtractMatchedToNewRowParams,
  options: ExecutionOptions,
): Promise<StepResult> {
  const {
    sourceRange,
    keyColumnIndexA,
    keyColumnIndexB,
    extractColumnIndexes,
    hasHeaders = true,
    caseSensitive = false,
  } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would extract matching rows in ${sourceRange} (key cols ${keyColumnIndexA}↔${keyColumnIndexB}) into new rows below.`,
    };
  }

  options.onProgress?.("Scanning for within-row column matches...");

  // Clip to used range.
  const rawRange = resolveRange(context, sourceRange);
  const used = rawRange.getUsedRange(false);
  used.load(["values", "rowCount", "columnCount", "address", "worksheet/name"]);
  await context.sync();

  // Merged cells break read-then-write: `range.values` returns nulls in
  // non-anchor cells of a merge. Refuse with a clear error instead of
  // silently destroying data via auto-unmerge.
  const mergeReport = await ensureUnmerged(context, used, {
    operation: "extractMatchedToNewRow",
    policy: "refuseWithError",
  });
  if (mergeReport.error) return mergeReport.error;
  if (mergeReport.hadMerges) {
    used.load(["values"]);
    await context.sync();
  }

  const values = (used.values ?? []) as (string | number | boolean | null)[][];
  if (values.length === 0) {
    return { stepId: "", status: "success", message: "Source range is empty." };
  }

  const origCols = used.columnCount;
  for (const idx of [keyColumnIndexA, keyColumnIndexB, ...extractColumnIndexes]) {
    if (idx < 0 || idx >= origCols) {
      return {
        stepId: "",
        status: "error",
        message: `Column index ${idx} is out of range (source has ${origCols} columns).`,
      };
    }
  }

  const headerRow = hasHeaders ? values[0] : null;
  const dataRows = hasHeaders ? values.slice(1) : values.slice();

  // Equality test — case-insensitive by default for string comparisons.
  // Strings are normalized (trim, NFC, strip bidi/zero-width, collapse
  // whitespace) so copy-pasted Hebrew with RTL marks or NBSP still matches.
  const eq = (a: unknown, b: unknown): boolean => {
    if (a === null || a === undefined || a === "" || b === null || b === undefined || b === "") {
      return false;
    }
    if (typeof a === "string" || typeof b === "string") {
      return caseSensitive
        ? normalizeString(a) === normalizeString(b)
        : normalizeForCompare(a) === normalizeForCompare(b);
    }
    return a === b;
  };

  // Build new data array — keep unmatched rows as-is, rewrite matched rows +
  // insert a new row immediately below.
  const newData: (string | number | boolean | null)[][] = [];
  let matchedCount = 0;
  const extractSet = new Set(extractColumnIndexes);

  for (const row of dataRows) {
    const a = row[keyColumnIndexA];
    const b = row[keyColumnIndexB];
    const isMatch = eq(a, b);

    if (!isMatch) {
      newData.push(row);
      continue;
    }

    matchedCount += 1;
    // Matched row: blank the extracted columns.
    const trimmed: (string | number | boolean | null)[] = [];
    for (let c = 0; c < origCols; c++) {
      trimmed.push(extractSet.has(c) ? null : row[c] ?? null);
    }
    newData.push(trimmed);

    // New row: key value in keyColumnIndexA's position + extracted values in
    // their ORIGINAL column positions (so the wide layout is preserved).
    const newRow: (string | number | boolean | null)[] = [];
    for (let c = 0; c < origCols; c++) {
      if (c === keyColumnIndexA) newRow.push(a ?? null);
      else if (extractSet.has(c)) newRow.push(row[c] ?? null);
      else newRow.push(null);
    }
    newData.push(newRow);
  }

  if (matchedCount === 0) {
    return {
      stepId: "",
      status: "success",
      message: `No matches found — columns ${keyColumnIndexA}/${keyColumnIndexB} never agreed on any row.`,
      outputs: { outputRange: used.address, matchedRowCount: 0 },
    };
  }

  // Compose final grid (header + processed data rows).
  const outputGrid: (string | number | boolean | null)[][] = [];
  if (headerRow) outputGrid.push(headerRow);
  for (const r of newData) outputGrid.push(r);

  // Write back anchored at the used-range top-left. The grid is TALLER than
  // the original (by `matchedCount` rows), so we may overwrite cells below
  // the original source range on the sheet. We accept that — the user's
  // stated intent is "insert new row below", which necessarily displaces
  // downstream content. If they need a safer fallback, they can clone the
  // sheet first.
  const sheet = used.worksheet;
  const addrPart = used.address.includes("!") ? used.address.split("!").pop()! : used.address;
  const [topLeftRef] = addrPart.split(":");
  const m = topLeftRef.match(/^([A-Z]+)(\d+)$/);
  if (!m) {
    return { stepId: "", status: "error", message: `Could not parse used range: ${used.address}` };
  }
  let startCol = 0;
  for (let c = 0; c < m[1].length; c++) startCol = startCol * 26 + (m[1].charCodeAt(c) - 64);
  startCol -= 1;
  const startRow = Number(m[2]) - 1;

  // Expand the snapshot footprint to cover the TALLER output region (extra
  // rows we're about to write below the source). Without this, undo can't
  // reach the inserted rows and leaves them populated.
  const idxToLetters = (idx: number): string => {
    let n = idx + 1;
    let out = "";
    while (n > 0) {
      const rem = (n - 1) % 26;
      out = String.fromCharCode(65 + rem) + out;
      n = Math.floor((n - 1) / 26);
    }
    return out;
  };
  try {
    sheet.load("name");
    await context.sync();
    const footprintAddr = `${sheet.name}!${idxToLetters(startCol)}${startRow + 1}:${idxToLetters(startCol + origCols - 1)}${startRow + outputGrid.length}`;
    await expandSnapshotFootprint(context, [footprintAddr]);
  } catch {
    // Non-fatal.
  }

  try {
    const outRange = sheet.getRangeByIndexes(startRow, startCol, outputGrid.length, origCols);
    outRange.values = outputGrid as unknown as (string | number | boolean)[][];
    await context.sync();
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return {
      stepId: "",
      status: "error",
      message: `Failed to write extracted rows: ${msg}`,
      error: msg,
    };
  }

  const indexToLetters = (idx: number): string => {
    let n = idx + 1;
    let out = "";
    while (n > 0) {
      const rem = (n - 1) % 26;
      out = String.fromCharCode(65 + rem) + out;
      n = Math.floor((n - 1) / 26);
    }
    return out;
  };
  const outputAddr = `${sheet.name}!${indexToLetters(startCol)}${startRow + 1}:${indexToLetters(startCol + origCols - 1)}${startRow + outputGrid.length}`;

  return {
    stepId: "",
    status: "success",
    message: `Extracted ${matchedCount} matched row(s) into new rows. Output: ${outputAddr}.${mergeReport.warning ?? ""}`,
    outputs: { outputRange: outputAddr, matchedRowCount: matchedCount },
  };
}

registry.register(meta, handler as any);
export { meta };
