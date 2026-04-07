/**
 * Workbook Snapshot — builds a lightweight, LLM-friendly summary of the
 * open workbook so the planner can reason about the actual data (not just
 * sheet names).
 *
 * For every worksheet (up to MAX_SHEETS), we capture:
 *   - sheetName
 *   - headers (first row of the used range)
 *   - rowCount, columnCount (used-range dimensions)
 *   - sampleRows (first N data rows — after the header)
 *   - dtypes (inferred per column: number|date|text|boolean|mixed|empty)
 *
 * Designed to be cheap: one Excel.run, one load+sync per sheet, cells
 * capped. Safe to call on every send; typical cost is a few hundred
 * milliseconds on a medium workbook.
 */
/* global Excel */

export interface SheetSnapshot {
  sheetName: string;
  rowCount: number;
  columnCount: number;
  headers: string[];
  sampleRows: (string | number | boolean | null)[][];
  dtypes: string[]; // "number" | "date" | "text" | "boolean" | "mixed" | "empty"
  /** Address of the used-range's top-left cell, e.g. "A1" or "C5" — tables
   *  don't always start at A1, and the LLM needs to know the real anchor. */
  anchorCell: string;
  /** Full used-range address, e.g. "Sheet1!C5:H40". */
  usedRangeAddress: string;
}

export interface WorkbookSnapshot {
  sheets: SheetSnapshot[];
  truncated: boolean; // true if we stopped at MAX_SHEETS
}

const MAX_SHEETS = 15;
const MAX_COLUMNS = 30;
const MAX_SAMPLE_ROWS = 5;

type Cell = string | number | boolean | null;

function inferDtype(values: Cell[]): string {
  let seenNumber = false;
  let seenText = false;
  let seenBool = false;
  let seenDate = false;
  let seenNonEmpty = false;

  const datePattern = /^\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}$/;

  for (const v of values) {
    if (v === null || v === "" || v === undefined) continue;
    seenNonEmpty = true;
    if (typeof v === "number") {
      seenNumber = true;
    } else if (typeof v === "boolean") {
      seenBool = true;
    } else if (typeof v === "string") {
      if (datePattern.test(v.trim())) seenDate = true;
      else seenText = true;
    }
  }

  if (!seenNonEmpty) return "empty";
  const hits = [seenNumber, seenText, seenBool, seenDate].filter(Boolean).length;
  if (hits > 1) return "mixed";
  if (seenNumber) return "number";
  if (seenDate) return "date";
  if (seenBool) return "boolean";
  return "text";
}

function toHeader(v: Cell, col: number): string {
  if (v === null || v === undefined || v === "") return `Col${col + 1}`;
  return String(v).slice(0, 80);
}

/** Extract the top-left cell from a range address like "Sheet1!C5:H40" → "C5",
 *  or "'My Sheet'!C5:H40" → "C5". Falls back to the raw input on parse failure. */
function topLeftCell(address: string): string {
  const idx = address.lastIndexOf("!");
  const body = idx >= 0 ? address.slice(idx + 1) : address;
  const colonIdx = body.indexOf(":");
  return colonIdx >= 0 ? body.slice(0, colonIdx) : body;
}

/**
 * Build a snapshot of the whole workbook. Never throws — returns null
 * if Excel context is unavailable.
 */
export async function buildWorkbookSnapshot(): Promise<WorkbookSnapshot | null> {
  try {
    const snapshot: WorkbookSnapshot = { sheets: [], truncated: false };

    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name,items/visibility");
      await context.sync();

      const visible = sheets.items.filter((s) => s.visibility === "Visible");
      const picked = visible.slice(0, MAX_SHEETS);
      if (visible.length > MAX_SHEETS) snapshot.truncated = true;

      // Phase 1: load used-range addresses for each sheet
      const used = picked.map((s) => {
        const ur = s.getUsedRangeOrNullObject(true);
        ur.load(["address", "rowCount", "columnCount", "isNullObject"]);
        return { sheet: s, used: ur };
      });
      await context.sync();

      // Phase 2: crop each used range to (header + MAX_SAMPLE_ROWS) × MAX_COLUMNS
      // and load its values in one shot.
      const probes = used.map((u) => {
        if (u.used.isNullObject) {
          return { crop: null, rowCount: 0, colCount: 0, anchorCell: "A1", usedRangeAddress: "" };
        }
        const totalRows = u.used.rowCount;
        const totalCols = Math.min(u.used.columnCount, MAX_COLUMNS);
        // We want the first (1 + sample) rows — but capped at totalRows.
        const wantRows = Math.min(totalRows, 1 + MAX_SAMPLE_ROWS);
        // getResizedRange takes deltas relative to the current range size.
        const crop = u.sheet
          .getRange(u.used.address)
          .getResizedRange(wantRows - totalRows, totalCols - u.used.columnCount);
        crop.load("values");
        return {
          crop,
          rowCount: totalRows,
          colCount: totalCols,
          anchorCell: topLeftCell(u.used.address),
          usedRangeAddress: u.used.address,
        };
      });
      await context.sync();

      for (let i = 0; i < picked.length; i++) {
        const p = probes[i];
        const sheetName = picked[i].name;
        if (!p.crop) {
          snapshot.sheets.push({
            sheetName,
            rowCount: 0,
            columnCount: 0,
            headers: [],
            sampleRows: [],
            dtypes: [],
            anchorCell: "A1",
            usedRangeAddress: "",
          });
          continue;
        }
        const allValues = (p.crop.values ?? []) as Cell[][];
        const headerRow = (allValues[0] ?? []) as Cell[];
        const headers = headerRow.map((v, ci) => toHeader(v, ci));
        const sampleRows = allValues.slice(1);

        // Infer dtypes from sample rows (or from header row if no body)
        const dtypes: string[] = [];
        for (let ci = 0; ci < headers.length; ci++) {
          const colVals: Cell[] = sampleRows.map((r) => (r?.[ci] ?? null) as Cell);
          dtypes.push(inferDtype(colVals));
        }

        // Truncate sample row strings (LLM context control)
        const trimmedSample: (string | number | boolean | null)[][] = sampleRows.map((r) =>
          r.slice(0, headers.length).map((v) => (typeof v === "string" ? v.slice(0, 60) : (v as Cell))),
        );

        snapshot.sheets.push({
          sheetName,
          rowCount: p.rowCount,
          columnCount: p.colCount,
          headers,
          sampleRows: trimmedSample,
          dtypes,
          anchorCell: p.anchorCell,
          usedRangeAddress: p.usedRangeAddress,
        });
      }
    });

    return snapshot;
  } catch {
    return null;
  }
}
