/**
 * analyzeApi.ts
 *
 * API client for POST /api/analyze.
 * Reads sheet data from the workbook via Office.js and sends it to the
 * backend analytical pipeline.
 */

/* eslint-disable @typescript-eslint/no-explicit-any */

// ── Types ─────────────────────────────────────────────────────────────────────

export interface SheetDataPayload {
  name: string;
  data: unknown[][];    // 2D array — first row is the header row
  headers?: string[];   // optional explicit header list
}

export interface AnalyzeRequest {
  userMessage: string;
  sheets: Record<string, SheetDataPayload>;
  activeSheet?: string;
  conversationHistory?: Array<{ role: string; content: string }>;
}

export interface AnalyzeResponse {
  intent: string;
  strategy?: string;
  message: string;
  results?: Record<string, unknown>;
  needs_clarification: boolean;
  clarification_question?: string;
  warnings: string[];
  execution_log: string[];
  confidence: number;
}

// ── Keyword detection ─────────────────────────────────────────────────────────

/** Returns true when the message is likely an analytical / data-science request. */
export function isAnalyticalRequest(message: string): boolean {
  const lower = message.toLowerCase();
  const keywords = [
    "match",
    "compare",
    "duplicate",
    "dedup",
    "group by",
    "aggregate",
    "sum by",
    "count by",
    "average by",
    "filter where",
    "find all",
    "find rows",
    "clean column",
    "profile",
    "analyse",
    "analyze",
    "show me rows where",
    "rows where",
    "similar",
    "fuzzy",
    "overlap",
    "diff",
  ];
  return keywords.some((kw) => lower.includes(kw));
}

// ── Sheet reading ─────────────────────────────────────────────────────────────

/**
 * Read one sheet from the workbook via Office.js.
 * Returns a SheetDataPayload with the used range as a 2D array
 * (first row = headers).
 */
export async function readSheetData(sheetName: string): Promise<SheetDataPayload> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const usedRange = sheet.getUsedRange();
    usedRange.load("values, address");
    await context.sync();

    const values: unknown[][] = usedRange.values as unknown[][];
    const headers = values.length > 0
      ? values[0].map((v) => (v != null ? String(v) : ""))
      : [];

    return { name: sheetName, data: values, headers };
  });
}

/**
 * Read all sheets in the workbook (up to MAX_SHEETS).
 * Skips sheets that are empty or fail to load.
 */
export async function readAllSheets(
  maxSheets = 5,
): Promise<Record<string, SheetDataPayload>> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    const result: Record<string, SheetDataPayload> = {};
    const toLoad = sheets.items.slice(0, maxSheets);

    for (const sheet of toLoad) {
      try {
        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load("values, isNullObject");
        await context.sync();

        if (usedRange.isNullObject || !usedRange.values || usedRange.values.length === 0) {
          continue;
        }

        const values: unknown[][] = usedRange.values as unknown[][];
        const headers = values[0].map((v) => (v != null ? String(v) : ""));
        result[sheet.name] = { name: sheet.name, data: values, headers };
      } catch {
        // Skip sheets that fail (protected, etc.)
      }
    }

    return result;
  });
}

/**
 * Read specific named sheets from the workbook.
 * Silently skips sheets that don't exist.
 */
export async function readSheetsForRequest(
  sheetNames: string[],
): Promise<Record<string, SheetDataPayload>> {
  const result: Record<string, SheetDataPayload> = {};

  for (const name of sheetNames) {
    try {
      const payload = await readSheetData(name);
      result[name] = payload;
    } catch {
      // sheet not found — skip
    }
  }

  return result;
}

// ── API call ──────────────────────────────────────────────────────────────────

/** POST /api/analyze with the provided request body. */
export async function analyzeRequest(
  request: AnalyzeRequest,
  signal?: AbortSignal,
): Promise<AnalyzeResponse> {
  const response = await fetch("/api/analyze", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(request),
    signal,
  });

  if (!response.ok) {
    let detail = `Server error ${response.status}`;
    try {
      const body = await response.json();
      detail = body?.detail?.message || body?.detail || detail;
    } catch {
      // ignore parse errors
    }
    throw new Error(detail);
  }

  return response.json() as Promise<AnalyzeResponse>;
}
