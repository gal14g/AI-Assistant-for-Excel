/**
 * useSelectionTracker
 *
 * Tracks the currently selected range in Excel and exposes a "range clipboard"
 * that the user can fill with Ctrl+C and paste into the chat input with Ctrl+V.
 *
 * Flow
 * ────
 * 1. As the user clicks around in Excel, currentSelectionAddress is kept up
 *    to date (used for the status hint in the input bar).
 * 2. When the user presses Ctrl+C while the chat input is focused,
 *    ChatInput calls captureCurrentSelection() which snapshots the current
 *    address into capturedRange.
 * 3. When the user presses Ctrl+V while the chat input is focused,
 *    ChatInput reads capturedRange and inserts [[address]] at the cursor.
 *    It then calls clearCapturedRange() to reset.
 *
 * No auto-insert happens on cell clicks — the user drives every insertion.
 *
 * Event source
 * ────────────
 * Office.context.document.addHandlerAsync(DocumentSelectionChanged) fires
 * reliably in every Excel version (desktop, web, Mac) without context issues.
 */

import { useState, useEffect } from "react";

export interface WorkbookContext {
  sheetName: string;
  workbookName: string;
}

interface SelectionTrackerResult {
  /** Address currently selected in Excel, e.g. "Sheet1!A1:C3" */
  currentSelectionAddress: string | null;
  /** Workbook + sheet context for the backend request */
  workbookContext: WorkbookContext | null;
}

export function useSelectionTracker(): SelectionTrackerResult {
  const [currentSelectionAddress, setCurrentSelectionAddress] = useState<string | null>(null);
  const [workbookContext, setWorkbookContext] = useState<WorkbookContext | null>(null);

  useEffect(() => {
    if (typeof Office === "undefined" || !Office.context?.document) return;

    const handleSelectionChange = async () => {
      try {
        await Excel.run(async (ctx) => {
          const ws = ctx.workbook.worksheets.getActiveWorksheet();
          ws.load("name");
          ctx.workbook.load("name");
          const range = ctx.workbook.getSelectedRange();
          range.load("address");
          await ctx.sync();

          const workbookName = ctx.workbook.name ?? "";
          const sheetName   = ws.name;

          // range.address is always fully qualified: "Sheet1!A1:B5"
          // Strip surrounding quotes from sheet names that contain spaces:
          // "'My Sheet'!A1:B5" → "My Sheet!A1:B5"
          const sheetAddress = (range.address ?? "").replace(/^'(.+)'!/, "$1!");

          // Build the display address.
          // Always include workbook name so cross-sheet references are
          // unambiguous when the LLM sees them.
          // Format: [WorkbookName.xlsx]Sheet1!A1:B5
          const displayAddress = workbookName
            ? `[${workbookName}]${sheetAddress}`
            : sheetAddress;

          setCurrentSelectionAddress(displayAddress);
          setWorkbookContext({ sheetName, workbookName });
        });
      } catch {
        // Not in Excel context — ignore.
      }
    };

    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      handleSelectionChange
    );

    return () => {
      Office.context.document.removeHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        { handler: handleSelectionChange }
      );
    };
  }, []);

  return {
    currentSelectionAddress,
    workbookContext,
  };
}
