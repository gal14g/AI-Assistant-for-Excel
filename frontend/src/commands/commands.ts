/**
 * Ribbon command handlers.
 *
 * These functions are referenced in the manifest and execute when the user
 * clicks ribbon buttons. Because we use a shared runtime, these handlers
 * have access to the same state as the task pane.
 *
 * Office.js notes:
 * - Functions must be registered with Office.actions.associate().
 * - The function name must match the <FunctionName> in the manifest.
 * - event.completed() must be called to signal the action finished.
 */

import { rollbackLastSnapshot, getSnapshotCount } from "../engine/snapshot";

/* global Office, Excel */

Office.onReady(() => {
  // Register ribbon command handlers
  Office.actions.associate("undoLastRun", undoLastRun);
});

/**
 * Undo the most recent AI Copilot execution.
 * Triggered from the ribbon "Undo Last Run" button.
 */
async function undoLastRun(event: Office.AddinCommands.Event): Promise<void> {
  try {
    if (getSnapshotCount() === 0) {
      // Nothing to undo – silently complete
      event.completed();
      return;
    }

    await Excel.run(async (context) => {
      const result = await rollbackLastSnapshot(context);
      if (result) {
        console.log(`Rolled back plan ${result.planId} (${result.cells.length} ranges)`);
      }
    });
  } catch (err) {
    console.error("Undo failed:", err);
  }

  event.completed();
}
