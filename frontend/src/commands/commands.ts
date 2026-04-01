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

/* global Office */

Office.onReady(() => {
  // Ribbon command handlers are registered here.
  // Undo is handled in the chat panel — no ribbon button needed.
});
