/**
 * Entry point for the AI Assistant For Excel task pane.
 *
 * Office.js notes:
 * - Office.onReady() must resolve before any Office.js API calls.
 * - We render React after Office is ready.
 * - The shared runtime keeps this JS context alive even if the task pane closes.
 */

import React from "react";
import { createRoot } from "react-dom/client";
import App from "./taskpane/App";

/* global Office */

function renderApp() {
  const container = document.getElementById("root");
  console.log("root element:", container);

  if (container) {
    const root = createRoot(container);
    console.log("rendering app now");
    root.render(<App />);
  } else {
    console.error("No #root element found.");
  }
}

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady()
    .then((info) => {
      console.log("Office ready:", info);
      renderApp();
    })
    .catch((err) => {
      console.error("Office.onReady failed:", err);
      renderApp();
    });
} else {
  console.log("Office not available, rendering in browser mode");
  renderApp();
}
