/**
 * App – Root component for the Excel AI Copilot task pane.
 */

import React from "react";
import { ChatPanel } from "./components/ChatPanel";
import "./App.css";

// Register all capabilities at startup
import "../engine/capabilities/index";

const App: React.FC = () => {
  return <ChatPanel />;
};

export default App;
