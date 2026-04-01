/**
 * SuggestedPrompts – Clickable prompt chips shown when the chat is fresh.
 * Mirrors the "Try asking" prompts in Microsoft Copilot for Excel.
 */

import React from "react";

interface Props {
  onSelect: (prompt: string) => void;
}

const PROMPTS: string[] = [
  // Add your own suggested prompts here
];

export const SuggestedPrompts: React.FC<Props> = React.memo(({ onSelect }) => (
  <div dir="auto" style={{ padding: "8px 12px 12px" }}>
    <div style={{ fontSize: 11, color: "#616161", marginBottom: 8, fontWeight: 500 }}>
      Try asking
    </div>
    <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
      {PROMPTS.map((prompt) => (
        <button
          key={prompt}
          onClick={() => onSelect(prompt)}
          style={{
            padding: "5px 12px",
            border: "1px solid #e0e0e0",
            borderRadius: 16,
            backgroundColor: "#ffffff",
            color: "#242424",
            fontSize: 12,
            cursor: "pointer",
            transition: "background 0.15s, border-color 0.15s",
            lineHeight: 1.4,
            textAlign: "left",
          }}
          onMouseEnter={(e) => {
            (e.currentTarget as HTMLButtonElement).style.backgroundColor = "#f0f4ff";
            (e.currentTarget as HTMLButtonElement).style.borderColor = "#5b5fc7";
          }}
          onMouseLeave={(e) => {
            (e.currentTarget as HTMLButtonElement).style.backgroundColor = "#ffffff";
            (e.currentTarget as HTMLButtonElement).style.borderColor = "#e0e0e0";
          }}
        >
          {prompt}
        </button>
      ))}
    </div>
  </div>
));
