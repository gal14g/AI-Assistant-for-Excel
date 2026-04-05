/**
 * SuggestedPrompts – Clickable prompt chips shown when the chat is fresh.
 */

import React from "react";

interface Props {
  onSelect: (prompt: string) => void;
}

const PROMPTS: string[] = [
  "Summarize this sheet",
  "Create a pivot table from the data",
  "Highlight duplicates in column A",
  "Sort by the first column descending",
];

export const SuggestedPrompts: React.FC<Props> = React.memo(({ onSelect }) => {
  if (PROMPTS.length === 0) return null;
  return (
    <div dir="auto" className="cc-suggested">
      <div className="cc-suggested-label">Try asking</div>
      <div className="cc-suggested-chips">
        {PROMPTS.map((prompt) => (
          <button
            key={prompt}
            className="cc-suggested-chip"
            onClick={() => onSelect(prompt)}
          >
            {prompt}
          </button>
        ))}
      </div>
    </div>
  );
});
SuggestedPrompts.displayName = "SuggestedPrompts";
