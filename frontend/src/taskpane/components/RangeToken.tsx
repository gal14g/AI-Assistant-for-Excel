/**
 * RangeToken – Inline chip representing a selected Excel range.
 *
 * Displays as [[Sheet1!A1:C20]] with a dismiss button.
 */

import React from "react";
import { RangeToken as RangeTokenType } from "../../engine/types";

interface Props {
  token: RangeTokenType;
  onRemove: (id: string) => void;
}

export const RangeTokenChip: React.FC<Props> = ({ token, onRemove }) => {
  return (
    <span
      style={{
        display: "inline-flex",
        alignItems: "center",
        gap: 4,
        padding: "2px 8px",
        borderRadius: 12,
        backgroundColor: "#e8f4e8",
        border: "1px solid #4caf50",
        fontSize: 12,
        fontFamily: "monospace",
        color: "#2e7d32",
        whiteSpace: "nowrap",
      }}
    >
      {token.display}
      <button
        onClick={() => onRemove(token.id)}
        style={{
          background: "none",
          border: "none",
          cursor: "pointer",
          padding: 0,
          fontSize: 14,
          color: "#888",
          lineHeight: 1,
        }}
        title="Remove range reference"
        aria-label={`Remove ${token.display}`}
      >
        x
      </button>
    </span>
  );
};
