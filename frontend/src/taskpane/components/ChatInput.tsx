/**
 * ChatInput
 *
 * Range reference insertion:
 *   1. Select any cell / range in Excel — hint bar shows the address.
 *   2. Click the input (or it's already focused).
 *   3. Press Ctrl+V (or ⌘V on Mac) → [[Sheet1!A1:C3]] is inserted at the
 *      exact cursor position.  Normal text paste is not affected because
 *      Ctrl+V only intercepts when there is a live Excel selection to insert.
 *      If no selection is known, Ctrl+V falls through to the browser default.
 *
 * Cursor position:
 *   Read directly from e.currentTarget.selectionStart at the moment the key
 *   is pressed — no async gap, no stale ref.
 */

import React, { useState, useRef, useCallback } from "react";

interface Props {
  onSend: (text: string, rangeTokens: { address: string; sheetName: string }[]) => void;
  disabled?: boolean;
  /** Currently selected address in Excel – updated live as the user clicks cells */
  currentSelectionAddress: string | null;
}

function extractRangeTokens(text: string): { address: string; sheetName: string }[] {
  // Matches [[Sheet1!A1:C3]] and also [[[WorkbookName.xlsx]Sheet1!A1:C3]]
  // The inner group allows exactly one ] so it can span [WorkbookName]SheetName!Addr.
  const regex = /\[\[([^\]]*(?:\][^\]]*)?)\]\]/g;
  const tokens: { address: string; sheetName: string }[] = [];
  let match: RegExpExecArray | null;
  while ((match = regex.exec(text)) !== null) {
    const ref  = match[1]; // e.g. "[Sales.xlsx]Sheet1!A:A" or "Sheet1!A:A"
    const bang = ref.indexOf("!");
    tokens.push({
      address:   ref,
      sheetName: bang > -1 ? ref.slice(0, bang).replace(/[[\]']/g, "") : "ActiveSheet",
    });
  }
  return tokens;
}

function spliceAt(
  value: string,
  pos: number,
  insertText: string
): { newValue: string; newPos: number } {
  const before  = value.slice(0, pos);
  const after   = value.slice(pos);
  const prefix  = before.length > 0 && !before.endsWith(" ") ? " " : "";
  const newValue = before + prefix + insertText + " " + after;
  const newPos   = pos + prefix.length + insertText.length + 1;
  return { newValue, newPos };
}

export const ChatInput: React.FC<Props> = ({
  onSend,
  disabled = false,
  currentSelectionAddress,
}) => {
  const [text, setText] = useState("");
  const inputRef = useRef<HTMLTextAreaElement>(null);

  const handleSend = useCallback(() => {
    const trimmed = text.trim();
    if (!trimmed) return;
    onSend(trimmed, extractRangeTokens(trimmed));
    setText("");
  }, [text, onSend]);

  const handleKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
      const ctrl = e.ctrlKey || e.metaKey;

      // Enter → send
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        handleSend();
        return;
      }

      // Ctrl+V → insert current Excel selection at cursor if one exists,
      // otherwise fall through to let the browser handle a normal paste.
      if (ctrl && e.key === "v" && currentSelectionAddress) {
        e.preventDefault();

        const el  = e.currentTarget;
        const pos = el.selectionStart ?? 0;
        const token = `[[${currentSelectionAddress}]]`;

        const { newValue, newPos } = spliceAt(el.value, pos, token);
        setText(newValue);

        requestAnimationFrame(() => {
          if (!inputRef.current) return;
          inputRef.current.setSelectionRange(newPos, newPos);
          inputRef.current.focus();
        });
      }
    },
    [handleSend, currentSelectionAddress]
  );

  const canSend = !disabled && text.trim().length > 0;

  return (
    <div style={{ boxShadow: "0 -1px 0 #e8e8e8", backgroundColor: "#ffffff", padding: "10px 12px" }}>

      {/* Live selection hint */}
      {currentSelectionAddress && (
        <div
          style={{
            fontSize: 11,
            color: "#5b5fc7",
            marginBottom: 4,
            fontFamily: "monospace",
            display: "flex",
            alignItems: "center",
            gap: 4,
          }}
        >
          <span>▸</span>
          <span>{currentSelectionAddress} — copy and paste here to insert at cursor</span>
        </div>
      )}

      {/* Input row */}
      <div style={{ display: "flex", gap: 8, alignItems: "flex-end" }}>
        <textarea
          ref={inputRef}
          value={text}
          onChange={(e) => setText(e.target.value)}
          onKeyDown={handleKeyDown}
          disabled={disabled}
          placeholder="Type a command… select a range in Excel, then copy and paste here to insert reference"
          rows={3}
          style={{
            flex: 1,
            padding: "8px 12px",
            border: "1px solid #e0e0e0",
            borderRadius: 8,
            fontSize: 13,
            fontFamily: "inherit",
            resize: "none",
            outline: "none",
            lineHeight: 1.5,
            transition: "border-color 0.2s",
          }}
        />
        <button
          onClick={handleSend}
          disabled={!canSend}
          style={{
            padding: "8px 16px",
            border: "none",
            borderRadius: 8,
            backgroundColor: canSend ? "#0f6cbd" : "#c8c6c4",
            color: "#fff",
            fontSize: 13,
            fontWeight: 600,
            cursor: canSend ? "pointer" : "default",
            height: 36,
            minWidth: 60,
            alignSelf: "flex-end",
          }}
        >
          Send
        </button>
      </div>

      <div style={{ fontSize: 10, color: "#aaa", marginTop: 4 }}>
        Enter to send · Shift+Enter for new line · Select range in Excel → copy and paste here to insert
      </div>
    </div>
  );
};
