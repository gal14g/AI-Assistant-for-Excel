/**
 * ChatInput
 *
 * Range reference insertion:
 *   1. Select any cell / range in Excel — hint bar shows the address.
 *   2. Click the input (or it's already focused).
 *   3. Press Ctrl+V (or Cmd+V on Mac) — [[Sheet1!A1:C3]] is inserted at the
 *      exact cursor position.  Normal text paste is not affected because
 *      Ctrl+V only intercepts when there is a live Excel selection to insert.
 *      If no selection is known, Ctrl+V falls through to the browser default.
 *
 * Toolbar row: Presets menu, Save preset, Undo, Send/Stop toggle.
 */

import React, { useState, useRef, useCallback, useEffect } from "react";
import { Preset } from "../../services/api";
import { PresetMenu } from "./PresetMenu";

interface Props {
  onSend: (text: string, rangeTokens: { address: string; sheetName: string }[]) => void;
  onStop: () => void;
  onUndo?: () => void;
  onSavePreset?: () => void;
  onDeletePreset?: (presetId: string) => void;
  onRenamePreset?: (presetId: string, newName: string) => void;
  disabled?: boolean;
  isLoading?: boolean;
  canUndo?: boolean;
  presets?: Preset[];
  currentSelectionAddress: string | null;
  /** When set, prefills the input and focuses it. Change seq to re-trigger for same text. */
  prefillText?: { text: string; seq: number };
}

function extractRangeTokens(text: string): { address: string; sheetName: string }[] {
  const regex = /\[\[([^\]]*(?:\][^\]]*)?)\]\]/g;
  const tokens: { address: string; sheetName: string }[] = [];
  let match: RegExpExecArray | null;
  while ((match = regex.exec(text)) !== null) {
    const ref  = match[1];
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
  onStop,
  onUndo,
  onSavePreset,
  onDeletePreset,
  onRenamePreset,
  disabled = false,
  isLoading = false,
  canUndo = false,
  presets = [],
  currentSelectionAddress,
  prefillText,
}) => {
  const [text, setText] = useState("");
  const inputRef = useRef<HTMLTextAreaElement>(null);
  const [presetMenuOpen, setPresetMenuOpen] = useState(false);

  // Prefill the input when an undo restores the previous user message.
  // eslint-disable-next-line react-hooks/set-state-in-effect -- Intentional: sync prefill from parent
  useEffect(() => {
    if (prefillText && prefillText.text) {
      setText(prefillText.text); // eslint-disable-line react-hooks/set-state-in-effect
      requestAnimationFrame(() => inputRef.current?.focus());
    }
  }, [prefillText?.seq]); // eslint-disable-line react-hooks/exhaustive-deps

  const handleSend = useCallback(() => {
    const trimmed = text.trim();
    if (!trimmed) return;
    onSend(trimmed, extractRangeTokens(trimmed));
    setText("");
  }, [text, onSend]);

  const handleKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
      const ctrl = e.ctrlKey || e.metaKey;

      // Enter -> send (only when not loading)
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        if (!isLoading) handleSend();
        return;
      }

      // Ctrl+V -> insert current Excel selection at cursor if one exists
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
    [handleSend, currentSelectionAddress, isLoading]
  );

  /** Load preset → prefill the input so user can review/edit before sending */
  const handlePresetSelect = useCallback((preset: Preset) => {
    setText(preset.userMessage);
    requestAnimationFrame(() => inputRef.current?.focus());
  }, []);

  const canSend = !disabled && !isLoading && text.trim().length > 0;

  // Shared compact button style
  const btnBase: React.CSSProperties = {
    fontSize: 11,
    padding: "4px 8px",
    borderRadius: 4,
    border: "1px solid #d1d1d1",
    backgroundColor: "#fff",
    cursor: "pointer",
    display: "inline-flex",
    alignItems: "center",
    gap: 3,
    whiteSpace: "nowrap",
  };

  return (
    <div dir="auto" style={{ boxShadow: "0 -1px 0 #e8e8e8", backgroundColor: "#ffffff", padding: "10px 12px" }}>

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
          <span>&#9658;</span>
          <span>{currentSelectionAddress} — Ctrl+V to insert</span>
        </div>
      )}

      {/* Textarea */}
      <textarea
        ref={inputRef}
        value={text}
        onChange={(e) => setText(e.target.value)}
        onKeyDown={handleKeyDown}
        disabled={disabled}
        placeholder="Type a command... select a range in Excel, then Ctrl+V to insert reference"
        rows={3}
        dir="auto"
        style={{
          width: "100%",
          boxSizing: "border-box",
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

      {/* Toolbar row */}
      <div style={{ display: "flex", alignItems: "center", gap: 6, padding: "4px 0 0 0", position: "relative" }}>

        {/* Presets menu button */}
        <button
          onClick={() => setPresetMenuOpen((prev) => !prev)}
          style={{
            ...btnBase,
            backgroundColor: presetMenuOpen ? "#f0f4ff" : "#fff",
            borderColor: presetMenuOpen ? "#5b5fc7" : "#d1d1d1",
            color: presetMenuOpen ? "#5b5fc7" : undefined,
          }}
          title="Saved presets"
        >
          &#128203; Presets
        </button>

        {/* Floating preset menu */}
        {presetMenuOpen && (
          <PresetMenu
            presets={presets}
            onSelect={handlePresetSelect}
            onDelete={(id) => { onDeletePreset?.(id); }}
            onRename={(id, name) => { onRenamePreset?.(id, name); }}
            onClose={() => setPresetMenuOpen(false)}
          />
        )}

        {/* Save preset button */}
        {!isLoading && onSavePreset && (
          <button onClick={onSavePreset} style={btnBase} title="Save last plan as preset">
            &#128190; Save
          </button>
        )}

        {/* Undo button */}
        {canUndo && onUndo && (
          <button onClick={onUndo} style={btnBase}>
            &#8617; Undo
          </button>
        )}

        {/* Send / Stop toggle — pushed to the right */}
        {isLoading ? (
          <button
            onClick={onStop}
            title="Stop generating"
            style={{
              marginLeft: "auto",
              width: 32,
              height: 32,
              borderRadius: "50%",
              border: "2px solid #c50f1f",
              backgroundColor: "#fff",
              cursor: "pointer",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              padding: 0,
              transition: "background-color 0.15s",
            }}
            onMouseEnter={(e) => { (e.currentTarget as HTMLButtonElement).style.backgroundColor = "#fdf3f3"; }}
            onMouseLeave={(e) => { (e.currentTarget as HTMLButtonElement).style.backgroundColor = "#fff"; }}
          >
            <div style={{
              width: 10,
              height: 10,
              backgroundColor: "#c50f1f",
              borderRadius: 2,
            }} />
          </button>
        ) : (
          <button
            onClick={handleSend}
            disabled={!canSend}
            style={{
              ...btnBase,
              marginLeft: "auto",
              border: "none",
              backgroundColor: canSend ? "#0f6cbd" : "#c8c6c4",
              color: "#fff",
              fontWeight: 600,
              padding: "4px 12px",
              cursor: canSend ? "pointer" : "default",
            }}
          >
            Send &#9654;
          </button>
        )}
      </div>

    </div>
  );
};
