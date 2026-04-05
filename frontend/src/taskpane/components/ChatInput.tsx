/**
 * ChatInput – Integrated textarea + toolbar (like modern chat apps).
 *
 * Range reference insertion:
 *   Select any cell/range in Excel → press Ctrl+V (or Cmd+V on Mac) to
 *   insert [[Sheet1!A1:C3]] at the cursor. Normal text paste still works
 *   when there is no Excel selection.
 */

import React, { useState, useRef, useCallback, useEffect } from "react";
import {
  Send16Filled, Stop16Filled, ArrowUndo16Regular,
  Save16Regular, List16Regular,
} from "@fluentui/react-icons";
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
  /** When set, prefills the input and focuses it. Change seq to re-trigger. */
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

function spliceAt(value: string, pos: number, insertText: string): { newValue: string; newPos: number } {
  const before  = value.slice(0, pos);
  const after   = value.slice(pos);
  const prefix  = before.length > 0 && !before.endsWith(" ") ? " " : "";
  const newValue = before + prefix + insertText + " " + after;
  const newPos   = pos + prefix.length + insertText.length + 1;
  return { newValue, newPos };
}

export const ChatInput: React.FC<Props> = ({
  onSend, onStop, onUndo, onSavePreset, onDeletePreset, onRenamePreset,
  disabled = false, isLoading = false, canUndo = false,
  presets = [], currentSelectionAddress, prefillText,
}) => {
  const [text, setText] = useState("");
  const inputRef = useRef<HTMLTextAreaElement>(null);
  const [presetMenuOpen, setPresetMenuOpen] = useState(false);

  // eslint-disable-next-line react-hooks/set-state-in-effect -- Intentional: sync prefill from parent
  useEffect(() => {
    if (prefillText && prefillText.text) {
      setText(prefillText.text); // eslint-disable-line react-hooks/set-state-in-effect
      requestAnimationFrame(() => inputRef.current?.focus());
    }
  }, [prefillText?.seq]); // eslint-disable-line react-hooks/exhaustive-deps

  // Auto-resize the textarea up to the CSS max-height
  useEffect(() => {
    const el = inputRef.current;
    if (!el) return;
    el.style.height = "auto";
    el.style.height = `${Math.min(el.scrollHeight, 160)}px`;
  }, [text]);

  const handleSend = useCallback(() => {
    const trimmed = text.trim();
    if (!trimmed) return;
    onSend(trimmed, extractRangeTokens(trimmed));
    setText("");
  }, [text, onSend]);

  const handleKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
      const ctrl = e.ctrlKey || e.metaKey;
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        if (!isLoading) handleSend();
        return;
      }
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

  const handlePresetSelect = useCallback((preset: Preset) => {
    setText(preset.userMessage);
    requestAnimationFrame(() => inputRef.current?.focus());
  }, []);

  const canSend = !disabled && !isLoading && text.trim().length > 0;

  return (
    <div dir="auto" className="cc-input-wrap">
      {currentSelectionAddress && (
        <div className="cc-selection-hint">
          <span>▸</span>
          <code>{currentSelectionAddress}</code>
          <span>— Ctrl+V to insert</span>
        </div>
      )}

      <div className="cc-input-box">
        <textarea
          ref={inputRef}
          className="cc-input-textarea"
          value={text}
          onChange={(e) => setText(e.target.value)}
          onKeyDown={handleKeyDown}
          disabled={disabled}
          placeholder="Ask Copilot… select a range in Excel, then Ctrl+V to insert"
          rows={1}
          dir="auto"
        />

        <div className="cc-input-toolbar">
          <button
            className={`cc-btn icon sm ${presetMenuOpen ? "active" : ""}`}
            onClick={() => setPresetMenuOpen((prev) => !prev)}
            title="Saved presets"
            aria-label="Presets"
          >
            <List16Regular />
          </button>
          {presetMenuOpen && (
            <PresetMenu
              presets={presets}
              onSelect={handlePresetSelect}
              onDelete={(id) => { onDeletePreset?.(id); }}
              onRename={(id, name) => { onRenamePreset?.(id, name); }}
              onClose={() => setPresetMenuOpen(false)}
            />
          )}

          {!isLoading && onSavePreset && (
            <button className="cc-btn icon sm" onClick={onSavePreset} title="Save last plan as preset" aria-label="Save preset">
              <Save16Regular />
            </button>
          )}

          {canUndo && onUndo && (
            <button className="cc-btn icon sm" onClick={onUndo} title="Undo last execution" aria-label="Undo">
              <ArrowUndo16Regular />
            </button>
          )}

          <div className="cc-input-toolbar-spacer" />

          {isLoading ? (
            <button className="cc-btn stop" onClick={onStop} title="Stop generating" aria-label="Stop">
              <Stop16Filled />
            </button>
          ) : (
            <button className="cc-btn send" onClick={handleSend} disabled={!canSend} title="Send (Enter)" aria-label="Send">
              <Send16Filled />
            </button>
          )}
        </div>
      </div>
    </div>
  );
};
