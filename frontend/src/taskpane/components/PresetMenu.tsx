/**
 * PresetMenu – Floating popup panel for managing presets.
 *
 * Opens from the "Presets" button in the toolbar. Shows a scrollable list of
 * saved presets with the ability to:
 *   - Click a preset to load its user message into the input
 *   - Rename a preset inline
 *   - Delete a preset with confirmation
 */

import React, { useState, useRef, useEffect } from "react";
import { Preset } from "../../services/api";

interface Props {
  presets: Preset[];
  onSelect: (preset: Preset) => void;
  onDelete: (presetId: string) => void;
  onRename: (presetId: string, newName: string) => void;
  onClose: () => void;
}

export const PresetMenu: React.FC<Props> = React.memo(({ presets, onSelect, onDelete, onRename, onClose }) => {
  const menuRef = useRef<HTMLDivElement>(null);
  const [renamingId, setRenamingId] = useState<string | null>(null);
  const [renameValue, setRenameValue] = useState("");
  const [confirmDeleteId, setConfirmDeleteId] = useState<string | null>(null);
  const renameInputRef = useRef<HTMLInputElement>(null);

  // Close menu when clicking outside
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (menuRef.current && !menuRef.current.contains(e.target as Node)) {
        onClose();
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, [onClose]);

  // Focus rename input when entering rename mode
  useEffect(() => {
    if (renamingId) {
      requestAnimationFrame(() => renameInputRef.current?.focus());
    }
  }, [renamingId]);

  const startRename = (preset: Preset, e: React.MouseEvent) => {
    e.stopPropagation();
    setRenamingId(preset.id);
    setRenameValue(preset.name);
  };

  const commitRename = () => {
    if (renamingId && renameValue.trim()) {
      onRename(renamingId, renameValue.trim());
    }
    setRenamingId(null);
    setRenameValue("");
  };

  const handleDelete = (presetId: string, e: React.MouseEvent) => {
    e.stopPropagation();
    if (confirmDeleteId === presetId) {
      // Second click = confirmed
      onDelete(presetId);
      setConfirmDeleteId(null);
    } else {
      // First click = ask for confirmation
      setConfirmDeleteId(presetId);
    }
  };

  return (
    <div
      ref={menuRef}
      dir="auto"
      style={{
        position: "absolute",
        bottom: "100%",
        left: 0,
        marginBottom: 6,
        width: 280,
        maxHeight: 300,
        backgroundColor: "#ffffff",
        borderRadius: 10,
        border: "1px solid #e0e0e0",
        boxShadow: "0 4px 16px rgba(0,0,0,0.12)",
        overflow: "hidden",
        zIndex: 100,
        animation: "fadeSlideIn 0.15s ease-out",
      }}
    >
      {/* Header */}
      <div style={{
        padding: "10px 14px",
        borderBottom: "1px solid #f0f0f0",
        display: "flex",
        justifyContent: "space-between",
        alignItems: "center",
      }}>
        <span style={{ fontSize: 13, fontWeight: 600, color: "#242424" }}>
          Saved Presets
        </span>
        <button
          onClick={onClose}
          style={{
            background: "none", border: "none", cursor: "pointer",
            fontSize: 16, color: "#888", padding: "0 2px", lineHeight: 1,
          }}
        >
          ×
        </button>
      </div>

      {/* Preset list */}
      <div style={{ overflowY: "auto", maxHeight: 240 }}>
        {presets.length === 0 && (
          <div style={{
            padding: "20px 14px",
            textAlign: "center",
            color: "#888",
            fontSize: 12,
          }}>
            No presets saved yet.
            <br />
            Use the 💾 button to save a plan as a preset.
          </div>
        )}

        {presets.map((preset) => (
          <div
            key={preset.id}
            onClick={() => { onSelect(preset); onClose(); }}
            style={{
              padding: "8px 14px",
              display: "flex",
              alignItems: "center",
              gap: 8,
              cursor: "pointer",
              borderBottom: "1px solid #f5f5f5",
              transition: "background-color 0.1s",
            }}
            onMouseEnter={(e) => { e.currentTarget.style.backgroundColor = "#f5f8ff"; }}
            onMouseLeave={(e) => { e.currentTarget.style.backgroundColor = "transparent"; }}
          >
            {/* Preset name or rename input */}
            <div style={{ flex: 1, minWidth: 0 }}>
              {renamingId === preset.id ? (
                <input
                  ref={renameInputRef}
                  value={renameValue}
                  onChange={(e) => setRenameValue(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") commitRename();
                    if (e.key === "Escape") { setRenamingId(null); }
                  }}
                  onBlur={commitRename}
                  onClick={(e) => e.stopPropagation()}
                  style={{
                    width: "100%",
                    fontSize: 12,
                    padding: "2px 6px",
                    border: "1px solid #5b5fc7",
                    borderRadius: 4,
                    outline: "none",
                  }}
                />
              ) : (
                <>
                  <div style={{
                    fontSize: 12,
                    fontWeight: 500,
                    color: "#242424",
                    whiteSpace: "nowrap",
                    overflow: "hidden",
                    textOverflow: "ellipsis",
                  }}>
                    {preset.name}
                  </div>
                  <div style={{
                    fontSize: 10,
                    color: "#999",
                    whiteSpace: "nowrap",
                    overflow: "hidden",
                    textOverflow: "ellipsis",
                  }}>
                    {preset.userMessage}
                  </div>
                </>
              )}
            </div>

            {/* Action buttons */}
            {renamingId !== preset.id && (
              <div style={{ display: "flex", gap: 4, flexShrink: 0 }}>
                {/* Rename */}
                <button
                  onClick={(e) => startRename(preset, e)}
                  title="Rename"
                  style={actionBtnStyle}
                >
                  ✎
                </button>
                {/* Delete — tap once to arm, tap again to confirm */}
                <button
                  onClick={(e) => handleDelete(preset.id, e)}
                  title={confirmDeleteId === preset.id ? "Click again to confirm" : "Delete"}
                  style={{
                    ...actionBtnStyle,
                    color: "#c50f1f",
                    backgroundColor: confirmDeleteId === preset.id ? "#fde8e8" : "transparent",
                    borderRadius: 4,
                    fontSize: confirmDeleteId === preset.id ? 10 : 13,
                    padding: confirmDeleteId === preset.id ? "2px 6px" : "2px 4px",
                  }}
                >
                  {confirmDeleteId === preset.id ? "Delete?" : "✕"}
                </button>
              </div>
            )}
          </div>
        ))}
      </div>
    </div>
  );
});
PresetMenu.displayName = "PresetMenu";

const actionBtnStyle: React.CSSProperties = {
  background: "none",
  border: "none",
  cursor: "pointer",
  fontSize: 13,
  color: "#666",
  padding: "2px 4px",
  borderRadius: 4,
  lineHeight: 1,
  transition: "color 0.1s",
};
