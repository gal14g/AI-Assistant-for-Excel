/**
 * HistoryDrawer – slide-out panel listing past conversations.
 *
 * Loads the server-persisted conversation list on open and lets the user
 * switch to an older chat, rename it, or delete it.
 */

import React, { useCallback, useEffect, useState } from "react";
import { Dismiss16Regular, Delete16Regular, Edit16Regular } from "@fluentui/react-icons";
import {
  listConversations,
  renameConversation,
  deleteConversation,
  ConversationSummary,
} from "../../services/api";

interface Props {
  open: boolean;
  activeConversationId: string | null;
  onClose: () => void;
  onSelect: (id: string) => void;
  /** Called after the active conversation is deleted, so the panel can clear. */
  onActiveDeleted: () => void;
}

export const HistoryDrawer: React.FC<Props> = ({
  open, activeConversationId, onClose, onSelect, onActiveDeleted,
}) => {
  const [items, setItems] = useState<ConversationSummary[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [renameId, setRenameId] = useState<string | null>(null);
  const [renameValue, setRenameValue] = useState("");

  const refresh = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      setItems(await listConversations());
    } catch (e) {
      setError(e instanceof Error ? e.message : "Failed to load history");
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    if (open) void refresh();
  }, [open, refresh]);

  const handleRenameSubmit = async (id: string) => {
    const title = renameValue.trim();
    if (!title) { setRenameId(null); return; }
    try {
      await renameConversation(id, title);
      setRenameId(null);
      await refresh();
    } catch (e) {
      setError(e instanceof Error ? e.message : "Rename failed");
    }
  };

  const handleDelete = async (id: string) => {
    try {
      await deleteConversation(id);
      if (id === activeConversationId) onActiveDeleted();
      await refresh();
    } catch (e) {
      setError(e instanceof Error ? e.message : "Delete failed");
    }
  };

  if (!open) return null;

  return (
    <>
      <div className="cc-drawer-backdrop" onClick={onClose} />
      <aside className="cc-drawer" role="dialog" aria-label="Chat history">
        <div className="cc-drawer-header">
          <div className="cc-drawer-title">History</div>
          <button className="cc-btn ghost icon sm" onClick={onClose} aria-label="Close history">
            <Dismiss16Regular />
          </button>
        </div>

        {loading && <div className="cc-drawer-empty">Loading…</div>}
        {error && <div className="cc-error" style={{ margin: 8 }}>{error}</div>}
        {!loading && !error && items.length === 0 && (
          <div className="cc-drawer-empty">No saved chats yet.</div>
        )}

        <ul className="cc-drawer-list">
          {items.map((c) => {
            const isActive = c.id === activeConversationId;
            const isRenaming = renameId === c.id;
            return (
              <li
                key={c.id}
                className={`cc-drawer-item${isActive ? " active" : ""}`}
              >
                {isRenaming ? (
                  <input
                    autoFocus
                    className="cc-drawer-rename-input"
                    value={renameValue}
                    onChange={(e) => setRenameValue(e.target.value)}
                    onBlur={() => handleRenameSubmit(c.id)}
                    onKeyDown={(e) => {
                      if (e.key === "Enter") void handleRenameSubmit(c.id);
                      if (e.key === "Escape") setRenameId(null);
                    }}
                  />
                ) : (
                  <button
                    className="cc-drawer-item-main"
                    onClick={() => { onSelect(c.id); onClose(); }}
                    title={c.title}
                  >
                    <div className="cc-drawer-item-title">{c.title}</div>
                    <div className="cc-drawer-item-meta">
                      {c.messageCount} msg · {new Date(c.updatedAt).toLocaleDateString()}
                    </div>
                  </button>
                )}
                <div className="cc-drawer-item-actions">
                  <button
                    className="cc-btn ghost icon sm"
                    title="Rename"
                    aria-label="Rename conversation"
                    onClick={() => { setRenameId(c.id); setRenameValue(c.title); }}
                  >
                    <Edit16Regular />
                  </button>
                  <button
                    className="cc-btn ghost icon sm"
                    title="Delete"
                    aria-label="Delete conversation"
                    onClick={() => void handleDelete(c.id)}
                  >
                    <Delete16Regular />
                  </button>
                </div>
              </li>
            );
          })}
        </ul>
      </aside>
    </>
  );
};

HistoryDrawer.displayName = "HistoryDrawer";
