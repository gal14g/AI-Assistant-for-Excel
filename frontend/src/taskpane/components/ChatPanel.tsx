/**
 * ChatPanel – Main chat interface.
 *
 * Execution timelines are scoped per-message: when the user runs a plan, the
 * resulting ExecutionState and progress log are attached to the assistant
 * message that produced that plan, so switching chats (or opening an older
 * conversation, in the future) naturally brings the timeline with it.
 */

import React, { useRef, useEffect, useCallback, useState } from "react";
import { Add16Regular, History16Regular } from "@fluentui/react-icons";
import { MessageBubble } from "./MessageBubble";
import { ChatInput } from "./ChatInput";
import { PlanPreview } from "./PlanPreview";
import { PlanOptionsPanel } from "./PlanOptionsPanel";
import { SuggestedPrompts } from "./SuggestedPrompts";
import { HistoryDrawer } from "./HistoryDrawer";
import { useChat } from "../hooks/useChat";
import { useSelectionTracker } from "../hooks/useSelectionTracker";
import { usePlanExecution } from "../hooks/usePlanExecution";
import {
  sendFeedback, listPresets, savePreset, deletePreset, renamePreset, Preset,
  patchMessageExecution,
} from "../../services/api";

export const ChatPanel: React.FC = () => {
  const chat = useChat();
  const selection = useSelectionTracker();
  const execution = usePlanExecution();
  const messagesEndRef = useRef<HTMLDivElement>(null);

  // Stable refs so effects don't need the whole `chat` object in their dep array
  const updateMessageRef = useRef(chat.updateMessage);
  updateMessageRef.current = chat.updateMessage;
  const conversationIdRef = useRef(chat.conversationId);
  conversationIdRef.current = chat.conversationId;

  // Track the last executed plan ID so undo works after execution completes
  const [lastExecutedPlanId, setLastExecutedPlanId] = useState<string | null>(null);
  // Track which assistant message owns the currently-running timeline
  const [activeTimelineMsgId, setActiveTimelineMsgId] = useState<string | null>(null);
  // Tracks the last status we already sent to the backend, to prevent duplicate PATCHes
  const patchedStatusRef = useRef<string | null>(null);
  const [presets, setPresets] = useState<Preset[]>([]);
  const [undoPrefill, setUndoPrefill] = useState<{ text: string; seq: number }>({ text: "", seq: 0 });
  const [savePresetMode, setSavePresetMode] = useState(false);
  const [savePresetName, setSavePresetName] = useState("");
  const [saveToast, setSaveToast] = useState("");
  const [historyOpen, setHistoryOpen] = useState(false);

  // Active option index — reset when options change
  const optionsKey = chat.currentOptions?.map((o) => o.plan.planId).join(",") ?? "";
  const [activeOptionIndex, setActiveOptionIndex] = useState(0);
  const prevOptionsKey = useRef(optionsKey);
  if (prevOptionsKey.current !== optionsKey) {
    prevOptionsKey.current = optionsKey;
    if (activeOptionIndex !== 0) setActiveOptionIndex(0);
  }

  const activePlan = chat.currentOptions?.[activeOptionIndex]?.plan ?? chat.currentPlan;
  const hasOptions = (chat.currentOptions?.length ?? 0) > 0;
  const showSuggested = chat.messages.length <= 1 && !chat.isLoading;

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [chat.messages]);

  // Sync live execution state onto the active message so the timeline renders
  // inline with the assistant bubble that proposed the plan.
  useEffect(() => {
    if (!activeTimelineMsgId || !execution.executionState) return;

    updateMessageRef.current(activeTimelineMsgId, {
      execution: execution.executionState,
      progressLog: execution.progressLog,
    });

    // Only PATCH the backend once per terminal state (completed/failed/rolledBack).
    const status = execution.executionState.status;
    const isTerminal = status === "completed" || status === "failed" || status === "rolledBack";
    const convId = conversationIdRef.current;
    if (isTerminal && convId && patchedStatusRef.current !== `${activeTimelineMsgId}:${status}`) {
      patchedStatusRef.current = `${activeTimelineMsgId}:${status}`;
      void patchMessageExecution(
        convId,
        activeTimelineMsgId,
        execution.executionState,
        execution.progressLog,
      );
    }
  }, [execution.executionState, execution.progressLog, activeTimelineMsgId]);

  const handleRun = async () => {
    if (!activePlan) return;
    // Bind the upcoming execution to the assistant message that carries this plan
    const msgId = chat.getLatestPlanMessageId();
    setActiveTimelineMsgId(msgId);
    patchedStatusRef.current = null;
    setLastExecutedPlanId(activePlan.planId);
    await execution.runPlan(activePlan);
    if (chat.interactionId) sendFeedback(chat.interactionId, activePlan.planId, "applied");
    chat.setCurrentPlan(null);
    chat.setCurrentOptions(null);
  };

  const handlePreview = async () => {
    if (!activePlan) return;
    const msgId = chat.getLatestPlanMessageId();
    setActiveTimelineMsgId(msgId);
    patchedStatusRef.current = null;
    await execution.previewPlan(activePlan);
  };

  const handleUndo = async () => {
    const planId = activePlan?.planId ?? lastExecutedPlanId;
    chat.setCurrentPlan(null);
    chat.setCurrentOptions(null);
    setLastExecutedPlanId(null);
    if (planId) await execution.undoLast(planId);
    const removedText = chat.removeLastExchange();
    if (removedText) setUndoPrefill((prev) => ({ text: removedText, seq: prev.seq + 1 }));
    setActiveTimelineMsgId(null);
    execution.reset();
  };

  const handleCancel = () => {
    if (chat.interactionId) sendFeedback(chat.interactionId, null, "dismissed");
    chat.setCurrentPlan(null);
    chat.setCurrentOptions(null);
    setLastExecutedPlanId(null);
    setActiveTimelineMsgId(null);
    execution.reset();
  };

  const handleSend = useCallback(async (text: string, rangeTokens?: { address: string; sheetName: string }[]) => {
    await chat.sendMessage(text, rangeTokens);
  }, [chat]);

  const handleNewChat = useCallback(() => {
    chat.clearHistory();
    setLastExecutedPlanId(null);
    setActiveTimelineMsgId(null);
    execution.reset();
  }, [chat, execution]);

  useEffect(() => { listPresets().then(setPresets); }, []);

  const handleSavePresetClick = useCallback(() => {
    const lastPlanMsg = [...chat.messages].reverse().find(m => m.role === "assistant" && m.plan);
    if (!lastPlanMsg) return;
    setSavePresetName("");
    setSavePresetMode(true);
  }, [chat.messages]);

  const handleSavePresetConfirm = useCallback(() => {
    const name = savePresetName.trim();
    if (!name) return;
    const lastPlanMsg = [...chat.messages].reverse().find(m => m.role === "assistant" && m.plan);
    if (!lastPlanMsg) return;
    const msgIndex = chat.messages.indexOf(lastPlanMsg);
    const userMsg = chat.messages.slice(0, msgIndex).reverse().find(m => m.role === "user");
    savePreset(
      name,
      userMsg?.content ?? "",
      JSON.stringify({ responseType: "plan", message: lastPlanMsg.content, plan: lastPlanMsg.plan })
    ).then(() => {
      listPresets().then(setPresets);
      setSavePresetMode(false);
      setSaveToast(`Preset "${name}" saved!`);
      setTimeout(() => setSaveToast(""), 2500);
    }).catch((err) => console.error("Failed to save preset:", err));
  }, [savePresetName, chat.messages]);

  const handleDeletePreset = useCallback((presetId: string) => {
    deletePreset(presetId).then(() => listPresets().then(setPresets));
  }, []);
  const handleRenamePreset = useCallback((presetId: string, newName: string) => {
    renamePreset(presetId, newName).then(() => listPresets().then(setPresets));
  }, []);
  const handleSuggestedPrompt = (prompt: string) => { handleSend(prompt, []); };

  return (
    <div dir="auto" className="cc-app">
      {/* Header */}
      <div className="cc-header">
        <div className="cc-header-brand">
          <div className="cc-header-logo" aria-hidden="true">
            {/* Spreadsheet grid icon */}
            <svg width="14" height="14" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
              <rect x="1" y="1" width="6" height="6" rx="1" fill="white" fillOpacity="0.9"/>
              <rect x="9" y="1" width="6" height="6" rx="1" fill="white" fillOpacity="0.6"/>
              <rect x="1" y="9" width="6" height="6" rx="1" fill="white" fillOpacity="0.6"/>
              <rect x="9" y="9" width="6" height="6" rx="1" fill="white" fillOpacity="0.9"/>
            </svg>
          </div>
          <div className="cc-header-title">Copilot</div>
        </div>
        <div className="cc-header-actions">
          <button
            className="cc-btn ghost icon sm"
            title="History"
            aria-label="Open chat history"
            onClick={() => setHistoryOpen(true)}
          >
            <History16Regular />
          </button>
          <button className="cc-btn ghost" onClick={handleNewChat} title="Start a new chat">
            <Add16Regular /> New chat
          </button>
        </div>
      </div>

      {/* Messages */}
      <div className="cc-messages">
        {chat.messages.map((msg) => (
          <MessageBubble key={msg.id} message={msg} />
        ))}

        {hasOptions && chat.currentOptions && (
          <div style={{ marginBottom: 14 }}>
            <PlanOptionsPanel
              options={chat.currentOptions}
              validation={execution.validationResult}
              isExecuting={execution.isExecuting}
              isPreviewing={execution.isPreviewing}
              onPreview={handlePreview}
              onRun={handleRun}
              onCancel={handleCancel}
              onUndo={handleUndo}
              canUndo={execution.executionState?.status === "completed" || lastExecutedPlanId !== null}
              onSelectOption={setActiveOptionIndex}
              activeIndex={activeOptionIndex}
            />
          </div>
        )}
        {!hasOptions && chat.currentPlan && (
          <div style={{ marginBottom: 14 }}>
            <PlanPreview
              plan={chat.currentPlan}
              validation={execution.validationResult}
              isExecuting={execution.isExecuting}
              isPreviewing={execution.isPreviewing}
              onPreview={handlePreview}
              onRun={handleRun}
              onCancel={handleCancel}
              onUndo={handleUndo}
              canUndo={execution.executionState?.status === "completed" || lastExecutedPlanId !== null}
            />
          </div>
        )}

        {(execution.lastError || chat.error) && (
          <div dir="auto" className="cc-error">{execution.lastError || chat.error}</div>
        )}

        <div ref={messagesEndRef} />
      </div>

      {showSuggested && <SuggestedPrompts onSelect={handleSuggestedPrompt} />}

      {chat.isLoading && (
        <div className="cc-thinking">
          <div className="cc-thinking-dots">
            <div className="cc-thinking-dot" />
            <div className="cc-thinking-dot" />
            <div className="cc-thinking-dot" />
          </div>
          Copilot is thinking…
        </div>
      )}

      {savePresetMode && (
        <>
          <div className="cc-modal-backdrop" onClick={() => setSavePresetMode(false)} />
          <div className="cc-modal" role="dialog" aria-label="Save preset">
            <div className="cc-modal-title">Save as preset</div>
            <input
              className="cc-modal-input"
              autoFocus
              value={savePresetName}
              onChange={(e) => setSavePresetName(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter") handleSavePresetConfirm();
                if (e.key === "Escape") setSavePresetMode(false);
              }}
              placeholder="Preset name…"
            />
            <div className="cc-modal-actions">
              <button className="cc-btn" onClick={() => setSavePresetMode(false)}>Cancel</button>
              <button className="cc-btn primary" onClick={handleSavePresetConfirm} disabled={!savePresetName.trim()}>
                Save
              </button>
            </div>
          </div>
        </>
      )}

      {saveToast && <div className="cc-toast">{saveToast}</div>}

      <HistoryDrawer
        open={historyOpen}
        activeConversationId={chat.conversationId}
        onClose={() => setHistoryOpen(false)}
        onSelect={(id) => {
          setLastExecutedPlanId(null);
          setActiveTimelineMsgId(null);
          execution.reset();
          void chat.loadConversation(id);
        }}
        onActiveDeleted={() => {
          setLastExecutedPlanId(null);
          setActiveTimelineMsgId(null);
          execution.reset();
          chat.clearHistory();
        }}
      />

      <ChatInput
        onSend={handleSend}
        onStop={chat.stopMessage}
        onUndo={lastExecutedPlanId ? handleUndo : undefined}
        onSavePreset={handleSavePresetClick}
        onDeletePreset={handleDeletePreset}
        onRenamePreset={handleRenamePreset}
        disabled={execution.isExecuting}
        isLoading={chat.isLoading}
        canUndo={!!lastExecutedPlanId}
        presets={presets}
        currentSelectionAddress={selection.currentSelectionAddress}
        prefillText={undoPrefill}
      />
    </div>
  );
};
