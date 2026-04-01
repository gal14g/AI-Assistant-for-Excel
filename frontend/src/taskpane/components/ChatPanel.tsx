/**
 * ChatPanel – Main chat interface. Microsoft Copilot for Excel style.
 */

import React, { useRef, useEffect, useCallback, useState } from "react";
import { MessageBubble } from "./MessageBubble";
import { ChatInput } from "./ChatInput";
import { PlanPreview } from "./PlanPreview";
import { PlanOptionsPanel } from "./PlanOptionsPanel";
import { ExecutionTimeline } from "./ExecutionTimeline";
import { SuggestedPrompts } from "./SuggestedPrompts";
import { useChat } from "../hooks/useChat";
import { useSelectionTracker } from "../hooks/useSelectionTracker";
import { usePlanExecution } from "../hooks/usePlanExecution";
import { sendFeedback, listPresets, savePreset, deletePreset, renamePreset, Preset } from "../../services/api";

// Copilot logo SVG
const CopilotLogo = () => (
  <svg width="20" height="20" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect width="20" height="20" rx="4" fill="url(#logo-grad)" />
    <defs>
      <linearGradient id="logo-grad" x1="0" y1="0" x2="20" y2="20" gradientUnits="userSpaceOnUse">
        <stop stopColor="#7719aa" />
        <stop offset="0.5" stopColor="#2764e7" />
        <stop offset="1" stopColor="#38b6ff" />
      </linearGradient>
    </defs>
    <path d="M10 3.5l1.4 4.2H15l-3.2 2.3 1.2 3.8L10 11.5l-3 2.3 1.2-3.8L5 7.7h3.6L10 3.5z" fill="white"/>
  </svg>
);

export const ChatPanel: React.FC = () => {
  const chat = useChat();
  const selection = useSelectionTracker();
  const execution = usePlanExecution();
  const messagesEndRef = useRef<HTMLDivElement>(null);

  // Track the last executed plan ID so undo works after execution completes
  const [lastExecutedPlanId, setLastExecutedPlanId] = useState<string | null>(null);
  const [presets, setPresets] = useState<Preset[]>([]);
  // When undo is triggered, prefill the input with the rolled-back user message.
  // The counter ensures repeated undos of the same text still trigger the effect.
  const [undoPrefill, setUndoPrefill] = useState<{ text: string; seq: number }>({ text: "", seq: 0 });
  // Save-preset naming mode
  const [savePresetMode, setSavePresetMode] = useState(false);
  const [savePresetName, setSavePresetName] = useState("");
  const [saveToast, setSaveToast] = useState("");
  // Track which option tab is active — reset when options change
  const optionsKey = chat.currentOptions?.map((o) => o.plan.planId).join(",") ?? "";
  const [activeOptionIndex, setActiveOptionIndex] = useState(0);
  const prevOptionsKey = useRef(optionsKey);
  if (prevOptionsKey.current !== optionsKey) {
    prevOptionsKey.current = optionsKey;
    if (activeOptionIndex !== 0) setActiveOptionIndex(0);
  }

  // The currently selected plan (from options or single plan)
  const activePlan = chat.currentOptions?.[activeOptionIndex]?.plan ?? chat.currentPlan;
  const hasOptions = (chat.currentOptions?.length ?? 0) > 0;

  // Only show suggested prompts if only the welcome system message is present
  const showSuggested = chat.messages.length <= 1 && !chat.isLoading;

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [chat.messages, chat.streamingText, execution.executionState]);

  const handleRun = async () => {
    if (activePlan) {
      setLastExecutedPlanId(activePlan.planId);
      await execution.runPlan(activePlan);
      // Record feedback (fire-and-forget)
      if (chat.interactionId) {
        sendFeedback(chat.interactionId, activePlan.planId, "applied");
      }
      chat.setCurrentPlan(null);
      chat.setCurrentOptions(null);
    }
  };

  const handlePreview = async () => {
    if (activePlan) await execution.previewPlan(activePlan);
  };

  const handleUndo = async () => {
    const planId = activePlan?.planId ?? lastExecutedPlanId;
    chat.setCurrentPlan(null);
    chat.setCurrentOptions(null);
    setLastExecutedPlanId(null);
    if (planId) {
      await execution.undoLast(planId);
    }
    // Remove the last user+assistant exchange and prefill the input with the rolled-back message
    const removedText = chat.removeLastExchange();
    if (removedText) {
      setUndoPrefill((prev) => ({ text: removedText, seq: prev.seq + 1 }));
    }
  };

  const handleCancel = () => {
    // Record dismiss feedback (fire-and-forget)
    if (chat.interactionId) {
      sendFeedback(chat.interactionId, null, "dismissed");
    }
    chat.setCurrentPlan(null);
    chat.setCurrentOptions(null);
    setLastExecutedPlanId(null);
    execution.reset();
  };

  const handleSend = useCallback(async (text: string, rangeTokens?: { address: string; sheetName: string }[]) => {
    await chat.sendMessage(text, rangeTokens);
  }, [chat]);

  // Load presets on mount
  useEffect(() => {
    listPresets().then(setPresets);
  }, []);

  const handleSavePresetClick = useCallback(() => {
    // Check there's a plan to save before entering naming mode
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

  const handleSuggestedPrompt = (prompt: string) => {
    handleSend(prompt, []);
  };

  return (
    <div dir="auto" style={{
      display: "flex", flexDirection: "column", height: "100vh",
      backgroundColor: "#fafafa",
      fontFamily: '"Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, sans-serif',
    }}>
      {/* Global animations */}
      <style>{`
        @keyframes bounce {
          0%, 80%, 100% { transform: translateY(0); }
          40% { transform: translateY(-4px); }
        }
        @keyframes fadeSlideIn {
          from { opacity: 0; transform: translateY(8px); }
          to   { opacity: 1; transform: translateY(0); }
        }
        @keyframes fadeIn {
          from { opacity: 0; }
          to   { opacity: 1; }
        }
        .chat-message-enter { animation: fadeSlideIn 0.25s ease-out; }
        .chat-scroll::-webkit-scrollbar { width: 4px; }
        .chat-scroll::-webkit-scrollbar-thumb { background: #d1d1d1; border-radius: 4px; }
        .chat-scroll::-webkit-scrollbar-thumb:hover { background: #a1a1a1; }
        .chat-scroll::-webkit-scrollbar-track { background: transparent; }
      `}</style>
      {/* Header */}
      <div style={{
        padding: "12px 16px",
        background: "linear-gradient(135deg, #f8f9ff 0%, #ffffff 100%)",
        borderBottom: "1px solid #e8e8e8",
        display: "flex", justifyContent: "space-between", alignItems: "center",
        flexShrink: 0,
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <CopilotLogo />
          <div>
            <div style={{ fontSize: 14, fontWeight: 600, color: "#242424", lineHeight: 1.2 }}>Copilot</div>
            <div style={{ fontSize: 10, color: "#616161" }}>Excel AI Assistant</div>
          </div>
        </div>
        <button
          onClick={chat.clearHistory}
          style={{
            background: "none", border: "1px solid #e8e8e8",
            borderRadius: 6, color: "#616161",
            padding: "4px 10px", fontSize: 12, cursor: "pointer",
          }}
        >
          New chat
        </button>
      </div>

      {/* Messages area */}
      <div className="chat-scroll" style={{ flex: 1, overflowY: "auto", padding: "16px 12px 0", scrollBehavior: "smooth" }}>
        {chat.messages.map((msg) => (
          <div key={msg.id} className="chat-message-enter">
            <MessageBubble message={msg} />
          </div>
        ))}

        {/* Streaming text */}
        {chat.streamingText && (
          <div style={{ display: "flex", gap: 8, alignItems: "flex-start", marginBottom: 16 }}>
            <div style={{
              padding: "12px 16px", backgroundColor: "#ffffff",
              borderRadius: "4px 18px 18px 18px", border: "1px solid #e8e8e8",
              fontSize: 13, color: "#616161", fontStyle: "italic",
              boxShadow: "0 1px 4px rgba(0,0,0,0.06)",
            }}>
              {chat.streamingText}
              <span style={{ opacity: 0.6 }}>▋</span>
            </div>
          </div>
        )}

        {/* Plan card — multi-option or single */}
        {hasOptions && chat.currentOptions && (
          <div style={{ marginBottom: 16 }}>
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
          <div style={{ marginBottom: 16 }}>
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

        {/* Execution timeline */}
        {execution.executionState && (
          <div style={{ marginBottom: 16 }}>
            <ExecutionTimeline
              state={execution.executionState}
              progressLog={execution.progressLog}
            />
          </div>
        )}

        {/* Error */}
        {(execution.lastError || chat.error) && (
          <div dir="auto" style={{
            padding: "10px 14px", backgroundColor: "#fdf3f3",
            borderRadius: 8, border: "1px solid #fcd6d6",
            color: "#c50f1f", fontSize: 12, marginBottom: 16,
          }}>
            {execution.lastError || chat.error}
          </div>
        )}

        <div ref={messagesEndRef} />
      </div>

      {/* Suggested prompts (shown only for fresh chat) */}
      {showSuggested && (
        <SuggestedPrompts onSelect={handleSuggestedPrompt} />
      )}

      {/* Thinking indicator (stop button is now in the input toolbar) */}
      {chat.isLoading && (
        <div style={{
          padding: "6px 16px", display: "flex", alignItems: "center", gap: 8,
          fontSize: 12, color: "#616161",
        }}>
          <div style={{ display: "flex", gap: 3 }}>
            {[0, 1, 2].map((i) => (
              <div key={i} style={{
                width: 6, height: 6, borderRadius: "50%",
                backgroundColor: "#5b5fc7",
                animation: `bounce 1.2s ease-in-out ${i * 0.2}s infinite`,
              }} />
            ))}
          </div>
          Copilot is thinking…
        </div>
      )}

      {/* Save preset naming bar */}
      {savePresetMode && (
        <div style={{
          display: "flex", alignItems: "center", gap: 6,
          padding: "8px 12px", borderTop: "1px solid #e8e8e8", backgroundColor: "#f8f9ff",
        }}>
          <input
            autoFocus
            value={savePresetName}
            onChange={(e) => setSavePresetName(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter") handleSavePresetConfirm();
              if (e.key === "Escape") setSavePresetMode(false);
            }}
            placeholder="Preset name..."
            style={{
              flex: 1, fontSize: 12, padding: "4px 8px",
              border: "1px solid #d1d1d1", borderRadius: 4, outline: "none",
            }}
          />
          <button
            onClick={handleSavePresetConfirm}
            disabled={!savePresetName.trim()}
            style={{
              fontSize: 11, padding: "4px 10px", borderRadius: 4, border: "none",
              backgroundColor: savePresetName.trim() ? "#0f6cbd" : "#c8c6c4",
              color: "#fff", fontWeight: 600, cursor: savePresetName.trim() ? "pointer" : "default",
            }}
          >
            Save
          </button>
          <button
            onClick={() => setSavePresetMode(false)}
            style={{
              fontSize: 11, padding: "4px 8px", borderRadius: 4,
              border: "1px solid #d1d1d1", backgroundColor: "#fff", cursor: "pointer",
            }}
          >
            Cancel
          </button>
        </div>
      )}

      {/* Save toast */}
      {saveToast && (
        <div style={{
          padding: "6px 12px", backgroundColor: "#dff6dd", color: "#107c10",
          fontSize: 12, fontWeight: 500, textAlign: "center",
          animation: "fadeIn 0.2s ease-out",
        }}>
          {saveToast}
        </div>
      )}

      {/* Chat input */}
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
