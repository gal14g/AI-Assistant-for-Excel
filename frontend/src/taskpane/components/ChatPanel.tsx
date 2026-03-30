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
import { sendFeedback } from "../../services/api";

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

  const handleSuggestedPrompt = (prompt: string) => {
    handleSend(prompt, []);
  };

  return (
    <div style={{
      display: "flex", flexDirection: "column", height: "100vh",
      backgroundColor: "#f5f5f5",
      fontFamily: '"Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, sans-serif',
    }}>
      {/* Header */}
      <div style={{
        padding: "10px 16px",
        backgroundColor: "#ffffff",
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
      <div style={{ flex: 1, overflowY: "auto", padding: "16px 12px 0" }}>
        {chat.messages.map((msg) => (
          <MessageBubble key={msg.id} message={msg} />
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
          <div style={{
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

      {/* Thinking indicator with stop button */}
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
          <button
            onClick={chat.stopMessage}
            style={{
              marginLeft: "auto",
              padding: "2px 10px",
              border: "1px solid #d1d1d1",
              borderRadius: 4,
              backgroundColor: "#fff",
              color: "#c50f1f",
              fontSize: 11,
              fontWeight: 600,
              cursor: "pointer",
            }}
          >
            Stop
          </button>
        </div>
      )}

      {/* Chat input */}
      <ChatInput
        onSend={handleSend}
        disabled={chat.isLoading || execution.isExecuting}
        currentSelectionAddress={selection.currentSelectionAddress}
      />
    </div>
  );
};
