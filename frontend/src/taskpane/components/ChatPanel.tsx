/**
 * ChatPanel – Main chat interface assembling messages, input, and execution.
 */

import React, { useRef, useEffect } from "react";
import { MessageBubble } from "./MessageBubble";
import { ChatInput } from "./ChatInput";
import { PlanPreview } from "./PlanPreview";
import { ExecutionTimeline } from "./ExecutionTimeline";
import { useChat } from "../hooks/useChat";
import { useSelectionTracker } from "../hooks/useSelectionTracker";
import { usePlanExecution } from "../hooks/usePlanExecution";

export const ChatPanel: React.FC = () => {
  const chat = useChat();
  const selection = useSelectionTracker();
  const execution = usePlanExecution();
  const messagesEndRef = useRef<HTMLDivElement>(null);

  // Auto-scroll to latest message
  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [chat.messages, chat.streamingText, execution.executionState]);

  const handlePreview = async () => {
    if (chat.currentPlan) {
      await execution.previewPlan(chat.currentPlan);
    }
  };

  const handleRun = async () => {
    if (chat.currentPlan) {
      const state = await execution.runPlan(chat.currentPlan);
      if (state?.status === "completed") {
        chat.setCurrentPlan(null);
      }
    }
  };

  const handleUndo = async () => {
    if (chat.currentPlan) {
      await execution.undoLast(chat.currentPlan.planId);
    }
  };

  const handleCancel = () => {
    chat.setCurrentPlan(null);
    execution.reset();
  };

  return (
    <div
      style={{
        display: "flex",
        flexDirection: "column",
        height: "100vh",
        backgroundColor: "#f5f6f7",
        fontFamily:
          '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
      }}
    >
      {/* Header */}
      <div
        style={{
          padding: "12px 16px",
          backgroundColor: "#217346",
          color: "#fff",
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          boxShadow: "0 2px 4px rgba(0,0,0,0.1)",
        }}
      >
        <div>
          <div style={{ fontSize: 16, fontWeight: 600 }}>Excel AI Copilot</div>
          <div style={{ fontSize: 11, opacity: 0.8 }}>
            Natural-language spreadsheet assistant
          </div>
        </div>
        <button
          onClick={chat.clearHistory}
          style={{
            background: "rgba(255,255,255,0.15)",
            border: "1px solid rgba(255,255,255,0.3)",
            borderRadius: 6,
            color: "#fff",
            padding: "4px 10px",
            fontSize: 12,
            cursor: "pointer",
          }}
        >
          Clear
        </button>
      </div>

      {/* Messages area */}
      <div
        style={{
          flex: 1,
          overflowY: "auto",
          padding: "12px 12px 0",
        }}
      >
        {chat.messages.map((msg) => (
          <MessageBubble key={msg.id} message={msg} />
        ))}

        {/* Streaming text */}
        {chat.streamingText && (
          <div
            style={{
              padding: "10px 14px",
              backgroundColor: "#fff",
              borderRadius: 12,
              border: "1px solid #e0e0e0",
              marginBottom: 12,
              fontSize: 13,
              color: "#555",
              fontStyle: "italic",
            }}
          >
            {chat.streamingText}
            <span style={{ animation: "blink 1s infinite" }}>|</span>
          </div>
        )}

        {/* Plan preview */}
        {chat.currentPlan && (
          <div style={{ marginBottom: 12 }}>
            <PlanPreview
              plan={chat.currentPlan}
              validation={execution.validationResult}
              isExecuting={execution.isExecuting}
              isPreviewing={execution.isPreviewing}
              onPreview={handlePreview}
              onRun={handleRun}
              onCancel={handleCancel}
              onUndo={handleUndo}
              canUndo={execution.executionState?.status === "completed"}
            />
          </div>
        )}

        {/* Execution timeline */}
        {execution.executionState && (
          <div style={{ marginBottom: 12 }}>
            <ExecutionTimeline
              state={execution.executionState}
              progressLog={execution.progressLog}
            />
          </div>
        )}

        {/* Error display */}
        {(execution.lastError || chat.error) && (
          <div
            style={{
              padding: "10px 14px",
              backgroundColor: "#fde7e7",
              borderRadius: 8,
              border: "1px solid #e53935",
              color: "#c62828",
              fontSize: 12,
              marginBottom: 12,
            }}
          >
            {execution.lastError || chat.error}
          </div>
        )}

        <div ref={messagesEndRef} />
      </div>

      {/* Loading indicator */}
      {chat.isLoading && (
        <div
          style={{
            padding: "8px 16px",
            textAlign: "center",
            fontSize: 12,
            color: "#217346",
          }}
        >
          Thinking...
        </div>
      )}

      {/* Chat input */}
      <ChatInput
        onSend={chat.sendMessage}
        disabled={chat.isLoading || execution.isExecuting}
        currentSelectionAddress={selection.currentSelectionAddress}
      />
    </div>
  );
};
