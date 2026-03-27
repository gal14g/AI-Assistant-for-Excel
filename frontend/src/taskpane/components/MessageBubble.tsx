/**
 * MessageBubble – Renders a single chat message.
 */

import React from "react";
import { ChatMessage } from "../../engine/types";

interface Props {
  message: ChatMessage;
}

export const MessageBubble: React.FC<Props> = ({ message }) => {
  const isUser = message.role === "user";
  const isSystem = message.role === "system";

  return (
    <div
      style={{
        display: "flex",
        justifyContent: isUser ? "flex-end" : "flex-start",
        marginBottom: 12,
      }}
    >
      <div
        style={{
          maxWidth: "85%",
          padding: "10px 14px",
          borderRadius: 12,
          backgroundColor: isUser
            ? "#217346"
            : isSystem
              ? "#f0f0f0"
              : "#ffffff",
          color: isUser ? "#ffffff" : "#333333",
          border: isUser ? "none" : "1px solid #e0e0e0",
          fontSize: 13,
          lineHeight: 1.5,
          boxShadow: "0 1px 2px rgba(0,0,0,0.08)",
        }}
      >
        {/* Role label for non-user messages */}
        {!isUser && (
          <div
            style={{
              fontSize: 11,
              fontWeight: 600,
              color: isSystem ? "#888" : "#217346",
              marginBottom: 4,
              textTransform: "uppercase",
              letterSpacing: 0.5,
            }}
          >
            {isSystem ? "System" : "Copilot"}
          </div>
        )}

        {/* Message content – [[range]] tokens are rendered inline as part of the text */}
        <div style={{ whiteSpace: "pre-wrap", wordBreak: "break-word" }}>
          {message.content}
        </div>

        {/* Plan summary */}
        {message.plan && (
          <div
            style={{
              marginTop: 8,
              padding: "8px 10px",
              backgroundColor: "#f8f9fa",
              borderRadius: 8,
              border: "1px solid #e0e0e0",
              fontSize: 12,
              color: "#333",
            }}
          >
            <div style={{ fontWeight: 600, marginBottom: 4 }}>
              Plan: {message.plan.steps.length} step(s)
            </div>
            {message.plan.steps.map((step, i) => (
              <div key={step.id} style={{ marginLeft: 8, color: "#555" }}>
                {i + 1}. {step.description}
              </div>
            ))}
            {message.plan.warnings && message.plan.warnings.length > 0 && (
              <div style={{ marginTop: 6, color: "#e65100", fontSize: 11 }}>
                Warnings: {message.plan.warnings.join("; ")}
              </div>
            )}
          </div>
        )}

        {/* Timestamp */}
        <div
          style={{
            fontSize: 10,
            color: isUser ? "rgba(255,255,255,0.7)" : "#aaa",
            marginTop: 4,
            textAlign: "right",
          }}
        >
          {new Date(message.timestamp).toLocaleTimeString()}
        </div>
      </div>
    </div>
  );
};
