/**
 * MessageBubble – Renders a single chat message in Microsoft Copilot style.
 */

import React from "react";
import { ChatMessage } from "../../engine/types";

interface Props {
  message: ChatMessage;
}

/** Very simple inline markdown: bold, italic, bullet lists, numbered lists */
function renderMarkdown(text: string): React.ReactNode[] {
  const lines = text.split("\n");
  const result: React.ReactNode[] = [];
  let listItems: React.ReactNode[] = [];
  let listType: "ul" | "ol" | null = null;

  const flushList = () => {
    if (listItems.length > 0) {
      if (listType === "ul") {
        result.push(<ul key={result.length} style={{ margin: "4px 0 4px 16px", padding: 0 }}>{listItems}</ul>);
      } else {
        result.push(<ol key={result.length} style={{ margin: "4px 0 4px 16px", padding: 0 }}>{listItems}</ol>);
      }
      listItems = [];
      listType = null;
    }
  };

  const inlineStyles = (s: string): React.ReactNode => {
    const parts = s.split(/(\*\*.*?\*\*|\*.*?\*)/g);
    return parts.map((p, i) => {
      if (p.startsWith("**") && p.endsWith("**")) return <strong key={i}>{p.slice(2, -2)}</strong>;
      if (p.startsWith("*") && p.endsWith("*")) return <em key={i}>{p.slice(1, -1)}</em>;
      return p;
    });
  };

  for (const line of lines) {
    const ulMatch = line.match(/^[-•] (.+)/);
    const olMatch = line.match(/^\d+\. (.+)/);

    if (ulMatch) {
      if (listType !== "ul") { flushList(); listType = "ul"; }
      listItems.push(<li key={listItems.length} style={{ marginBottom: 2 }}>{inlineStyles(ulMatch[1])}</li>);
    } else if (olMatch) {
      if (listType !== "ol") { flushList(); listType = "ol"; }
      listItems.push(<li key={listItems.length} style={{ marginBottom: 2 }}>{inlineStyles(olMatch[1])}</li>);
    } else {
      flushList();
      if (line.trim() === "") {
        result.push(<div key={result.length} style={{ height: 6 }} />);
      } else {
        result.push(<div key={result.length}>{inlineStyles(line)}</div>);
      }
    }
  }
  flushList();
  return result;
}

// Copilot sparkle icon SVG
const CopilotIcon = () => (
  <svg width="28" height="28" viewBox="0 0 28 28" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ flexShrink: 0 }}>
    <rect width="28" height="28" rx="6" fill="url(#copilot-grad)" />
    <defs>
      <linearGradient id="copilot-grad" x1="0" y1="0" x2="28" y2="28" gradientUnits="userSpaceOnUse">
        <stop stopColor="#7719aa" />
        <stop offset="0.5" stopColor="#2764e7" />
        <stop offset="1" stopColor="#38b6ff" />
      </linearGradient>
    </defs>
    <path d="M14 6l1.8 5.4H21l-4.5 3.3 1.7 5.3L14 17l-4.2 3 1.7-5.3L7 11.4h5.2L14 6z" fill="white" />
  </svg>
);

export const MessageBubble: React.FC<Props> = ({ message }) => {
  const isUser = message.role === "user";
  const isSystem = message.role === "system";

  if (isSystem) {
    return (
      <div style={{ textAlign: "center", padding: "6px 0 10px", color: "#616161", fontSize: 12 }}>
        {message.content}
      </div>
    );
  }

  return (
    <div style={{ display: "flex", justifyContent: isUser ? "flex-end" : "flex-start", marginBottom: 16, gap: 8, alignItems: "flex-start" }}>
      {!isUser && <CopilotIcon />}

      <div style={{
        maxWidth: "82%",
        padding: isUser ? "10px 14px" : "12px 16px",
        borderRadius: isUser ? "18px 18px 4px 18px" : "4px 18px 18px 18px",
        backgroundColor: isUser ? "#0f6cbd" : "#ffffff",
        color: isUser ? "#ffffff" : "#242424",
        border: isUser ? "none" : "1px solid #e8e8e8",
        fontSize: 13,
        lineHeight: 1.6,
        boxShadow: "0 1px 4px rgba(0,0,0,0.06)",
      }}>
        {!isUser && (
          <div style={{ fontSize: 11, fontWeight: 600, color: "#5b5fc7", marginBottom: 6, letterSpacing: 0.3 }}>
            Copilot
          </div>
        )}

        <div style={{ wordBreak: "break-word" }}>
          {renderMarkdown(message.content)}
        </div>

        {message.plan && (
          <div style={{
            marginTop: 10,
            padding: "8px 12px",
            backgroundColor: "#f0f4ff",
            borderRadius: 8,
            border: "1px solid #c5cae9",
            fontSize: 12,
          }}>
            <div style={{ fontWeight: 600, color: "#5b5fc7", marginBottom: 4 }}>
              Plan: {message.plan.steps.length} step{message.plan.steps.length !== 1 ? "s" : ""}
            </div>
            {message.plan.steps.map((step, i) => (
              <div key={step.id} style={{ color: "#424242", marginLeft: 4, marginBottom: 2 }}>
                {i + 1}. {step.description}
              </div>
            ))}
          </div>
        )}

        <div style={{ fontSize: 10, color: isUser ? "rgba(255,255,255,0.65)" : "#bdbdbd", marginTop: 6, textAlign: "right" }}>
          {new Date(message.timestamp).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })}
        </div>
      </div>

      {isUser && (
        <div style={{
          width: 28, height: 28, borderRadius: "50%", backgroundColor: "#0f6cbd",
          display: "flex", alignItems: "center", justifyContent: "center",
          color: "#fff", fontSize: 12, fontWeight: 700, flexShrink: 0,
        }}>
          You
        </div>
      )}
    </div>
  );
};
