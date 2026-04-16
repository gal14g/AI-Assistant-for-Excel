/**
 * MessageBubble – Renders a single chat message. The execution timeline for
 * a plan that was run from this message is embedded inline inside the bubble,
 * so it stays scoped to the conversation turn that produced it.
 */

import React from "react";
import { ChatMessage } from "../../engine/types";
import { ExecutionTimeline } from "./ExecutionTimeline";

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
      if (listType === "ul") result.push(<ul key={result.length}>{listItems}</ul>);
      else result.push(<ol key={result.length}>{listItems}</ol>);
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
      listItems.push(<li key={listItems.length}>{inlineStyles(ulMatch[1])}</li>);
    } else if (olMatch) {
      if (listType !== "ol") { flushList(); listType = "ol"; }
      listItems.push(<li key={listItems.length}>{inlineStyles(olMatch[1])}</li>);
    } else {
      flushList();
      if (line.trim() === "") result.push(<div key={result.length} style={{ height: 6 }} />);
      else result.push(<div key={result.length}>{inlineStyles(line)}</div>);
    }
  }
  flushList();
  return result;
}

export const MessageBubble: React.FC<Props> = React.memo(({ message }) => {
  const isUser = message.role === "user";
  const isSystem = message.role === "system";

  if (isSystem) {
    return <div dir="auto" className="cc-msg-system">{message.content}</div>;
  }

  return (
    <div dir="auto" className={`cc-msg-row ${isUser ? "user" : "assistant"}`}>
      <div className={`cc-msg-avatar ${isUser ? "user" : "assistant"}`} aria-hidden="true">
        {isUser ? "You" : (
          /* Sparkle / AI wand icon */
          <svg width="15" height="15" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M10 2v3M10 15v3M2 10h3M15 10h3" stroke="white" strokeWidth="1.8" strokeLinecap="round"/>
            <circle cx="10" cy="10" r="3" fill="white"/>
            <path d="M4.93 4.93l2.12 2.12M12.95 12.95l2.12 2.12M4.93 15.07l2.12-2.12M12.95 7.05l2.12-2.12" stroke="white" strokeWidth="1.4" strokeLinecap="round"/>
          </svg>
        )}
      </div>
      <div className="cc-msg-body">
        <div className={`cc-msg-bubble ${isUser ? "user" : "assistant"}`}>
          {renderMarkdown(message.content)}

          {message.plan && (
            <div className="cc-msg-plan-inline">
              <div className="cc-msg-plan-inline-title">
                Plan · {message.plan.steps.length} step{message.plan.steps.length !== 1 ? "s" : ""}
              </div>
              {message.plan.steps.map((step, i) => (
                <div key={step.id} className="cc-msg-plan-inline-step" dir="auto">
                  {i + 1}. {step.descriptionLocalized || step.description}
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Inline execution timeline — scoped to this message's plan run */}
        {message.execution && (
          <ExecutionTimeline
            state={message.execution}
            progressLog={message.progressLog ?? []}
          />
        )}

        <div className="cc-msg-time">
          {new Date(message.timestamp).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })}
        </div>
      </div>
    </div>
  );
});
MessageBubble.displayName = "MessageBubble";
