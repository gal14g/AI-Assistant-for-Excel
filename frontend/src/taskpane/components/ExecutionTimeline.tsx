/**
 * ExecutionTimeline – Shows step-by-step progress during plan execution.
 */

import React from "react";
import { ExecutionState, StepStatus } from "../../engine/types";

interface Props {
  state: ExecutionState;
  progressLog: { stepId: string; message: string; timestamp: string }[];
}

const STATUS_STYLES: Record<StepStatus, { color: string; icon: string }> = {
  pending: { color: "#888", icon: "○" },
  running: { color: "#217346", icon: "◎" },
  success: { color: "#4caf50", icon: "●" },
  error: { color: "#e53935", icon: "✕" },
  skipped: { color: "#ff9800", icon: "⊘" },
  preview: { color: "#7b1fa2", icon: "◇" },
};

export const ExecutionTimeline: React.FC<Props> = ({ state, progressLog }) => {
  return (
    <div
      style={{
        padding: "12px 14px",
        backgroundColor: "#fafafa",
        borderRadius: 8,
        border: "1px solid #e0e0e0",
        fontSize: 12,
      }}
    >
      <div
        style={{
          fontWeight: 600,
          marginBottom: 8,
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
        }}
      >
        <span>Execution Timeline</span>
        <span
          style={{
            fontSize: 11,
            padding: "2px 8px",
            borderRadius: 10,
            backgroundColor: getStatusBg(state.status),
            color: "#fff",
          }}
        >
          {state.status}
        </span>
      </div>

      {/* Step results */}
      <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
        {state.stepResults.map((result) => {
          const style = STATUS_STYLES[result.status] ?? STATUS_STYLES.pending;
          return (
            <div
              key={result.stepId}
              style={{
                display: "flex",
                alignItems: "flex-start",
                gap: 8,
              }}
            >
              <span style={{ color: style.color, fontSize: 14, lineHeight: "18px" }}>
                {style.icon}
              </span>
              <div style={{ flex: 1 }}>
                <div style={{ color: style.color, fontWeight: 500 }}>
                  {result.stepId}
                </div>
                <div style={{ color: "#666" }}>{result.message}</div>
                {result.durationMs !== undefined && (
                  <div style={{ color: "#aaa", fontSize: 11 }}>
                    {result.durationMs}ms
                  </div>
                )}
                {result.error && (
                  <div style={{ color: "#e53935", fontSize: 11, marginTop: 2 }}>
                    {result.error}
                  </div>
                )}
              </div>
            </div>
          );
        })}
      </div>

      {/* Progress log (collapsible) */}
      {progressLog.length > 0 && (
        <details style={{ marginTop: 8 }}>
          <summary style={{ cursor: "pointer", color: "#666", fontSize: 11 }}>
            Detailed log ({progressLog.length} entries)
          </summary>
          <div
            style={{
              marginTop: 4,
              maxHeight: 150,
              overflowY: "auto",
              fontFamily: "monospace",
              fontSize: 11,
              color: "#555",
              lineHeight: 1.6,
            }}
          >
            {progressLog.map((entry, i) => (
              <div key={i}>
                <span style={{ color: "#999" }}>
                  {new Date(entry.timestamp).toLocaleTimeString()}
                </span>{" "}
                [{entry.stepId}] {entry.message}
              </div>
            ))}
          </div>
        </details>
      )}
    </div>
  );
};

function getStatusBg(status: string): string {
  switch (status) {
    case "completed": return "#4caf50";
    case "failed": return "#e53935";
    case "running":
    case "previewing": return "#217346";
    case "rolledBack": return "#ff9800";
    default: return "#888";
  }
}
