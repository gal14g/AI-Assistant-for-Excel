/**
 * PlanPreview – Shows a plan before execution with Preview/Run/Cancel buttons.
 */

import React from "react";
import { ExecutionPlan } from "../../engine/types";
import { ValidationResult } from "../../engine/validator";

interface Props {
  plan: ExecutionPlan;
  validation: ValidationResult | null;
  isExecuting: boolean;
  isPreviewing: boolean;
  onPreview: () => void;
  onRun: () => void;
  onCancel: () => void;
  onUndo: () => void;
  canUndo: boolean;
}

export const PlanPreview: React.FC<Props> = ({
  plan,
  validation,
  isExecuting,
  isPreviewing,
  onPreview,
  onRun,
  onCancel,
  onUndo,
  canUndo,
}) => {
  return (
    <div
      style={{
        padding: 14,
        backgroundColor: "#fff",
        borderRadius: 8,
        border: "1px solid #217346",
        boxShadow: "0 2px 8px rgba(33,115,70,0.12)",
      }}
    >
      {/* Header */}
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          marginBottom: 10,
        }}
      >
        <span style={{ fontWeight: 600, fontSize: 14, color: "#217346" }}>
          Execution Plan
        </span>
        <span
          style={{
            fontSize: 11,
            color: "#666",
            padding: "2px 6px",
            backgroundColor: "#f0f0f0",
            borderRadius: 4,
          }}
        >
          {plan.steps.length} step(s) | confidence: {Math.round(plan.confidence * 100)}%
        </span>
      </div>

      {/* Summary */}
      <div style={{ fontSize: 13, color: "#333", marginBottom: 10 }}>
        {plan.summary}
      </div>

      {/* Steps list */}
      <div
        style={{
          marginBottom: 12,
          padding: "8px 10px",
          backgroundColor: "#f8f9fa",
          borderRadius: 6,
          fontSize: 12,
        }}
      >
        {plan.steps.map((step, i) => (
          <div
            key={step.id}
            style={{
              display: "flex",
              gap: 8,
              padding: "4px 0",
              borderBottom:
                i < plan.steps.length - 1 ? "1px solid #eee" : "none",
            }}
          >
            <span style={{ color: "#217346", fontWeight: 600, minWidth: 20 }}>
              {i + 1}.
            </span>
            <div style={{ flex: 1 }}>
              <span style={{ color: "#333" }}>{step.description}</span>
              <span
                style={{
                  marginLeft: 6,
                  fontSize: 10,
                  color: "#888",
                  fontFamily: "monospace",
                }}
              >
                ({step.action})
              </span>
            </div>
          </div>
        ))}
      </div>

      {/* Validation messages */}
      {validation && (
        <div style={{ marginBottom: 10, fontSize: 12 }}>
          {validation.errors.map((e, i) => (
            <div key={i} style={{ color: "#e53935", marginBottom: 2 }}>
              Error: {e.message}
            </div>
          ))}
          {validation.warnings.map((w, i) => (
            <div key={i} style={{ color: "#ff9800", marginBottom: 2 }}>
              Warning: {w.message}
            </div>
          ))}
          {validation.valid && validation.errors.length === 0 && (
            <div style={{ color: "#4caf50" }}>Validation passed</div>
          )}
        </div>
      )}

      {/* Warnings from planner */}
      {plan.warnings && plan.warnings.length > 0 && (
        <div style={{ marginBottom: 10, fontSize: 12 }}>
          {plan.warnings.map((w, i) => (
            <div key={i} style={{ color: "#ff9800" }}>
              Note: {w}
            </div>
          ))}
        </div>
      )}

      {/* Formatting preservation indicator */}
      <div style={{ fontSize: 11, color: "#666", marginBottom: 12 }}>
        {plan.preserveFormatting
          ? "Formatting will be preserved"
          : "This plan may modify formatting"}
      </div>

      {/* Action buttons */}
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
        <button
          onClick={onPreview}
          disabled={isPreviewing || isExecuting}
          style={{
            ...buttonStyle,
            backgroundColor: "#7b1fa2",
            opacity: isPreviewing || isExecuting ? 0.5 : 1,
          }}
        >
          {isPreviewing ? "Previewing..." : "Preview"}
        </button>
        <button
          onClick={onRun}
          disabled={isExecuting || isPreviewing}
          style={{
            ...buttonStyle,
            backgroundColor: "#4caf50",
            opacity: isExecuting || isPreviewing ? 0.5 : 1,
          }}
        >
          {isExecuting ? "Running..." : "Run Plan"}
        </button>
        <button
          onClick={onCancel}
          disabled={isExecuting}
          style={{
            ...buttonStyle,
            backgroundColor: "#666",
          }}
        >
          Cancel
        </button>
        {canUndo && (
          <button
            onClick={onUndo}
            style={{
              ...buttonStyle,
              backgroundColor: "#ff9800",
            }}
          >
            Undo Last
          </button>
        )}
      </div>
    </div>
  );
};

const buttonStyle: React.CSSProperties = {
  padding: "8px 16px",
  border: "none",
  borderRadius: 6,
  color: "#fff",
  fontSize: 13,
  fontWeight: 500,
  cursor: "pointer",
  transition: "opacity 0.2s",
};
