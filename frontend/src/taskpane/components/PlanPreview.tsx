/**
 * PlanPreview – Action card shown before executing a plan.
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

const ACTION_LABELS: Record<string, string> = {
  readRange: "Read", writeValues: "Write values", writeFormula: "Write formula",
  matchRecords: "Match records", groupSum: "Group & sum", createTable: "Create table",
  applyFilter: "Apply filter", sortRange: "Sort", createPivot: "Create pivot",
  createChart: "Create chart", addConditionalFormat: "Conditional format",
  cleanupText: "Clean text", removeDuplicates: "Remove duplicates",
  freezePanes: "Freeze panes", findReplace: "Find & replace",
  addValidation: "Add validation", addSheet: "Add sheet",
  renameSheet: "Rename sheet", deleteSheet: "Delete sheet",
  copySheet: "Copy sheet", protectSheet: "Protect sheet",
  autoFitColumns: "Auto-fit", mergeCells: "Merge cells",
  setNumberFormat: "Number format",
};

export const PlanPreview: React.FC<Props> = React.memo(({
  plan, validation, isExecuting, isPreviewing,
  onPreview, onRun, onCancel, onUndo, canUndo,
}) => {
  const hasErrors = validation && !validation.valid;
  const hasWarnings = (validation?.warnings.length ?? 0) > 0 || (plan.warnings?.length ?? 0) > 0;

  return (
    <div style={{
      backgroundColor: "#ffffff",
      borderRadius: 12,
      border: "1px solid #e8e8e8",
      overflow: "hidden",
      boxShadow: "0 2px 8px rgba(0,0,0,0.08)",
    }}>
      {/* Card header */}
      <div style={{
        padding: "12px 16px",
        borderBottom: "1px solid #f0f0f0",
        background: "linear-gradient(135deg, #f0f4ff 0%, #f5f0ff 100%)",
        display: "flex", justifyContent: "space-between", alignItems: "center",
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <span style={{ fontSize: 16 }}>⚡</span>
          {/* Prefer localized translation when the planner emitted one.
             Canonical English `summary` remains the source of truth for logs/
             persistence; this only affects what the user sees. */}
          <span dir="auto" style={{ fontWeight: 600, fontSize: 14, color: "#242424" }}>
            {plan.summaryLocalized || plan.summary}
          </span>
        </div>
        <span style={{
          fontSize: 11, color: "#5b5fc7", fontWeight: 600,
          backgroundColor: "#ede7f6", padding: "2px 8px", borderRadius: 10,
        }}>
          {plan.steps.length} step{plan.steps.length !== 1 ? "s" : ""} · {Math.round(plan.confidence * 100)}% confidence
        </span>
      </div>

      {/* Steps */}
      <div style={{ padding: "10px 16px" }}>
        {plan.steps.map((step, i) => (
          <div key={step.id} style={{
            display: "flex", gap: 10, alignItems: "flex-start",
            padding: "6px 0",
            borderBottom: i < plan.steps.length - 1 ? "1px solid #f5f5f5" : "none",
          }}>
            <div style={{
              width: 22, height: 22, borderRadius: "50%",
              backgroundColor: "#f0f4ff", color: "#5b5fc7",
              display: "flex", alignItems: "center", justifyContent: "center",
              fontSize: 11, fontWeight: 700, flexShrink: 0,
            }}>
              {i + 1}
            </div>
            <div style={{ flex: 1 }}>
              <span dir="auto" style={{ fontSize: 13, color: "#242424" }}>
                {step.descriptionLocalized || step.description}
              </span>
              {/* marginInlineStart instead of marginLeft: follows writing direction
                 so the action badge sits after the description in both LTR and RTL. */}
              <span style={{
                marginInlineStart: 6, fontSize: 10, color: "#ffffff",
                backgroundColor: "#5b5fc7", padding: "1px 6px", borderRadius: 4,
                fontWeight: 500,
              }}>
                {ACTION_LABELS[step.action] ?? step.action}
              </span>
            </div>
          </div>
        ))}
      </div>

      {/* Validation / warnings */}
      {(hasErrors || hasWarnings) && (
        <div style={{ padding: "8px 16px", borderTop: "1px solid #f0f0f0" }}>
          {validation?.errors.map((e, i) => (
            <div key={i} style={{ display: "flex", gap: 6, alignItems: "flex-start", color: "#c50f1f", fontSize: 12, marginBottom: 3 }}>
              <span>✕</span><span dir="auto">{e.message}</span>
            </div>
          ))}
          {validation?.warnings.map((w, i) => (
            <div key={i} style={{ display: "flex", gap: 6, alignItems: "flex-start", color: "#c47f17", fontSize: 12, marginBottom: 3 }}>
              <span>⚠</span><span dir="auto">{w.message}</span>
            </div>
          ))}
          {plan.warnings?.map((w, i) => (
            <div key={i} style={{ display: "flex", gap: 6, alignItems: "flex-start", color: "#c47f17", fontSize: 12, marginBottom: 3 }}>
              <span>⚠</span><span dir="auto">{w}</span>
            </div>
          ))}
        </div>
      )}

      {/* Status line */}
      <div style={{ padding: "4px 16px 10px", fontSize: 11, color: "#616161", display: "flex", gap: 12 }}>
        {validation?.valid && <span style={{ color: "#107c10" }}>✓ Validation passed</span>}
        {plan.preserveFormatting && <span>Formatting will be preserved</span>}
      </div>

      {/* Action buttons */}
      <div style={{ padding: "10px 16px 14px", display: "flex", gap: 8, borderTop: "1px solid #f0f0f0" }}>
        <button onClick={onRun} disabled={isExecuting || isPreviewing} style={primaryBtn(isExecuting || isPreviewing)}>
          {isExecuting ? "Applying…" : "Apply"}
        </button>
        <button onClick={onPreview} disabled={isPreviewing || isExecuting} style={secondaryBtn(isPreviewing || isExecuting)}>
          {isPreviewing ? "Previewing…" : "Preview"}
        </button>
        {canUndo && (
          <button onClick={onUndo} style={secondaryBtn(false)}>Undo</button>
        )}
        {/* marginInlineStart: auto pushes Dismiss to the trailing edge
           (right in LTR, left in RTL). */}
        <button onClick={onCancel} disabled={isExecuting} style={{ ...secondaryBtn(isExecuting), marginInlineStart: "auto" }}>
          Dismiss
        </button>
      </div>
    </div>
  );
});
PlanPreview.displayName = "PlanPreview";

const primaryBtn = (disabled: boolean): React.CSSProperties => ({
  padding: "7px 18px", border: "none", borderRadius: 6,
  backgroundColor: disabled ? "#c8c6c4" : "#0f6cbd",
  color: "#fff", fontSize: 13, fontWeight: 600,
  cursor: disabled ? "default" : "pointer",
});

const secondaryBtn = (disabled: boolean): React.CSSProperties => ({
  padding: "7px 14px", border: "1px solid #d1d1d1", borderRadius: 6,
  backgroundColor: "#ffffff", color: disabled ? "#a0a0a0" : "#242424",
  fontSize: 13, fontWeight: 500, cursor: disabled ? "default" : "pointer",
});
