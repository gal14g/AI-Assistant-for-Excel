/**
 * PlanOptionsPanel – Tab-based multi-option plan selector.
 *
 * Shows pill tabs (Option A, Option B, ...) at the top.
 * The selected tab renders a full PlanPreview below.
 * The user picks an option, then Apply / Preview / Dismiss as usual.
 */

import React from "react";
import { PlanOption } from "../../engine/types";
import { PlanPreview } from "./PlanPreview";
import { ValidationResult } from "../../engine/validator";

interface Props {
  options: PlanOption[];
  validation: ValidationResult | null;
  isExecuting: boolean;
  isPreviewing: boolean;
  onPreview: () => void;
  onRun: () => void;
  onCancel: () => void;
  onUndo: () => void;
  canUndo: boolean;
  /** Called when the user switches tabs — parent should update the active plan. */
  onSelectOption: (index: number) => void;
  activeIndex: number;
}

export const PlanOptionsPanel: React.FC<Props> = React.memo(({
  options,
  validation,
  isExecuting,
  isPreviewing,
  onPreview,
  onRun,
  onCancel,
  onUndo,
  canUndo,
  onSelectOption,
  activeIndex,
}) => {
  if (options.length === 0) return null;

  const activePlan = options[activeIndex]?.plan;
  if (!activePlan) return null;

  return (
    <div dir="auto">
      {/* Tab bar — only show when there are multiple options */}
      {options.length > 1 && (
        <div
          style={{
            display: "flex",
            gap: 6,
            marginBottom: 8,
            flexWrap: "wrap",
          }}
        >
          {options.map((opt, i) => (
            <button
              key={i}
              onClick={() => onSelectOption(i)}
              style={{
                padding: "5px 12px",
                borderRadius: 16,
                border: i === activeIndex ? "2px solid #5b5fc7" : "1px solid #d1d1d1",
                backgroundColor: i === activeIndex ? "#ede7f6" : "#ffffff",
                color: i === activeIndex ? "#5b5fc7" : "#424242",
                fontSize: 12,
                fontWeight: i === activeIndex ? 600 : 400,
                cursor: "pointer",
                transition: "all 0.15s ease",
              }}
            >
              <span dir="auto">{opt.optionLabelLocalized || opt.optionLabel}</span>
            </button>
          ))}
        </div>
      )}

      {/* Active plan preview */}
      <PlanPreview
        plan={activePlan}
        validation={validation}
        isExecuting={isExecuting}
        isPreviewing={isPreviewing}
        onPreview={onPreview}
        onRun={onRun}
        onCancel={onCancel}
        onUndo={onUndo}
        canUndo={canUndo}
      />
    </div>
  );
});
PlanOptionsPanel.displayName = "PlanOptionsPanel";
