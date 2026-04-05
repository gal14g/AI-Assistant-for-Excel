/**
 * ExecutionTimeline – Shows step-by-step progress during plan execution.
 */

import React from "react";
import {
  Circle16Regular,
  CircleHalfFill16Regular,
  CheckmarkCircle16Filled,
  DismissCircle16Filled,
  SubtractCircle16Filled,
  Eye16Regular,
} from "@fluentui/react-icons";
import { ExecutionState, StepStatus } from "../../engine/types";

interface Props {
  state: ExecutionState;
  progressLog: { stepId: string; message: string; timestamp: string }[];
}

const STEP_ICONS: Record<StepStatus, React.ReactNode> = {
  pending: <Circle16Regular />,
  running: <CircleHalfFill16Regular />,
  success: <CheckmarkCircle16Filled />,
  error:   <DismissCircle16Filled />,
  skipped: <SubtractCircle16Filled />,
  preview: <Eye16Regular />,
};

export const ExecutionTimeline: React.FC<Props> = React.memo(({ state, progressLog }) => {
  return (
    <div dir="auto" className="cc-timeline">
      <div className="cc-timeline-head">
        <span className="cc-timeline-title">Execution Timeline</span>
        <span className={`cc-status-pill ${state.status}`}>{state.status}</span>
      </div>

      <div className="cc-timeline-steps">
        {state.stepResults.map((result) => (
          <div key={result.stepId} className="cc-timeline-step">
            <span className={`cc-timeline-step-icon ${result.status}`}>
              {STEP_ICONS[result.status] ?? STEP_ICONS.pending}
            </span>
            <div className="cc-timeline-step-body">
              <div className="cc-timeline-step-id">{result.stepId}</div>
              <div className="cc-timeline-step-msg">{result.message}</div>
              {result.durationMs !== undefined && (
                <div className="cc-timeline-step-duration">{result.durationMs}ms</div>
              )}
              {result.error && <div className="cc-timeline-step-error">{result.error}</div>}
            </div>
          </div>
        ))}
      </div>

      {progressLog.length > 0 && (
        <details className="cc-timeline-log">
          <summary>Detailed log ({progressLog.length} entries)</summary>
          <div className="cc-timeline-log-body">
            {progressLog.map((entry, i) => (
              <div key={i}>
                <span style={{ opacity: 0.6 }}>
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
});
ExecutionTimeline.displayName = "ExecutionTimeline";
