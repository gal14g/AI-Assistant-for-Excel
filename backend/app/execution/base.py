"""
Abstract contracts for Python-side plan execution.

`PlanExecutor` is the interface that both the MCP/xlwings bridge and any
future alternate runner (e.g. a headless COM-free executor driving openpyxl
for round-tripping .xlsx files without Excel) must implement. Keeping this
abstract decouples the MCP tool surface from the fact that xlwings is our
current only concrete runtime.

Type-shape stays aligned with `frontend/src/engine/types.ts` so a plan
produced by the backend planner can be executed identically regardless of
target: the Office.js executor in the add-in or the xlwings executor in
MCP mode.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Any, Callable, Literal, Optional

from app.models.plan import ExecutionPlan, PlanStep


# ── Callbacks ───────────────────────────────────────────────────────────────
# Progress callback shape mirrors the frontend's `onProgress` in
# `ExecutionOptions` so MCP can stream `notifications/progress` events to the
# client as each step advances.
ProgressCallback = Callable[[str, str], None]
"""(step_id, human_readable_message) -> None"""


# ── Per-step results ────────────────────────────────────────────────────────

StepStatus = Literal["success", "error", "skipped", "preview"]


@dataclass
class StepResult:
    """
    Mirrors the frontend `StepResult` interface (see
    `frontend/src/engine/types.ts`). Returned by each capability handler and
    aggregated into `ExecutionState`.
    """

    step_id: str
    status: StepStatus
    message: str
    outputs: dict[str, Any] = field(default_factory=dict)
    error: Optional[str] = None
    duration_ms: Optional[int] = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "stepId": self.step_id,
            "status": self.status,
            "message": self.message,
            "outputs": self.outputs,
            "error": self.error,
            "durationMs": self.duration_ms,
        }


# ── Execution context ───────────────────────────────────────────────────────


@dataclass
class ExecutorContext:
    """
    Per-plan execution state. Passed into each capability handler so it can
    access the target workbook, resolve cross-step bindings, push snapshots
    for undo, and stream progress back to MCP.

    `workbook_handle` is intentionally `Any` here — concrete executors
    narrow it:
    - `XlwingsExecutor` sets it to `xlwings.Book`
    - A hypothetical openpyxl executor would set it to `openpyxl.Workbook`
    """

    plan_id: str
    workbook_handle: Any
    """The target workbook (concrete type depends on the runtime)."""

    # Step-output bindings for {{step_N.field}} resolution. Populated as each
    # step completes successfully. Mirrors the TS executor's `resultsMap`.
    step_results: dict[str, StepResult] = field(default_factory=dict)

    # Pre-mutation snapshots for undo. Each entry captures range values +
    # formulas + number_format before a mutating step runs. Keyed by plan_id
    # so the whole plan can be rolled back atomically.
    snapshot_stack: list[dict[str, Any]] = field(default_factory=list)

    # Execution flags (mirror frontend ExecutionOptions).
    dry_run: bool = False
    preserve_formatting: bool = True

    on_progress: Optional[ProgressCallback] = None

    def emit_progress(self, step_id: str, message: str) -> None:
        if self.on_progress is not None:
            try:
                self.on_progress(step_id, message)
            except Exception:  # noqa: BLE001 — progress callbacks must not break execution
                pass


# ── Executor ABC ────────────────────────────────────────────────────────────


class PlanExecutor(ABC):
    """
    Abstract interface for running an `ExecutionPlan` against a concrete
    backend (desktop Excel via xlwings, or anything else). Concrete
    implementations should:

    1. Resolve `target_workbook` into a handle native to the backend.
    2. Topologically sort `plan.steps` by `dependsOn` (reuse the same
       algorithm as the Office.js executor — order must match).
    3. For each step, capture a pre-mutation snapshot when the capability
       is mutating and `dry_run` is False.
    4. Resolve `{{step_N.outputField}}` bindings in params before calling
       the handler (use `resolve_bindings` below).
    5. Dispatch to the registered capability handler.
    6. Aggregate `StepResult`s into an `ExecutionState`-shaped dict.
    7. On any step failure: stop immediately, report via `emit_progress`,
       and leave `snapshot_stack` intact so the caller can undo.

    The contract is deliberately simple so MCP tools like `undo_last` can be
    implemented in the concrete class without this ABC knowing anything
    about xlwings.
    """

    @abstractmethod
    async def execute_plan(
        self,
        plan: ExecutionPlan,
        target_workbook: Optional[str] = None,
        dry_run: bool = False,
        on_progress: Optional[ProgressCallback] = None,
    ) -> dict[str, Any]:
        """
        Run `plan` and return an ExecutionState-shaped dict:
            {
              "planId": str,
              "status": "completed" | "failed",
              "stepResults": [StepResult.to_dict(), ...],
              "startedAt": ISO-8601,
              "completedAt": ISO-8601,
            }

        `target_workbook` is a book name or path; `None` means "active
        workbook". Concrete executors that support cross-workbook references
        (like `XlwingsExecutor`) will also resolve `[Book2.xlsx]Sheet!A1`
        range tokens inside step params — the runtime just needs to know
        which workbook is the *default* for otherwise-unqualified addresses.
        """
        raise NotImplementedError

    @abstractmethod
    async def undo_last(self, target_workbook: Optional[str] = None) -> dict[str, Any]:
        """
        Pop the most recent snapshot off the stack and restore. Returns:
            { "restored": int, "message": str }
        """
        raise NotImplementedError


# ── Binding resolution (shared helper) ──────────────────────────────────────

import json
import re

_BINDING_RE = re.compile(r"\{\{(step_\w+)\.(\w+)\}\}")


def resolve_bindings(
    step: PlanStep,
    step_results: dict[str, StepResult],
) -> PlanStep:
    """
    Replace every `{{step_N.field}}` token in `step.params` with the value
    from a previously-completed step's outputs. Mirrors the TS executor's
    binding resolution (see `frontend/src/engine/executor.ts:resolveBindings`).

    Raises `ValueError` with a clear message when a binding can't be
    resolved — the MCP tool surfaces this as an `isError` response so the
    chat client can explain the problem to the user.
    """
    params_json = json.dumps(step.params, default=str)
    errors: list[str] = []

    def _replace(m: re.Match[str]) -> str:
        full, sid, field = m.group(0), m.group(1), m.group(2)
        result = step_results.get(sid)
        if result is None:
            errors.append(
                f"{full}: step '{sid}' has not run (missing from plan or earlier failure)"
            )
            return full
        if not result.outputs or field not in result.outputs:
            available = ", ".join(result.outputs.keys()) if result.outputs else "none"
            errors.append(
                f"{full}: step '{sid}' did not produce '{field}' (available: {available})"
            )
            return full
        value = result.outputs[field]
        # JSON-encode values so the resulting blob is still parseable when
        # we reinflate it (strings get their quotes, numbers stay numbers).
        return json.dumps(value, default=str).strip('"') if isinstance(value, str) else json.dumps(value, default=str)

    resolved_json = _BINDING_RE.sub(_replace, params_json)

    if errors:
        raise ValueError(
            "Could not resolve step output bindings:\n  " + "\n  ".join(errors)
        )

    resolved_params = json.loads(resolved_json)
    # PlanStep is a Pydantic model — model_copy preserves validation.
    return step.model_copy(update={"params": resolved_params})
