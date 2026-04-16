"""
Concrete `PlanExecutor` backed by xlwings — the desktop Excel bridge.

Mirrors the behavior of `frontend/src/engine/executor.ts`:
- Topological sort by `dependsOn`
- Pre-mutation snapshots (captured before each mutating step)
- `{{step_N.outputField}}` binding resolution
- Per-step StepResult aggregation
- Progress callback fires after every step (for MCP notifications/progress)

xlwings gives us one piece Office.js can't: access to *every* workbook
open in the same Excel app. Cross-workbook addresses in step params
(e.g. `[Budget.xlsx]Sheet1!A1`) resolve via the shared `xlwings.apps`
registry — see `range_utils.py`.
"""

from __future__ import annotations

import asyncio
import logging
import time
from datetime import datetime
from typing import Any, Optional

from app.execution.base import (
    ExecutorContext,
    PlanExecutor,
    ProgressCallback,
    StepResult,
    resolve_bindings,
)
from app.execution.capability_registry import registry
from app.execution.snapshot import (
    MAX_SNAPSHOT_STACK,
    capture_snapshot,
    restore_snapshot,
)
from app.models.plan import ExecutionPlan, PlanStep

log = logging.getLogger(__name__)


# Snapshot stacks are per-book, stored at module scope so `undo_last` works
# across separate `execute_plan` calls in the same MCP session.
_SNAPSHOT_STACKS: dict[str, list[dict[str, Any]]] = {}


def _stack_for(book_name: str) -> list[dict[str, Any]]:
    return _SNAPSHOT_STACKS.setdefault(book_name, [])


class XlwingsExecutor(PlanExecutor):
    """
    Execute plans against the currently-running Excel desktop instance.

    Instantiation is cheap — the xlwings app handle is resolved lazily on
    first use so importing this module on a headless server doesn't fail.
    """

    def __init__(self) -> None:
        self._app = None  # type: ignore[assignment]

    # ── Public PlanExecutor contract ────────────────────────────────────────

    async def execute_plan(
        self,
        plan: ExecutionPlan,
        target_workbook: Optional[str] = None,
        dry_run: bool = False,
        on_progress: Optional[ProgressCallback] = None,
    ) -> dict[str, Any]:
        started = datetime.utcnow().isoformat()
        start_time = time.monotonic()

        book = self._resolve_book(target_workbook)
        ctx = ExecutorContext(
            plan_id=plan.planId,
            workbook_handle=book,
            dry_run=dry_run,
            preserve_formatting=plan.preserveFormatting,
            on_progress=on_progress,
        )

        # Topological order — plan.steps are already produced in dependency
        # order by the validator, but we defend against out-of-order inputs.
        ordered_steps = _topological_sort(list(plan.steps))

        step_results: list[StepResult] = []
        status = "completed"

        for step in ordered_steps:
            # Binding resolution: substitute {{step_N.field}} before handing
            # params to the handler.
            try:
                resolved = resolve_bindings(step, ctx.step_results)
            except ValueError as exc:
                res = StepResult(step_id=step.id, status="error", message=str(exc), error=str(exc))
                step_results.append(res)
                ctx.step_results[step.id] = res
                ctx.emit_progress(step.id, f"Binding error: {exc}")
                status = "failed"
                break

            cap = registry.get(step.action)
            if cap is None:
                res = StepResult(
                    step_id=step.id,
                    status="error",
                    message=f"Unknown action {step.action!r} — no handler registered.",
                    error=f"unknown_action:{step.action}",
                )
                step_results.append(res)
                ctx.step_results[step.id] = res
                ctx.emit_progress(step.id, res.message)
                status = "failed"
                break

            # Snapshot-before-mutate.
            if cap.mutates and not dry_run:
                try:
                    snap = capture_snapshot(
                        book,
                        _addresses_from_params(resolved.params),
                        default_sheet=book.sheets.active.name,
                    )
                    stack = _stack_for(book.name)
                    stack.append(snap)
                    if len(stack) > MAX_SNAPSHOT_STACK:
                        stack.pop(0)
                except Exception as exc:  # noqa: BLE001 — snapshot is best-effort
                    log.warning("Snapshot failed for step %s: %s", step.id, exc)

            ctx.emit_progress(step.id, f"Running {step.action}…")
            step_start = time.monotonic()
            try:
                # Handlers are synchronous (xlwings is blocking COM), so we
                # offload to a worker thread so the event loop stays free
                # for MCP progress notifications.
                raw = await asyncio.to_thread(cap.handler, ctx, resolved.params)
            except Exception as exc:  # noqa: BLE001
                log.exception("Step %s (%s) crashed", step.id, step.action)
                raw = {"status": "error", "message": f"Exception: {exc}", "error": str(exc)}

            res = StepResult(
                step_id=step.id,
                status=raw.get("status", "error"),
                message=raw.get("message", ""),
                outputs=raw.get("outputs", {}) or {},
                error=raw.get("error"),
                duration_ms=int((time.monotonic() - step_start) * 1000),
            )
            step_results.append(res)
            ctx.step_results[step.id] = res
            ctx.emit_progress(step.id, res.message)
            if res.status == "error":
                status = "failed"
                break

        return {
            "planId": plan.planId,
            "status": status,
            "stepResults": [r.to_dict() for r in step_results],
            "startedAt": started,
            "completedAt": datetime.utcnow().isoformat(),
            "durationMs": int((time.monotonic() - start_time) * 1000),
        }

    async def undo_last(self, target_workbook: Optional[str] = None) -> dict[str, Any]:
        book = self._resolve_book(target_workbook)
        stack = _stack_for(book.name)
        if not stack:
            return {"restored": 0, "message": "Nothing to undo."}

        snap = stack.pop()
        restored = await asyncio.to_thread(restore_snapshot, book.app.books, snap)
        return {
            "restored": restored,
            "message": f"Restored {restored} range(s) from snapshot at {snap['timestamp']}.",
        }

    # ── Helpers ─────────────────────────────────────────────────────────────

    def _resolve_book(self, target: Optional[str]):
        """
        Return an `xlwings.Book` for `target` (book name / path / None for
        active workbook). Raises with a clear message if no Excel is running.
        """
        try:
            import xlwings as xw
        except ImportError as exc:  # pragma: no cover
            raise RuntimeError(
                "xlwings isn't installed. `pip install xlwings` on a machine "
                "with desktop Excel to use MCP mode."
            ) from exc

        if self._app is None:
            if not xw.apps:
                raise RuntimeError(
                    "No running Excel instance found. Open Excel (with at "
                    "least one workbook) and retry."
                )
            self._app = xw.apps.active

        if target is None:
            return self._app.books.active

        # Exact name match first, then stem-insensitive match.
        for b in self._app.books:
            if b.name == target or b.name.lower() == target.lower():
                return b
            if b.name.rsplit(".", 1)[0].lower() == target.rsplit(".", 1)[0].lower():
                return b
        raise ValueError(f"Workbook {target!r} is not open in Excel.")


# ── Plan helpers ────────────────────────────────────────────────────────────


def _topological_sort(steps: list[PlanStep]) -> list[PlanStep]:
    """
    Sort `steps` by `dependsOn`. Handles missing dependencies defensively:
    any `dependsOn` entry that doesn't exist in the plan is ignored (the
    validator would have rejected the plan upstream if this were an issue).
    """
    by_id = {s.id: s for s in steps}
    visited: set[str] = set()
    out: list[PlanStep] = []

    def _visit(step: PlanStep) -> None:
        if step.id in visited:
            return
        for dep in step.dependsOn or []:
            dep_step = by_id.get(dep)
            if dep_step is not None:
                _visit(dep_step)
        visited.add(step.id)
        out.append(step)

    for s in steps:
        _visit(s)
    return out


def _addresses_from_params(params: dict[str, Any]) -> list[str]:
    """
    Pluck every address-looking string out of a step's params for snapshotting.
    Mirrors `frontend/src/engine/executor.ts:extractRangesFromParams`.

    We look at the canonical keys used across our action schemas:
    range, cell, sourceRange, destinationRange, target — plus anything
    nested under params that happens to be a string containing '!' (a
    sheet-qualified reference). This is intentionally permissive because
    over-snapshotting is far cheaper than missing a mutation.
    """
    found: list[str] = []
    candidates = ("range", "cell", "sourceRange", "destinationRange", "target", "location")
    for key in candidates:
        v = params.get(key)
        if isinstance(v, str):
            found.append(v)
    # Also catch values under common nested dicts (e.g. conditional format rules).
    for v in params.values():
        if isinstance(v, str) and "!" in v and v not in found and len(v) < 200:
            found.append(v)
    return found
