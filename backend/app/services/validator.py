"""
Server-side plan validator.

Provides a second layer of validation beyond the frontend validator.
Checks schema conformance, business rules, and safety constraints.
"""

from __future__ import annotations

from pydantic import ValidationError

from ..models.plan import (
    ExecutionPlan,
    PlanStep,
    StepAction,
    ACTION_PARAM_MODELS,
)
from ..models.request import ValidationIssue, ValidationResponse


def validate_plan(plan: ExecutionPlan) -> ValidationResponse:
    """Validate an execution plan. Returns errors and warnings."""
    errors: list[ValidationIssue] = []
    warnings: list[ValidationIssue] = []

    # Check step IDs are unique
    step_ids = set()
    for step in plan.steps:
        if step.id in step_ids:
            errors.append(
                ValidationIssue(
                    message=f"Duplicate step id: {step.id}",
                    code="DUPLICATE_STEP_ID",
                    stepId=step.id,
                )
            )
        step_ids.add(step.id)

    # Validate each step
    for step in plan.steps:
        _validate_step(step, step_ids, plan.preserveFormatting, errors, warnings)

    # Check for dependency cycles
    if _has_cycle(plan.steps):
        errors.append(
            ValidationIssue(
                message="Dependency cycle detected in plan steps",
                code="DEPENDENCY_CYCLE",
            )
        )

    # Check confidence range
    if plan.confidence < 0 or plan.confidence > 1:
        warnings.append(
            ValidationIssue(
                message=f"Confidence {plan.confidence} outside [0, 1]",
                code="INVALID_CONFIDENCE",
            )
        )

    return ValidationResponse(
        valid=len(errors) == 0,
        errors=errors,
        warnings=warnings,
    )


def _validate_step(
    step: PlanStep,
    all_ids: set[str],
    preserve_formatting: bool,
    errors: list[ValidationIssue],
    warnings: list[ValidationIssue],
) -> None:
    """Validate a single plan step."""

    # Check action is valid
    try:
        StepAction(step.action)
    except ValueError:
        errors.append(
            ValidationIssue(
                message=f"Unknown action: {step.action}",
                code="UNKNOWN_ACTION",
                stepId=step.id,
            )
        )
        return

    # Validate dependencies exist
    if step.dependsOn:
        for dep in step.dependsOn:
            if dep not in all_ids:
                errors.append(
                    ValidationIssue(
                        message=f"Dependency '{dep}' not found",
                        code="INVALID_DEPENDENCY",
                        stepId=step.id,
                    )
                )

    # Validate params against the action-specific model
    param_model = ACTION_PARAM_MODELS.get(StepAction(step.action))
    if param_model:
        try:
            param_model(**step.params)
        except ValidationError as e:
            for err in e.errors():
                errors.append(
                    ValidationIssue(
                        message=f"Param error: {err['msg']} at {'.'.join(str(loc) for loc in err['loc'])}",
                        code="INVALID_PARAMS",
                        stepId=step.id,
                        field=str(err["loc"][-1]) if err["loc"] else None,
                    )
                )

    # Formatting safety check
    formatting_actions = {
        StepAction.createTable,
        StepAction.createChart,
        StepAction.createPivot,
        StepAction.addConditionalFormat,
        StepAction.mergeCells,
        StepAction.setNumberFormat,
        StepAction.autoFitColumns,
    }
    if preserve_formatting and StepAction(step.action) in formatting_actions:
        warnings.append(
            ValidationIssue(
                message=f"Action '{step.action}' may affect formatting (preserveFormatting=true)",
                code="FORMAT_SAFETY_WARNING",
                stepId=step.id,
            )
        )

    # Safety: check for dangerously large ranges
    _check_range_safety(step, warnings)


def _check_range_safety(
    step: PlanStep,
    warnings: list[ValidationIssue],
) -> None:
    """Warn about potentially dangerous range sizes."""
    params = step.params
    for key in ["range", "outputRange", "dataRange", "sourceRange"]:
        if key in params:
            val = str(params[key])
            # Full column references like "A:A" or "A:Z" can be huge
            if ":" in val and not any(c.isdigit() for c in val.split("!")[-1]):
                warnings.append(
                    ValidationIssue(
                        message=f"Full column reference '{val}' may be very large",
                        code="LARGE_RANGE_WARNING",
                        stepId=step.id,
                        field=key,
                    )
                )


def _has_cycle(steps: list[PlanStep]) -> bool:
    """Check if the step dependency graph has a cycle."""
    adj: dict[str, list[str]] = {s.id: (s.dependsOn or []) for s in steps}
    visited: set[str] = set()
    in_stack: set[str] = set()

    def dfs(node: str) -> bool:
        if node in in_stack:
            return True
        if node in visited:
            return False
        visited.add(node)
        in_stack.add(node)
        for dep in adj.get(node, []):
            if dfs(dep):
                return True
        in_stack.discard(node)
        return False

    return any(dfs(s.id) for s in steps if s.id not in visited)
