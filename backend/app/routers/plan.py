"""
Plan API router – handles plan generation, validation, and capability listing.
"""

from __future__ import annotations

from fastapi import APIRouter, HTTPException
from pydantic import ValidationError

from ..models.plan import ExecutionPlan
from ..models.request import PlanRequest, PlanResponse, ValidationResponse, CapabilityInfo
from ..services.planner import generate_plan, CAPABILITY_DESCRIPTIONS
from ..services.validator import validate_plan

router = APIRouter(prefix="/api", tags=["plan"])


@router.post("/plan", response_model=PlanResponse)
async def create_plan(request: PlanRequest) -> PlanResponse:
    """
    Generate an execution plan from a natural-language request.

    The LLM planner produces a typed JSON plan. The plan is validated
    before being returned. If validation fails, a 422 is returned.
    """
    try:
        plan, explanation = await generate_plan(request)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Plan generation failed: {str(e)}")

    # Validate the generated plan
    validation = validate_plan(plan)
    if not validation.valid:
        error_msgs = "; ".join(e.message for e in validation.errors)
        raise HTTPException(
            status_code=422,
            detail=f"Generated plan failed validation: {error_msgs}",
        )

    return PlanResponse(
        plan=plan,
        explanation=explanation,
        alternatives=None,
    )


@router.post("/validate", response_model=ValidationResponse)
async def validate_plan_endpoint(plan: ExecutionPlan) -> ValidationResponse:
    """
    Validate an execution plan without executing it.
    Useful for client-side plans or manual testing.
    """
    return validate_plan(plan)


@router.get("/capabilities", response_model=list[CapabilityInfo])
async def list_capabilities() -> list[CapabilityInfo]:
    """List all available capabilities that the planner can use."""
    # Mapping of which actions mutate vs affect formatting
    mutating = {
        "writeValues", "writeFormula", "matchRecords", "groupSum",
        "createTable", "sortRange", "createPivot", "createChart",
        "cleanupText", "removeDuplicates", "findReplace",
        "addSheet", "renameSheet", "deleteSheet", "copySheet",
        "mergeCells", "setNumberFormat", "autoFitColumns",
    }
    formatting = {
        "createTable", "createPivot", "createChart",
        "addConditionalFormat", "mergeCells", "setNumberFormat", "autoFitColumns",
    }

    return [
        CapabilityInfo(
            action=action,
            description=desc,
            mutates=action in mutating,
            affectsFormatting=action in formatting,
        )
        for action, desc in CAPABILITY_DESCRIPTIONS.items()
    ]
