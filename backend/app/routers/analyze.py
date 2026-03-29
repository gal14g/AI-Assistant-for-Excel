"""
POST /api/analyze — Analytical pipeline endpoint.

Accepts sheet data sent from the frontend (read via Office.js), runs the
AnalyticalPlanner → Orchestrator → ExplanationService pipeline, and returns
structured results plus a human-readable explanation.
"""
from __future__ import annotations

import logging
from typing import Any, Optional

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel, Field

from ..models.analytical_plan import SheetData, IntentType
from ..models.request import ConversationMessage
from ..planner import AnalyticalPlanner
from ..orchestrator import Orchestrator
from ..services.explanation_service import ExplanationService
from ..services.clarification_service import ClarificationPolicy
from ..services.matching_service import ColumnSufficiencyEvaluator

router = APIRouter(prefix="/api", tags=["analyze"])
logger = logging.getLogger(__name__)

# Module-level singletons (stateless services)
_planner = AnalyticalPlanner()
_explanation_svc = ExplanationService()
_clarification_policy = ClarificationPolicy()
_sufficiency_evaluator = ColumnSufficiencyEvaluator()


# ── Request / Response models ─────────────────────────────────────────────────


class AnalyzeRequest(BaseModel):
    userMessage: str
    sheets: dict[str, SheetData]
    activeSheet: Optional[str] = None
    conversationHistory: Optional[list[ConversationMessage]] = None


class AnalyzeResponse(BaseModel):
    intent: str
    strategy: Optional[str] = None
    message: str
    results: Optional[dict[str, Any]] = None
    needs_clarification: bool = False
    clarification_question: Optional[str] = None
    warnings: list[str] = Field(default_factory=list)
    execution_log: list[str] = Field(default_factory=list)
    confidence: float = 0.0


# ── Endpoints ─────────────────────────────────────────────────────────────────


@router.post("/analyze", response_model=AnalyzeResponse)
async def analyze(request: AnalyzeRequest) -> AnalyzeResponse:
    """
    Run the full analytical pipeline on provided sheet data.

    Flow:
    1. AnalyticalPlanner (LLM) → AnalyticalPlan JSON
    2. Orchestrator validates + executes tool chain
    3. ExplanationService generates the human-readable response
    """
    if not request.sheets:
        raise HTTPException(
            status_code=422,
            detail="At least one sheet must be provided in 'sheets'.",
        )

    # ── Step 1: Plan ──────────────────────────────────────────────────────────
    logger.info(
        "analyze: user_message=%r sheets=%s",
        request.userMessage[:120],
        list(request.sheets.keys()),
    )

    plan = await _planner.plan(
        user_message=request.userMessage,
        sheets=request.sheets,
        conversation_history=request.conversationHistory,
    )
    logger.info(
        "analyze: intent=%s confidence=%.2f needs_clarification=%s tool_chain=%s",
        plan.intent,
        plan.confidence,
        plan.needs_clarification,
        [op.value for op in plan.selected_tool_chain],
    )

    # ── Step 2: Column sufficiency gate ──────────────────────────────────────
    # If the planner included profile results in context we can evaluate them,
    # but at this point we only have the plan — use a lightweight heuristic.
    sufficiency: dict = {"sufficient": True, "score": 1.0, "weak_columns": [], "strong_columns": []}

    ask_clarification = _clarification_policy.should_ask(
        plan=plan,
        column_sufficiency=sufficiency,
        user_message=request.userMessage,
    )
    if ask_clarification:
        all_columns: list[str] = []
        for sheet in request.sheets.values():
            all_columns.extend(sheet.header_row)
        question = _clarification_policy.generate_question(plan, sufficiency, all_columns)
        return AnalyzeResponse(
            intent=plan.intent.value,
            message=question,
            needs_clarification=True,
            clarification_question=question,
            confidence=plan.confidence,
        )

    # ── Step 3: Orchestrate ───────────────────────────────────────────────────
    orchestrator = Orchestrator(request.sheets)
    result = orchestrator.execute(plan)

    if result.needs_clarification:
        return AnalyzeResponse(
            intent=plan.intent.value,
            message=result.clarification_question or "Could you provide more details?",
            needs_clarification=True,
            clarification_question=result.clarification_question,
            confidence=plan.confidence,
            execution_log=result.context.execution_log if result.context else [],
        )

    if not result.success:
        error_msg = "; ".join(result.errors) or "Pipeline failed."
        logger.warning("analyze pipeline failed: %s", error_msg)
        raise HTTPException(
            status_code=422,
            detail={
                "message": error_msg,
                "errors": result.errors,
                "execution_log": result.context.execution_log if result.context else [],
            },
        )

    # ── Step 4: Explain ───────────────────────────────────────────────────────
    explanation = _explanation_svc.generate(
        plan=plan,
        context=result.context,  # type: ignore[arg-type]
        final_data=result.final_data,
    )

    # Re-evaluate column sufficiency from profiling results (if available)
    if result.context:
        profile_data = result.context.get_data("profile_columns")
        if isinstance(profile_data, list):
            sufficiency = _sufficiency_evaluator.evaluate(profile_data)

    warnings: list[str] = []
    if result.context:
        warnings = result.context.all_warnings()
    if not sufficiency.get("sufficient", True):
        warnings.insert(0, sufficiency.get("recommendation", ""))

    return AnalyzeResponse(
        intent=plan.intent.value,
        strategy=plan.parameters.get("strategy"),
        message=explanation,
        results=result.final_data if result.final_data else None,
        needs_clarification=False,
        warnings=warnings[:10],
        execution_log=result.context.execution_log if result.context else [],
        confidence=plan.confidence,
    )


@router.get("/analyze/intents")
async def list_intents() -> dict:
    """Return all supported analytical intent types."""
    return {
        "intents": [intent.value for intent in IntentType],
        "description": "Supported analytical pipeline intents.",
    }
