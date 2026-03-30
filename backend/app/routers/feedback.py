"""
Feedback router – records the user's plan choice.

POST /api/feedback
"""

from __future__ import annotations

from typing import Literal, Optional

from fastapi import APIRouter
from pydantic import BaseModel

from ..db import log_choice, get_interaction

router = APIRouter(prefix="/api", tags=["feedback"])


class FeedbackRequest(BaseModel):
    interactionId: str
    chosenPlanId: Optional[str] = None   # null = dismissed all options
    action: Literal["applied", "dismissed"]


@router.post("/feedback")
async def record_feedback(request: FeedbackRequest):
    await log_choice(
        interaction_id=request.interactionId,
        chosen_plan_id=request.chosenPlanId,
        action=request.action,
    )

    # Promote applied interactions into the few-shot example pool
    if request.action == "applied":
        try:
            interaction = await get_interaction(request.interactionId)
            if interaction and interaction.get("plans_json"):
                from ..services.example_store import add_user_example

                await add_user_example(
                    interaction_id=request.interactionId,
                    user_message=interaction["user_message"],
                    assistant_response=interaction["plans_json"],
                )
        except Exception:
            pass  # promotion failure should never break the feedback endpoint

    return {"status": "ok"}
