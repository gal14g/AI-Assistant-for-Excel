"""
Feedback router – records the user's plan choice.

POST /api/feedback
"""

from __future__ import annotations

from typing import Literal, Optional

from fastapi import APIRouter, Request
from pydantic import BaseModel, Field
from slowapi import Limiter
from slowapi.util import get_remote_address

from ..db import log_choice, get_interaction

limiter = Limiter(key_func=get_remote_address)

router = APIRouter(prefix="/api", tags=["feedback"])


class FeedbackRequest(BaseModel):
    interactionId: str = Field(..., max_length=100)
    chosenPlanId: Optional[str] = Field(None, max_length=100)
    action: Literal["applied", "dismissed"]


@router.post("/feedback")
@limiter.limit("30/minute")
async def record_feedback(request: Request, body: FeedbackRequest):
    await log_choice(
        interaction_id=body.interactionId,
        chosen_plan_id=body.chosenPlanId,
        action=body.action,
    )

    # Promote applied interactions into the few-shot example pool
    if body.action == "applied":
        try:
            interaction = await get_interaction(body.interactionId)
            if interaction and interaction.get("plans_json"):
                from ..services.example_store import add_user_example

                await add_user_example(
                    interaction_id=body.interactionId,
                    user_message=interaction["user_message"],
                    assistant_response=interaction["plans_json"],
                )
        except Exception:
            pass  # promotion failure should never break the feedback endpoint

    return {"status": "ok"}
