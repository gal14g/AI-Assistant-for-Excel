"""
Chat API router – conversational AI with Excel planning capability.
"""

from __future__ import annotations

from fastapi import APIRouter, HTTPException

from ..models.chat import ChatRequest, ChatResponse
from ..services.chat_service import chat
from ..services.validator import validate_plan

router = APIRouter(prefix="/api", tags=["chat"])


@router.post("/chat", response_model=ChatResponse)
async def chat_endpoint(request: ChatRequest) -> ChatResponse:
    """
    Conversational AI endpoint.

    The LLM decides whether to:
    - Return a conversational reply (responseType="message")
    - Return an execution plan (responseType="plan")

    Plans are validated before being returned.
    """
    try:
        result = await chat(request)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Chat failed: {str(e)}")

    # Validate the plan if one was generated
    if result.responseType == "plan" and result.plan:
        validation = validate_plan(result.plan)
        if not validation.valid:
            error_msgs = "; ".join(e.message for e in validation.errors)
            raise HTTPException(
                status_code=422,
                detail=f"Generated plan failed validation: {error_msgs}",
            )

    return result
