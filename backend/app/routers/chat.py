"""
Chat API router – conversational AI with Excel planning capability.
"""

from __future__ import annotations

import logging

from fastapi import APIRouter, HTTPException, Request
from fastapi.responses import StreamingResponse
from slowapi import Limiter
from slowapi.util import get_remote_address

from ..models.chat import ChatRequest, ChatResponse
from ..services.chat_service import chat, chat_stream
from ..services.validator import validate_plan

logger = logging.getLogger(__name__)
limiter = Limiter(key_func=get_remote_address)

router = APIRouter(prefix="/api", tags=["chat"])


@router.post("/chat", response_model=ChatResponse)
@limiter.limit("15/minute")
async def chat_endpoint(request: Request, body: ChatRequest) -> ChatResponse:
    """
    Conversational AI endpoint.

    The LLM decides whether to:
    - Return a conversational reply (responseType="message")
    - Return an execution plan (responseType="plan")

    Plans are validated before being returned.
    """
    try:
        result = await chat(body)
    except Exception as e:
        logger.error("Chat failed: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail="An error occurred processing your request. Please try again.",
        )

    # Validate all plans before returning
    if result.responseType == "plans" and result.plans:
        valid_options = []
        all_errors = []
        for i, option in enumerate(result.plans):
            validation = validate_plan(option.plan)
            if validation.valid:
                valid_options.append(option)
            else:
                all_errors.extend(f"Option {chr(65+i)}: {e.message}" for e in validation.errors)
        if not valid_options:
            error_msgs = "; ".join(all_errors)
            raise HTTPException(
                status_code=422,
                detail=f"Generated plan options failed validation: {error_msgs}",
            )
        result.plans = valid_options
    elif result.responseType == "plan" and result.plan:
        validation = validate_plan(result.plan)
        if not validation.valid:
            error_msgs = "; ".join(e.message for e in validation.errors)
            raise HTTPException(
                status_code=422,
                detail=f"Generated plan failed validation: {error_msgs}",
            )

    return result


@router.post("/chat/stream")
@limiter.limit("15/minute")
async def chat_stream_endpoint(request: Request, body: ChatRequest) -> StreamingResponse:
    """
    Streaming chat endpoint — returns SSE.

    Yields:
      data: {"type": "chunk", "text": "..."}  — LLM token as it arrives
      data: {"type": "done",  "result": {...}} — final ChatResponse JSON
    """
    async def generate():
        async for sse_line in chat_stream(body):
            yield sse_line

    return StreamingResponse(
        generate(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
            "Connection": "keep-alive",
        },
    )
