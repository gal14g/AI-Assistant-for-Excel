"""
Streaming API router – SSE endpoints for real-time plan generation updates.
"""

from __future__ import annotations

import json
from fastapi import APIRouter
from sse_starlette.sse import EventSourceResponse

from ..models.request import PlanRequest
from ..services.planner import generate_plan_stream

router = APIRouter(prefix="/api", tags=["stream"])


@router.post("/plan/stream")
async def stream_plan_generation(request: PlanRequest):
    """
    Stream plan generation via Server-Sent Events.

    Events:
    - type=explanation: Partial explanation text as it's generated
    - type=plan_ready: Complete JSON plan
    - type=error: Error message
    - type=done: Stream complete
    """

    async def event_generator():
        async for event in generate_plan_stream(request):
            yield {
                "event": "message",
                "data": json.dumps(event),
            }

    return EventSourceResponse(event_generator())
