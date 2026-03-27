"""
Chat service – conversational AI layer for Excel AI Copilot.

A single LLM call handles both routing and execution planning.
The LLM decides whether to:
  - Respond conversationally (questions, explanations, greetings)
  - Generate an Excel execution plan (operations that modify the workbook)

The response is always a JSON object with a "responseType" field.
"""

from __future__ import annotations

import json
import uuid
from datetime import datetime, timezone

import litellm

from ..config import settings
from ..models.chat import ChatRequest, ChatResponse
from ..models.plan import ExecutionPlan
from ..services.planner import _litellm_kwargs, CAPABILITY_DESCRIPTIONS, extract_json

# Silence verbose LiteLLM logs
litellm.success_callback = []


def _build_chat_system_prompt() -> str:
    caps = "\n".join(f"  - {k}: {v}" for k, v in CAPABILITY_DESCRIPTIONS.items())
    return f"""You are Excel AI Copilot, an intelligent assistant for Microsoft Excel.

You help users in two ways:
1. ANSWER QUESTIONS — explain Excel concepts, formulas, best practices, what you can do, etc.
2. EXECUTE EXCEL OPERATIONS — perform actions on the spreadsheet when the user asks you to do something

You MUST respond with a single valid JSON object in EXACTLY this format — no other text:
{{
  "responseType": "message" | "plan",
  "message": "<your reply to the user>",
  "plan": null
}}

Or when executing an operation:
{{
  "responseType": "plan",
  "message": "<plain-English explanation of what the plan will do>",
  "plan": {{
    "planId": "<uuid>",
    "createdAt": "<ISO timestamp>",
    "userRequest": "<original user message>",
    "summary": "<one-line summary>",
    "steps": [
      {{
        "id": "step_1",
        "description": "<what this step does>",
        "action": "<actionName>",
        "params": {{ }},
        "dependsOn": []
      }}
    ],
    "preserveFormatting": true,
    "confidence": 0.9,
    "warnings": []
  }}
}}

DECISION RULES:
- Use responseType "plan" only when the user wants to DO something to their spreadsheet (write data, apply formatting, create charts, sort, filter, etc.)
- Use responseType "message" for everything else: questions, greetings, explanations, "what can you do?", asking for advice, etc.
- For "message" type, set plan to null
- Always write a friendly, concise "message" in plain English

AVAILABLE EXCEL ACTIONS:
{caps}

PLAN RULES:
1. Never output executable code — only JSON plans
2. Prefer native Excel formulas (writeFormula) over computed values (writeValues)
3. Use exact range references from [[...]] tokens in the user message when provided
4. If no [[...]] tokens are given, use the sheet/column names the user describes (e.g. "Sheet1!A:A")
5. Set preserveFormatting: true unless the user explicitly asks to change formatting
6. Each step must have a unique id like step_1, step_2, etc.
7. Keep plans minimal — use only the steps actually needed

RANGE RULES (very important):
- Every "range" param must be a SINGLE range address — never comma-separated multi-ranges
  WRONG: "Sheet1!A:A,Sheet2!A:A"  RIGHT: "Sheet1!A:A"
- Full column references like "A:A" or "Sheet1!A:A" are allowed and preferred when the user says "column A"
- readRange reads ONE range per step — use multiple readRange steps if you need data from multiple ranges
- matchRecords, groupSum, writeValues, writeFormula ALL handle their own data reading internally — do NOT add a readRange step before them
- For matchRecords specifically: lookupRange and sourceRange are the two columns to join on; the action reads both internally

STEP MINIMALISM:
- matchRecords: needs 1 step only — it reads lookupRange and sourceRange itself
- groupSum: needs 1 step only — it reads dataRange itself
- writeFormula: needs 1 step — never precede it with readRange
- Only use readRange when you genuinely need to inspect data before deciding what to do next

EXAMPLES OF responseType "message":
- "What can you do?" → explain capabilities
- "What does VLOOKUP do?" → explain the function
- "Hi" / "Hello" → greet back
- "Should I use XLOOKUP or VLOOKUP?" → give advice

EXAMPLES OF responseType "plan":
- "Fill column B with the sum of A and C" → writeFormula, 1 step
- "Create a chart from [[Sheet1!A1:B20]]" → createChart, 1 step
- "Sort this table by the date column" → sortRange, 1 step
- "Match column A in Sheet1 to column A in Sheet2 and write column B" → matchRecords, 1 step

Respond ONLY with the JSON object. No preamble, no markdown fences."""


def _build_chat_messages(request: ChatRequest) -> list[dict]:
    messages: list[dict] = [{"role": "system", "content": _build_chat_system_prompt()}]

    if request.conversationHistory:
        for msg in request.conversationHistory[-8:]:
            messages.append({"role": msg.role, "content": msg.content})

    # Build user message with context
    parts = [request.userMessage]
    if request.rangeTokens:
        refs = ", ".join(f"[[{t.address}]]" for t in request.rangeTokens)
        parts.append(f"\nReferenced ranges: {refs}")
    if request.activeSheet:
        parts.append(f"\nActive sheet: {request.activeSheet}")
    if request.workbookName:
        parts.append(f"\nWorkbook: {request.workbookName}")

    messages.append({"role": "user", "content": "\n".join(parts)})
    return messages


async def chat(request: ChatRequest) -> ChatResponse:
    """
    Send a user message to the chat AI.
    Returns either a conversational reply or an execution plan.
    """
    response = await litellm.acompletion(
        messages=_build_chat_messages(request),
        **_litellm_kwargs(),
    )

    text: str = response.choices[0].message.content or ""
    parsed = extract_json(text)

    response_type = parsed.get("responseType", "message")
    message = parsed.get("message", "")

    if response_type == "plan" and parsed.get("plan"):
        plan_data: dict = parsed["plan"]
        # Ensure required fields
        if "planId" not in plan_data:
            plan_data["planId"] = str(uuid.uuid4())
        if "createdAt" not in plan_data:
            plan_data["createdAt"] = datetime.now(timezone.utc).isoformat()
        if "userRequest" not in plan_data:
            plan_data["userRequest"] = request.userMessage

        plan = ExecutionPlan(**plan_data)
        return ChatResponse(responseType="plan", message=message, plan=plan)

    return ChatResponse(responseType="message", message=message, plan=None)
