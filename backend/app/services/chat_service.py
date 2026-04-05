"""
Chat service – conversational AI layer for Excel AI Copilot.

A single LLM call handles both routing and execution planning.
The LLM decides whether to:
  - Respond conversationally (questions, explanations, greetings)
  - Generate an Excel execution plan (operations that modify the workbook)

The response is always a JSON object with a "responseType" field.
"""

from __future__ import annotations

import functools
import logging
import re
import time
import uuid
from datetime import datetime, timezone

from ..config import settings
from ..models.chat import ChatRequest, ChatResponse, PlanOption
from ..models.plan import ExecutionPlan
from ..services.planner import CAPABILITY_DESCRIPTIONS, extract_json
from ..services.llm_client import acompletion

logger = logging.getLogger(__name__)


@functools.lru_cache(maxsize=64)
def _build_chat_system_prompt(relevant_actions: tuple[str, ...] | None = None) -> str:
    if relevant_actions:
        filtered = {k: v for k, v in CAPABILITY_DESCRIPTIONS.items() if k in relevant_actions}
    else:
        filtered = CAPABILITY_DESCRIPTIONS
    caps = "\n".join(f"  - {k}: {v}" for k, v in filtered.items())
    return f"""You are Excel AI Copilot, an intelligent assistant for Microsoft Excel.

You help users in two ways:
1. ANSWER QUESTIONS — explain Excel concepts, formulas, best practices, what you can do, etc.
2. EXECUTE EXCEL OPERATIONS — perform actions on the spreadsheet when the user asks you to do something

You MUST respond with a single valid JSON object in EXACTLY this format — no other text:

For questions / greetings / explanations:
{{
  "responseType": "message",
  "message": "<your reply to the user>",
  "plans": null
}}

For spreadsheet operations — provide 2-3 DIFFERENT approaches the user can choose from:
{{
  "responseType": "plans",
  "message": "<brief overview of the options>",
  "plans": [
    {{
      "optionLabel": "Option A: <short approach name>",
      "plan": {{
        "planId": "<uuid>",
        "createdAt": "<ISO timestamp>",
        "userRequest": "<original user message>",
        "summary": "<one-line summary of THIS approach>",
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
    }},
    {{
      "optionLabel": "Option B: <short approach name>",
      "plan": {{ ... }}
    }}
  ]
}}

MULTI-OPTION RULES:
- Each option MUST use a meaningfully DIFFERENT strategy (e.g. formula-based vs PivotTable vs computed values)
- Give each option a short optionLabel starting with "Option A:", "Option B:", etc.
- If there is genuinely only ONE reasonable approach, return just one option in the array — do NOT invent bad alternatives
- 2-3 options is ideal. Never more than 3.
- Each plan in the array must be a complete, valid plan with its own planId, steps, etc.

DECISION RULES:
- Use responseType "plans" when the user wants to DO something to their spreadsheet (write data, apply formatting, create charts, sort, filter, etc.)
- Use responseType "message" for everything else: questions, greetings, explanations, "what can you do?", asking for advice, etc.
- For "message" type, set plans to null
- LANGUAGE: Always reply in the SAME language the user writes in. If the user writes in Hebrew, reply in Hebrew. If in English, reply in English. Match the user's language naturally.
- Always write a friendly, concise "message"

AVAILABLE EXCEL ACTIONS:
{caps}

PLAN RULES:
1. Never output executable code — only JSON plans
2. Prefer native Excel formulas (writeFormula) over computed values (writeValues)
3. Use exact range references from [[...]] tokens in the user message when provided.
   CRITICAL: Extract only the address INSIDE the [[...]] markers — do NOT include [[ or ]] in the JSON.
   The "Referenced ranges" list at the bottom of each user message shows the clean sheet-qualified addresses.
   ALWAYS copy those exact strings into your params — sheet names may be in Hebrew or other languages.
   Example: user says "column [[Sheet1!A:A]]" → use "Sheet1!A:A" in params (NOT "[[Sheet1!A:A]]")
   Example: Referenced ranges has [[תוכנה!A:B]] → use "תוכנה!A:B" exactly — do NOT replace with "Sheet1"
4. If no [[...]] tokens are given, use the sheet/column names the user describes (e.g. "Sheet1!A:A")
5. Set preserveFormatting: true unless the user explicitly asks to change formatting
6. Each step must have a unique id like step_1, step_2, etc.

MULTI-STEP PLANS:
- Use as many steps as the task genuinely requires — do not artificially limit to one step
- Use "dependsOn": ["step_id"] when a step must run after another (e.g. step_2 depends on step_1)
- Steps without dependsOn run first; steps with dependsOn run after their dependencies complete
- Examples of when to use multiple steps:
    - Read data → clean it → write results → create chart (4 steps)
    - Add new sheet → write headers → create table → create pivot (4 steps)
    - Match records → sort result → add conditional formatting (3 steps)
    - Clean up column → remove duplicates → auto-fit columns (3 steps)
    - DASHBOARD: addSheet("Dashboard") → createPivot → createChart → addConditionalFormat → autoFitColumns

writeFormula RULES (critical):
- Params: cell (string — SINGLE cell like "A1" or "Sheet1!D2"), formula (string starting with =), fillDown (int, optional)
- The param is "cell" NOT "range" — writeFormula takes a single cell, not a range
- COMPLEX FORMULAS supported: nested functions, LAMBDA/LET, dynamic arrays, XLOOKUP, etc.
- Dynamic arrays (Excel 365): =UNIQUE(A:A), =FILTER(A:B, B:B>0), =SORT(A:A) — these spill automatically
- When the user asks for a "complex formula" or "dynamic formula", use writeFormula

FIXING SPILL / #REF / #VALUE ERRORS:
- When a user reports a #SPILL error: the formula's spill range is blocked by other cells. Fix by:
  1. Clear the blocking cells first (use clearRange on the spill target area)
  2. Then rewrite the formula
- When a user reports a #REF error: a referenced range/sheet was deleted. Fix by rewriting the formula with correct references.
- When a user says "fix the formula" or "there's an error in column X": first READ the range to see the current formulas and values, then determine the fix.
- ALWAYS ask what the user intended if the fix is ambiguous (respond with responseType "message")

PIVOT FIELD RULES:
- rows and values accept either header names ("Department") or range addresses ("Sheet2!A:A")
- ALWAYS include rows and values in createPivot when the user specifies them — never drop them
- Example: user says "rows = [[Sheet2!A:A]]" → use rows: ["Sheet2!A:A"] (the handler resolves it)

CONDITIONAL FORMAT RULES:
- To highlight a row based on a condition in another column (e.g. "highlight row if column D is blank"): use ruleType="formula" with formula="=$D2=\"\"" applied to the WHOLE ROW range (e.g. "Sheet1!A2:Z1000"). NEVER use ruleType="cellValue" for cross-column or whole-row rules.
- To highlight cells >100 in red: ruleType="cellValue", operator="greaterThan", values=[100], format={{"fillColor":"#ffcccc"}}
- To show a green→red gradient: ruleType="colorScale" (no extra params needed)
- formula examples: blank check: =$D2="" | not blank: =$D2<>"" | cross-col: =$B2>$C2 | text: =$A2="LATE"

VALIDATION RULES:
- For dropdown from a range: validationType="list", formula="=Sheet2!A:A" (do NOT use listValues for range sources)
- For dropdown from fixed values: validationType="list", listValues=["Apple","Banana","Cherry"]
- For date range: validationType="date", operator="between", min="DATE(2024,1,1)", max="DATE(2024,12,31)"
- For unique values only: validationType="custom", formula="=COUNTIF($A$1:$A$100,A1)<=1"

MATCH RULES (critical):
- To write a constant ("pass", "yes", "✓") to a column based on a match: use matchRecords with writeValue="pass" — NEVER use writeValues for this
- For single-column lookup: matchRecords with lookupRange="Sheet1!A:A", sourceRange="Sheet2!A:A", returnColumns=[2], outputRange="Sheet1!C:C"
- For MULTI-COLUMN composite match (matching 2+ columns together): matchRecords with lookupRange="Sheet1!C:D" (2-col range), sourceRange="Sheet2!A:B" (2-col range), outputRange="Sheet1!I:I", writeValue="pass"
- NEVER set values: ["pass"] in a writeValues step to simulate a match — use matchRecords with writeValue instead

writeValues FORMAT RULES (critical — validation will reject if wrong):
- "range" is REQUIRED — always include a target range like "A1:B31" or "Sheet1!A1:C10"
- "values" MUST be a 2D array (list of lists) — each inner list is ONE ROW
- The FIRST ROW of values must ALWAYS be column headers/titles (e.g. ["Date", "Hours"])
  CORRECT: "values": [["Date", "Hours"], ["01/04/2026", ""], ["02/04/2026", ""]]
  WRONG:   "values": [["01/04/2026", ""], ["02/04/2026", ""]]  ← missing headers
  WRONG:   "values": ["01/04/2026", "02/04/2026"]  ← flat list, will be rejected
  WRONG:   "values": "01/04/2026"  ← not an array at all
- For generating lists/tables (dates, sequences, templates): use writeValues with the full 2D array
  Example task: "create a list of dates for April with a column for hours"
  → writeValues with range: "A1:B31", values: [["Date","Hours"],["01/04/2026",""],["02/04/2026",""],...]

RANGE RULES (very important):
- Every "range" param must be a SINGLE range address — never comma-separated multi-ranges
  WRONG: "Sheet1!A:A,Sheet2!A:A"  RIGHT: "Sheet1!A:A"
- Full column references like "A:A" or "Sheet1!A:A" are valid and preferred when the user says "column A"
- readRange reads ONE range per step — use multiple readRange steps for multiple ranges
- matchRecords, groupSum, writeValues, writeFormula read their own data — do NOT precede them with readRange
- For matchRecords: lookupRange and sourceRange are the two columns to match; it reads both itself

OPTIONAL PARAMS — many params have smart defaults, so you can omit them:
- createPivot: only sourceRange is required; destinationRange, pivotName, rows, values are all optional (auto-detected from headers)
- createTable: only range is required; tableName is optional (auto-generated)
- sortRange: only range is required; sortFields defaults to first column ascending
- matchRecords: returnColumns defaults to column 1 of sourceRange
- autoFitColumns: range is optional — omit to auto-fit all used columns on the active sheet
- setNumberFormat: common formats: "#,##0.00" (number), "0%" (percent), "dd/mm/yyyy" (date), "$#,##0.00" (USD), "General" (reset)

EXAMPLES OF responseType "message":
- "What can you do?" → explain capabilities
- "What does VLOOKUP do?" → explain the function
- "Hi" / "Hello" → greet back
- "Should I use XLOOKUP or VLOOKUP?" → give advice

EXAMPLES OF responseType "plan":
- "Read column A, match it to Sheet2, write results, then create a chart"
    → step_1: readRange, step_2: matchRecords (dependsOn step_1), step_3: createChart (dependsOn step_2)
- "Create a pivot from [[Sheet2!A1:B6]]" → createPivot, 1 step (sourceRange only needed)
- "Sort this table then add green highlighting to values above 100"
    → step_1: sortRange, step_2: addConditionalFormat (dependsOn step_1)
- "Clean up column A then remove duplicates" → step_1: cleanupText, step_2: removeDuplicates (dependsOn step_1)

Respond ONLY with the JSON object. No preamble, no markdown fences."""


async def _dynamic_few_shot_examples(user_message: str) -> list[dict]:
    """
    Retrieve the most relevant few-shot examples for the current query
    from the vector-backed example store. Returns user/assistant message
    pairs in the format expected by the LLM.
    """
    from .example_store import search_examples

    examples = await search_examples(user_message)
    messages: list[dict] = []
    for ex in examples:
        messages.append({"role": "user", "content": ex["user_message"]})
        messages.append({"role": "assistant", "content": ex["assistant_response"]})
    return messages


def _strip_wb_qualifier(address: str) -> str:
    """
    Strip the workbook qualifier from an address.
    "[WorkbookName.xlsx]Sheet1!A:A" → "Sheet1!A:A"
    "Sheet1!A:A" → "Sheet1!A:A" (unchanged)
    Uses rsplit on ']' so it's safe even when the workbook name contains ']'.
    """
    if "]" in address:
        return address.rsplit("]", 1)[-1]
    return address


_TRIPLE_BRACKET_RE = re.compile(r'\[\[\[([^\]]*\][^\]]*)\]\]')


def _clean_user_message(message: str, token_map: dict[str, str]) -> str:
    """
    Replace workbook-qualified range tokens in the message body with clean ones.

    Converts "[[[WB.xlsx]Sheet!Range]]" → "[[Sheet!Range]]" so the LLM sees
    a consistent, unambiguous format throughout the entire prompt.

    token_map: {raw_inner_address: clean_sheet_qualified_address}
    """
    def replace(m: re.Match[str]) -> str:
        # m.group(1) = inner address like "[WB.xlsx]Sheet!Range"
        inner = m.group(1)
        clean = token_map.get(inner, _strip_wb_qualifier(inner))
        return f"[[{clean}]]"

    return _TRIPLE_BRACKET_RE.sub(replace, message)


def _build_user_content(request: ChatRequest) -> str:
    # Build the clean address map from rangeTokens
    token_map: dict[str, str] = {}
    if request.rangeTokens:
        for t in request.rangeTokens:
            token_map[t.address] = _strip_wb_qualifier(t.address)

    # Clean the message body: replace [[[WB.xlsx]Sheet!Range]] → [[Sheet!Range]]
    clean_message = _clean_user_message(request.userMessage, token_map)

    parts = [clean_message]

    # Inject current date/time so the LLM knows "next month", "today", etc.
    now = datetime.now(timezone.utc)
    parts.append(f"\nCurrent date: {now.strftime('%A, %d/%m/%Y')} (UTC)")

    # Tell the LLM which date format to use based on the user's locale
    locale = getattr(request, "locale", None) or ""
    # Most locales use dd/mm/yyyy; US/Canada/Philippines/etc use mm/dd/yyyy
    mm_dd_locales = {"en-US", "en-PH", "en-CA", "fr-CA", "ko-KR", "ja-JP", "zh-TW"}
    locale_prefix = locale.split("-")[0] if locale else ""
    if locale in mm_dd_locales or locale_prefix in {"ja", "ko", "zh"}:
        date_fmt = "mm/dd/yyyy"
    else:
        date_fmt = "dd/mm/yyyy"
    parts.append(f"User date format: {date_fmt} — ALWAYS use this format consistently for ALL dates in your response")

    # If the frontend reported the used range, tell the LLM where free space starts
    if getattr(request, "usedRangeEnd", None):
        parts.append(f"\nSheet used range ends at: {request.usedRangeEnd} — place new data below or beside it")

    if request.rangeTokens:
        # Strip workbook qualifiers before sending to the LLM.
        # The LLM must produce sheet-qualified addresses ("Sheet!A:A") in its JSON,
        # not workbook-qualified ones ("[WB.xlsx]Sheet!A:A").  Feeding workbook-
        # qualified tokens causes the LLM to hallucinate generic names like "Sheet1"
        # instead of preserving the real (possibly Hebrew) sheet name.
        clean = list(token_map.values())
        refs = ", ".join(f"[[{a}]]" for a in clean)
        parts.append(f"\nReferenced ranges: {refs}")
    if request.activeSheet:
        parts.append(f"\nActive sheet: {request.activeSheet}")
    if request.workbookName:
        parts.append(f"\nWorkbook: {request.workbookName}")
    return "\n".join(parts)


async def _build_chat_messages(request: ChatRequest, relevant_actions: list[str] | None = None) -> list[dict]:
    actions_key = tuple(relevant_actions) if relevant_actions else None
    messages: list[dict] = [{"role": "system", "content": _build_chat_system_prompt(actions_key)}]

    # Dynamic few-shot examples — retrieved via vector similarity to the user's query
    few_shot = await _dynamic_few_shot_examples(request.userMessage)
    messages.extend(few_shot)

    if request.conversationHistory:
        for msg in request.conversationHistory[-8:]:
            if msg.role in ("user", "assistant"):
                messages.append({"role": msg.role, "content": msg.content[:5000]})

    messages.append({"role": "user", "content": _build_user_content(request)})
    return messages


def _build_retry_messages(request: ChatRequest, relevant_actions: list[str] | None = None) -> list[dict]:
    """Stripped-down prompt for retry — no few-shot examples, harder JSON enforcement."""
    actions_key = tuple(relevant_actions) if relevant_actions else None
    system = (
        _build_chat_system_prompt(actions_key)
        + "\n\nCRITICAL: Your ENTIRE response must be ONE valid JSON object and nothing else. "
        "No prose, no markdown, no explanation — just the JSON object starting with { and ending with }."
    )
    return [
        {"role": "system", "content": system},
        {"role": "user", "content": _build_user_content(request)},
    ]


def _fill_plan_defaults(plan_data: dict, request: ChatRequest) -> ExecutionPlan:
    """Ensure required top-level plan fields exist, then parse."""
    if "planId" not in plan_data:
        plan_data["planId"] = str(uuid.uuid4())
    if "createdAt" not in plan_data:
        plan_data["createdAt"] = datetime.now(timezone.utc).isoformat()
    if "userRequest" not in plan_data:
        plan_data["userRequest"] = request.userMessage
    return ExecutionPlan(**plan_data)


def _parse_response(text: str, request: ChatRequest) -> ChatResponse:
    parsed = extract_json(text)
    response_type = parsed.get("responseType", "message")
    message = parsed.get("message", "")

    # New multi-option format: responseType "plans" with array
    if response_type == "plans" and parsed.get("plans"):
        options: list[PlanOption] = []
        for i, opt in enumerate(parsed["plans"]):
            plan_data = opt.get("plan", opt)  # handle both {optionLabel, plan} and bare plan
            plan = _fill_plan_defaults(plan_data, request)
            label = opt.get("optionLabel", f"Option {chr(65 + i)}")
            options.append(PlanOption(optionLabel=label, plan=plan))
        if options:
            return ChatResponse(
                responseType="plans",
                message=message or "Here are a few approaches:",
                plans=options,
            )

    # Backward compat: single plan (from few-shot examples or simpler LLM output)
    if response_type == "plan" and parsed.get("plan"):
        plan = _fill_plan_defaults(parsed["plan"], request)
        option = PlanOption(optionLabel="Option A", plan=plan)
        return ChatResponse(
            responseType="plans",
            message=message or plan.summary,
            plans=[option],
        )

    # Empty message = LLM returned valid JSON but with no content — treat as failure
    if not message.strip():
        raise ValueError("LLM returned empty message")

    return ChatResponse(responseType="message", message=message, plan=None)


async def _log_interaction_safe(
    interaction_id: str,
    request: ChatRequest,
    result: ChatResponse,
    latency_ms: int,
) -> None:
    """Log to the feedback DB — never let DB errors break the chat flow."""
    try:
        from ..db import log_interaction

        await log_interaction(
            interaction_id=interaction_id,
            user_message=request.userMessage,
            active_sheet=request.activeSheet,
            workbook_name=request.workbookName,
            range_tokens=request.rangeTokens,
            response=result,
            model_used=settings.llm_model,
            latency_ms=latency_ms,
        )
    except Exception:
        pass


async def chat(request: ChatRequest) -> ChatResponse:
    """
    Send a user message to the chat AI.
    Returns either a conversational reply or multiple plan options.
    On JSON parse failure, retries once with a stripped-down prompt.
    If both attempts fail, returns a friendly error message instead of crashing.
    """
    from .capability_store import search_capabilities

    relevant_actions = search_capabilities(request.userMessage)
    interaction_id = str(uuid.uuid4())
    start = time.monotonic()

    result: ChatResponse | None = None

    try:
        text = await acompletion(
            messages=await _build_chat_messages(request, relevant_actions),
        )
        logger.debug("LLM raw response (attempt 1): %s", text[:500])
        result = _parse_response(text, request)
    except Exception as exc:
        logger.warning("Chat attempt 1 failed: %s", exc)

    # Retry with no few-shot examples and a stronger JSON-only instruction
    if result is None:
        try:
            text = await acompletion(
                messages=_build_retry_messages(request, relevant_actions),
            )
            logger.debug("LLM raw response (attempt 2): %s", text[:500])
            result = _parse_response(text, request)
        except Exception as exc:
            logger.error("Chat attempt 2 failed: %s", exc)
            result = ChatResponse(
                responseType="message",
                message="Sorry, I couldn't process that request. Try rephrasing it more simply, or break it into smaller steps.",
                plan=None,
            )

    latency_ms = int((time.monotonic() - start) * 1000)
    result.interactionId = interaction_id
    await _log_interaction_safe(interaction_id, request, result, latency_ms)
    await _persist_conversation_turn(request, result)

    return result


async def _persist_conversation_turn(request: ChatRequest, result: ChatResponse) -> None:
    """
    Append the user turn + assistant reply to the conversations store.
    Creates a new conversation on the fly when conversationId is omitted.
    Failures never break the chat flow.
    """
    import uuid as _uuid

    try:
        from ..db import append_conv_message, create_conversation

        conv_id = request.conversationId
        if not conv_id:
            # Title from the first user message, trimmed.
            title = request.userMessage.strip().splitlines()[0][:60] or "New chat"
            conv_id = await create_conversation(title)

        # Save the user message first
        user_msg_id = request.userMessageId or str(_uuid.uuid4())
        range_tokens = None
        if request.rangeTokens:
            range_tokens = [{"address": t.address, "sheetName": t.sheetName} for t in request.rangeTokens]
        await append_conv_message(
            conversation_id=conv_id, message_id=user_msg_id, role="user",
            content=request.userMessage, range_tokens=range_tokens,
        )

        # Save the assistant message (may carry a plan)
        assistant_msg_id = str(_uuid.uuid4())
        plan_json: object | None = None
        if result.plans and len(result.plans) > 0:
            plan_json = result.plans[0].plan.model_dump(mode="json")
        elif result.plan:
            plan_json = result.plan.model_dump(mode="json")
        await append_conv_message(
            conversation_id=conv_id, message_id=assistant_msg_id, role="assistant",
            content=result.message, plan=plan_json,
        )

        result.conversationId = conv_id
        result.assistantMessageId = assistant_msg_id
    except Exception as exc:  # noqa: BLE001
        logger.warning("Failed to persist conversation turn: %s", exc)
