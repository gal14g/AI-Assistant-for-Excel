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
3. Use exact range references from [[...]] tokens in the user message when provided.
   CRITICAL: Extract only the address INSIDE the [[...]] markers — do NOT include [[ or ]] in the JSON.
   Example: user says "column [[Sheet1!A:A]]" → use "Sheet1!A:A" in params (NOT "[[Sheet1!A:A]]")
   Example: user says "[[[WorkbookName.xlsx]Sheet1!A:A]]" → use "[WorkbookName.xlsx]Sheet1!A:A"
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

COMPLEX FORMULAS — writeFormula supports any Excel formula including:
- Nested functions: =IF(ISNUMBER(MATCH(A2,Sheet2!A:A,0)),"Found","Not Found")
- LAMBDA / LET: =LET(x, A2*1.2, IF(x>100, x, 0))
- Dynamic arrays (Excel 365): =UNIQUE(A:A), =FILTER(A:B, B:B>0), =SORT(A:A)
- Array formulas: =SUM(IF(A:A="X", B:B, 0)) — wrap in ARRAYFORMULA for older Excel
- Lookup chains: =IFERROR(XLOOKUP(A2,Sheet2!A:A,Sheet2!C:C),VLOOKUP(A2,Sheet3!A:C,3,0))
- When the user asks for a "complex formula" or "dynamic formula", use writeFormula

PIVOT FIELD RULES:
- rows and values accept either header names ("Department") or range addresses ("Sheet2!A:A")
- ALWAYS include rows and values in createPivot when the user specifies them — never drop them
- Example: user says "rows = [[Sheet2!A:A]]" → use rows: ["Sheet2!A:A"] (the handler resolves it)

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


def _few_shot_examples() -> list[dict]:
    """
    Few-shot examples injected as user/assistant turns before the real request.

    Kept to 7 high-signal pairs — enough to establish the JSON format and
    the message/plan split without bloating the context window (which causes
    some models to hallucinate prose instead of JSON).
    """
    return [
        # 1. Greeting → message
        {"role": "user", "content": "hi"},
        {"role": "assistant", "content": '{"responseType":"message","message":"Hi! I\'m Copilot, your Excel assistant. Tell me what you\'d like to do with your spreadsheet.","plan":null}'},

        # 2. Question → message
        {"role": "user", "content": "what can you do?"},
        {"role": "assistant", "content": '{"responseType":"message","message":"I can write formulas, create charts and pivot tables, sort and filter data, apply conditional formatting, clean text, remove duplicates, match records between sheets, manage sheets, and more. Select a range in Excel, paste it here with Ctrl+V, and tell me what you need!","plan":null}'},

        # 3. Match between sheets with range tokens (hardest pattern to get right)
        {"role": "user", "content": "write to [[Sheet2!C:C]] the values from [[Sheet1!B:B]] where [[Sheet1!A:A]] matches [[Sheet2!A:A]]"},
        {"role": "assistant", "content": '{"responseType":"plan","message":"I\'ll use XLOOKUP to match Sheet2 column A against Sheet1 column A, and write the corresponding Sheet1 column B values into Sheet2 column C.","plan":{"planId":"ex-match-1","createdAt":"2024-01-01T00:00:00Z","userRequest":"write to Sheet2!C:C values from Sheet1!B:B where Sheet1!A:A matches Sheet2!A:A","summary":"XLOOKUP Sheet1!B into Sheet2!C based on column A match","steps":[{"id":"step_1","description":"Match Sheet2!A against Sheet1!A and write Sheet1!B values into Sheet2!C","action":"matchRecords","params":{"lookupRange":"Sheet2!A:A","sourceRange":"Sheet1!A:A","returnColumns":[2],"matchType":"exact","outputRange":"Sheet2!C:C","preferFormula":true},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.95,"warnings":[]}}'},

        # 4. Chart with range token → plan
        {"role": "user", "content": "create a bar chart from [[Sheet1!A1:B20]]"},
        {"role": "assistant", "content": '{"responseType":"plan","message":"I\'ll create a bar chart using the data in Sheet1!A1:B20.","plan":{"planId":"ex-chart-1","createdAt":"2024-01-01T00:00:00Z","userRequest":"create a bar chart from [[Sheet1!A1:B20]]","summary":"Create bar chart from Sheet1!A1:B20","steps":[{"id":"step_1","description":"Create a bar chart from the selected range","action":"createChart","params":{"dataRange":"Sheet1!A1:B20","chartType":"bar"},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.95,"warnings":[]}}'},

        # 5a. Pivot table — auto-detect (user gives only source)
        {"role": "user", "content": "create a pivot table from [[Sheet1!A1:D50]]"},
        {"role": "assistant", "content": '{"responseType":"plan","message":"I\'ll create a PivotTable from the selected range. Fields will be auto-detected from the column headers.","plan":{"planId":"ex-pivot-1","createdAt":"2024-01-01T00:00:00Z","userRequest":"create a pivot table from [[Sheet1!A1:D50]]","summary":"Create PivotTable from Sheet1!A1:D50","steps":[{"id":"step_1","description":"Create PivotTable (rows and values auto-detected from headers)","action":"createPivot","params":{"sourceRange":"Sheet1!A1:D50"},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.92,"warnings":[]}}'},

        # 5b. Pivot table — user explicitly specifies rows and values via range tokens
        {"role": "user", "content": "create a pivot from [[Sheet2!A1:D7]], set rows to [[Sheet2!A:A]] and values to sum of [[Sheet2!D:D]]"},
        {"role": "assistant", "content": '{"responseType":"plan","message":"I\'ll create a PivotTable grouping by column A (rows) and summing column D values.","plan":{"planId":"ex-pivot-2","createdAt":"2024-01-01T00:00:00Z","userRequest":"create pivot from Sheet2!A1:D7, rows=Sheet2!A:A, values=sum Sheet2!D:D","summary":"PivotTable rows=col A, values=SUM col D","steps":[{"id":"step_1","description":"Create PivotTable with rows from column A and sum of column D as values","action":"createPivot","params":{"sourceRange":"Sheet2!A1:D7","rows":["Sheet2!A:A"],"values":[{"field":"Sheet2!D:D","summarizeBy":"sum"}]},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.97,"warnings":[]}}'},

        # 6. Multi-step: sort → chart
        {"role": "user", "content": "sort by column B descending then create a chart"},
        {"role": "assistant", "content": '{"responseType":"plan","message":"I\'ll sort the data by column B descending, then create a chart from the sorted result.","plan":{"planId":"ex-sort-chart","createdAt":"2024-01-01T00:00:00Z","userRequest":"sort by column B descending then create a chart","summary":"Sort by column B desc then create chart","steps":[{"id":"step_1","description":"Sort data by column B descending","action":"sortRange","params":{"range":"Sheet1!A:B","sortFields":[{"columnIndex":1,"ascending":false}],"hasHeaders":true},"dependsOn":[]},{"id":"step_2","description":"Create a chart from the sorted data","action":"createChart","params":{"dataRange":"Sheet1!A:B","chartType":"columnClustered"},"dependsOn":["step_1"]}],"preserveFormatting":true,"confidence":0.9,"warnings":[]}}'},

        # 7. Conditional formatting
        {"role": "user", "content": "highlight cells in column B above 100 in red"},
        {"role": "assistant", "content": '{"responseType":"plan","message":"I\'ll add a conditional formatting rule to highlight cells in column B with values greater than 100 in red.","plan":{"planId":"ex-cf-1","createdAt":"2024-01-01T00:00:00Z","userRequest":"highlight cells in column B above 100 in red","summary":"Red highlight for Sheet1!B values > 100","steps":[{"id":"step_1","description":"Apply red fill to cells in column B where value > 100","action":"addConditionalFormat","params":{"range":"Sheet1!B:B","ruleType":"cellValue","operator":"greaterThan","values":[100],"format":{"fillColor":"#ffcccc","fontColor":"#c50f1f"}},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.93,"warnings":[]}}'},
    ]


def _build_user_content(request: ChatRequest) -> str:
    parts = [request.userMessage]
    if request.rangeTokens:
        refs = ", ".join(f"[[{t.address}]]" for t in request.rangeTokens)
        parts.append(f"\nReferenced ranges: {refs}")
    if request.activeSheet:
        parts.append(f"\nActive sheet: {request.activeSheet}")
    if request.workbookName:
        parts.append(f"\nWorkbook: {request.workbookName}")
    return "\n".join(parts)


def _build_chat_messages(request: ChatRequest) -> list[dict]:
    messages: list[dict] = [{"role": "system", "content": _build_chat_system_prompt()}]

    # Few-shot examples teach the model the expected response format and patterns
    messages.extend(_few_shot_examples())

    if request.conversationHistory:
        for msg in request.conversationHistory[-8:]:
            messages.append({"role": msg.role, "content": msg.content})

    messages.append({"role": "user", "content": _build_user_content(request)})
    return messages


def _build_retry_messages(request: ChatRequest) -> list[dict]:
    """Stripped-down prompt for retry — no few-shot examples, harder JSON enforcement."""
    system = (
        _build_chat_system_prompt()
        + "\n\nCRITICAL: Your ENTIRE response must be ONE valid JSON object and nothing else. "
        "No prose, no markdown, no explanation — just the JSON object starting with { and ending with }."
    )
    return [
        {"role": "system", "content": system},
        {"role": "user", "content": _build_user_content(request)},
    ]


def _parse_response(text: str, request: ChatRequest) -> ChatResponse:
    parsed = extract_json(text)
    response_type = parsed.get("responseType", "message")
    message = parsed.get("message", "")

    if response_type == "plan" and parsed.get("plan"):
        plan_data: dict = parsed["plan"]
        if "planId" not in plan_data:
            plan_data["planId"] = str(uuid.uuid4())
        if "createdAt" not in plan_data:
            plan_data["createdAt"] = datetime.now(timezone.utc).isoformat()
        if "userRequest" not in plan_data:
            plan_data["userRequest"] = request.userMessage

        plan = ExecutionPlan(**plan_data)
        return ChatResponse(responseType="plan", message=message, plan=plan)

    return ChatResponse(responseType="message", message=message, plan=None)


async def chat(request: ChatRequest) -> ChatResponse:
    """
    Send a user message to the chat AI.
    Returns either a conversational reply or an execution plan.
    On JSON parse failure, retries once with a stripped-down prompt.
    """
    try:
        response = await litellm.acompletion(
            messages=_build_chat_messages(request),
            **_litellm_kwargs(),
        )
        text: str = response.choices[0].message.content or ""
        return _parse_response(text, request)
    except Exception:
        # Retry with no few-shot examples and a stronger JSON-only instruction
        response = await litellm.acompletion(
            messages=_build_retry_messages(request),
            **_litellm_kwargs(),
        )
        text = response.choices[0].message.content or ""
        return _parse_response(text, request)
