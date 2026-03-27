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

MULTI-STEP PLANS:
- Use as many steps as the task genuinely requires — do not artificially limit to one step
- Use "dependsOn": ["step_id"] when a step must run after another (e.g. step_2 depends on step_1)
- Steps without dependsOn run first; steps with dependsOn run after their dependencies complete
- Examples of when to use multiple steps:
    - Read data → clean it → write results → create chart (4 steps)
    - Add new sheet → write headers → create table → create pivot (4 steps)
    - Match records → sort result → add conditional formatting (3 steps)
    - Clean up column → remove duplicates → auto-fit columns (3 steps)

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

    Each pair shows a realistic user phrasing and the ideal JSON response.
    Covers: greetings, questions, all major action types, range tokens,
    multi-step plans with dependsOn, and optional-param omission.
    """
    return [

        # ── 1. Greeting / small talk → message ─────────────────────────────
        {"role": "user", "content": "hi"},
        {"role": "assistant", "content": '{"responseType":"message","message":"Hi! I\'m Copilot, your Excel assistant. Ask me anything or tell me what you\'d like to do with your spreadsheet.","plan":null}'},

        # ── 2. Capability question → message ───────────────────────────────
        {"role": "user", "content": "what can you do? / what are your capabilities? / how can you help me?"},
        {"role": "assistant", "content": '{"responseType":"message","message":"I can help you with a wide range of Excel tasks:\\n\\n- **Formulas** — write XLOOKUP, SUMIF, IF, and any other formula\\n- **Charts** — create column, bar, line, pie, scatter charts\\n- **Pivot tables** — build and configure pivot tables\\n- **Sorting & filtering** — sort by any column, apply filters\\n- **Conditional formatting** — highlight cells by value, color scales, data bars\\n- **Data cleaning** — trim spaces, fix case, remove duplicates\\n- **Lookup & match** — match records between sheets using XLOOKUP\\n- **Tables** — convert ranges to structured Excel tables\\n- **Sheet management** — add, rename, copy, delete sheets\\n\\nSelect a range in Excel, paste it here with Ctrl+V, and tell me what to do with it!","plan":null}'},

        # ── 3. Formula explanation → message ───────────────────────────────
        {"role": "user", "content": "what does XLOOKUP do? / explain VLOOKUP / when should I use SUMIF?"},
        {"role": "assistant", "content": '{"responseType":"message","message":"**XLOOKUP** searches a range for a value and returns a corresponding result from another range.\\n\\nExample: `=XLOOKUP(A2, Sheet2!A:A, Sheet2!B:B, \\"Not found\\")`\\n\\nThis looks for the value in A2 inside Sheet2 column A, and returns the matching value from Sheet2 column B. If nothing matches, it returns \\"Not found\\".\\n\\nIt\'s the modern replacement for VLOOKUP — it can search left or right, handles exact and approximate matches, and doesn\'t break when you insert columns.","plan":null}'},

        # ── 4. Create chart – plain text request ───────────────────────────
        {"role": "user", "content": "make a chart from my data / visualize column A and B / plot this as a bar chart\nActive sheet: Sheet1"},
        {"role": "assistant", "content": '{"responseType":"plan","message":"I\'ll create a bar chart from the data on Sheet1.","plan":{"planId":"ex-chart-1","createdAt":"2024-01-01T00:00:00Z","userRequest":"make a chart","summary":"Create bar chart from Sheet1 data","steps":[{"id":"step_1","description":"Create a bar chart from the sheet data","action":"createChart","params":{"dataRange":"Sheet1!A:B","chartType":"bar"},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.85,"warnings":[]}}'},

        # ── 5. Create chart – with range token ─────────────────────────────
        {"role": "user", "content": "create a line chart from [[Sheet1!A1:C20]]"},
        {"role": "assistant", "content": '{"responseType":"plan","message":"I\'ll create a line chart using the data in Sheet1!A1:C20.","plan":{"planId":"ex-chart-2","createdAt":"2024-01-01T00:00:00Z","userRequest":"create a line chart from [[Sheet1!A1:C20]]","summary":"Create line chart from Sheet1!A1:C20","steps":[{"id":"step_1","description":"Create a line chart from the selected range","action":"createChart","params":{"dataRange":"Sheet1!A1:C20","chartType":"line"},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.95,"warnings":[]}}'},

        # ── 6. Pivot table – minimal params ────────────────────────────────
        {"role": "user", "content": "create a pivot table / make a pivot from [[Sheet1!A1:D50]] / summarize this data with a pivot"},
        {"role": "assistant", "content": '{"responseType":"plan","message":"I\'ll create a PivotTable from your data. The fields will be auto-detected from your column headers.","plan":{"planId":"ex-pivot-1","createdAt":"2024-01-01T00:00:00Z","userRequest":"create a pivot table","summary":"Create PivotTable from Sheet1 data","steps":[{"id":"step_1","description":"Create a PivotTable (fields auto-detected from headers)","action":"createPivot","params":{"sourceRange":"Sheet1!A1:D50"},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.9,"warnings":[]}}'},

        # ── 7. Lookup / match between sheets ───────────────────────────────
        {"role": "user", "content": "fill column C in Sheet1 with matching values from Sheet2 based on column A / vlookup Sheet1 column A against Sheet2 / match IDs in [[Sheet1!A:A]] with [[Sheet2!A:A]] and bring over [[Sheet2!B:B]]"},
        {"role": "assistant", "content": '{"responseType":"plan","message":"I\'ll match values in Sheet1 column A against Sheet2 column A, and write the corresponding values from Sheet2 column B into Sheet1 column C using XLOOKUP formulas.","plan":{"planId":"ex-match-1","createdAt":"2024-01-01T00:00:00Z","userRequest":"fill column C in Sheet1 with matching values from Sheet2 based on column A","summary":"XLOOKUP Sheet1!A against Sheet2!A, write Sheet2!B to Sheet1!C","steps":[{"id":"step_1","description":"Match Sheet1 column A to Sheet2 column A and write Sheet2 column B values to Sheet1 column C","action":"matchRecords","params":{"lookupRange":"Sheet1!A:A","sourceRange":"Sheet2!A:A","returnColumns":[2],"matchType":"exact","outputRange":"Sheet1!C:C","preferFormula":true},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.92,"warnings":[]}}'},

        # ── 8. Sort ─────────────────────────────────────────────────────────
        {"role": "user", "content": "sort by column B / order the table by date descending / sort [[Sheet1!A1:D100]] by the second column\nActive sheet: Sheet1"},
        {"role": "assistant", "content": '{"responseType":"plan","message":"I\'ll sort the data by column B in descending order.","plan":{"planId":"ex-sort-1","createdAt":"2024-01-01T00:00:00Z","userRequest":"sort by column B descending","summary":"Sort Sheet1 data by column B descending","steps":[{"id":"step_1","description":"Sort the range by column B descending","action":"sortRange","params":{"range":"Sheet1!A:D","sortFields":[{"columnIndex":1,"ascending":false}],"hasHeaders":true},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.9,"warnings":[]}}'},

        # ── 9. Conditional formatting ───────────────────────────────────────
        {"role": "user", "content": "highlight cells above 100 in red / color rows where value is negative / mark cells greater than average"},
        {"role": "assistant", "content": '{"responseType":"plan","message":"I\'ll add conditional formatting to highlight cells with values above 100 in red.","plan":{"planId":"ex-cf-1","createdAt":"2024-01-01T00:00:00Z","userRequest":"highlight cells above 100 in red","summary":"Red conditional format for values > 100","steps":[{"id":"step_1","description":"Highlight cells greater than 100 with a red fill","action":"addConditionalFormat","params":{"range":"Sheet1!A:A","ruleType":"cellValue","operator":"greaterThan","values":[100],"format":{"fillColor":"#ffcccc","fontColor":"#c50f1f","bold":true}},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.92,"warnings":[]}}'},

        # ── 10. Write a formula ─────────────────────────────────────────────
        {"role": "user", "content": "add a SUM formula in C1 for column A / write a formula to calculate the total / put =AVERAGE(B:B) in D1"},
        {"role": "assistant", "content": '{"responseType":"plan","message":"I\'ll write a SUM formula in C1 that totals all values in column A.","plan":{"planId":"ex-formula-1","createdAt":"2024-01-01T00:00:00Z","userRequest":"add a SUM formula in C1 for column A","summary":"Write =SUM(A:A) in C1","steps":[{"id":"step_1","description":"Write SUM formula for column A into cell C1","action":"writeFormula","params":{"cell":"Sheet1!C1","formula":"=SUM(A:A)"},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.95,"warnings":[]}}'},

        # ── 11. Format as table ─────────────────────────────────────────────
        {"role": "user", "content": "make this a table / format as Excel table / convert [[Sheet1!A1:E50]] to a table"},
        {"role": "assistant", "content": '{"responseType":"plan","message":"I\'ll convert your data into a structured Excel Table with auto-filter and styling.","plan":{"planId":"ex-table-1","createdAt":"2024-01-01T00:00:00Z","userRequest":"make this a table","summary":"Convert range to Excel Table","steps":[{"id":"step_1","description":"Convert range to a formatted Excel Table","action":"createTable","params":{"range":"Sheet1!A1:E50","hasHeaders":true},"dependsOn":[]}],"preserveFormatting":true,"confidence":0.93,"warnings":[]}}'},

        # ── 12. Clean data → remove duplicates (multi-step) ────────────────
        {"role": "user", "content": "clean up the text in column A and then remove duplicates / trim spaces and deduplicate / fix text formatting then remove duplicate rows"},
        {"role": "assistant", "content": '{"responseType":"plan","message":"I\'ll first trim spaces and fix the text in column A, then remove any duplicate rows.","plan":{"planId":"ex-clean-1","createdAt":"2024-01-01T00:00:00Z","userRequest":"clean up the text in column A and then remove duplicates","summary":"Clean text in column A then remove duplicates","steps":[{"id":"step_1","description":"Trim whitespace and normalize text in column A","action":"cleanupText","params":{"range":"Sheet1!A:A","operations":["trim","normalizeWhitespace"]},"dependsOn":[]},{"id":"step_2","description":"Remove duplicate rows based on all columns","action":"removeDuplicates","params":{"range":"Sheet1!A:A"},"dependsOn":["step_1"]}],"preserveFormatting":true,"confidence":0.9,"warnings":[]}}'},

        # ── 13. Sort → create chart (multi-step) ───────────────────────────
        {"role": "user", "content": "sort by sales descending and then make a chart / order by column B and visualize it / sort the data and create a bar chart"},
        {"role": "assistant", "content": '{"responseType":"plan","message":"I\'ll sort the data by the sales column descending, then create a bar chart from the sorted data.","plan":{"planId":"ex-sort-chart-1","createdAt":"2024-01-01T00:00:00Z","userRequest":"sort by sales descending and then make a chart","summary":"Sort by sales descending then create bar chart","steps":[{"id":"step_1","description":"Sort data by column B (sales) descending","action":"sortRange","params":{"range":"Sheet1!A:B","sortFields":[{"columnIndex":1,"ascending":false}],"hasHeaders":true},"dependsOn":[]},{"id":"step_2","description":"Create a bar chart from the sorted data","action":"createChart","params":{"dataRange":"Sheet1!A:B","chartType":"columnClustered","title":"Sales Chart"},"dependsOn":["step_1"]}],"preserveFormatting":true,"confidence":0.88,"warnings":[]}}'},

        # ── 14. Add sheet → write headers → create table (multi-step) ──────
        {"role": "user", "content": "add a new sheet called Summary and create a table there / create a Summary sheet with a table / make a new sheet and set it up as a table"},
        {"role": "assistant", "content": '{"responseType":"plan","message":"I\'ll add a new sheet called Summary, write column headers, then convert that range to an Excel Table.","plan":{"planId":"ex-newsheet-1","createdAt":"2024-01-01T00:00:00Z","userRequest":"add a new sheet called Summary and create a table there","summary":"Add Summary sheet and create a table","steps":[{"id":"step_1","description":"Add a new worksheet named Summary","action":"addSheet","params":{"sheetName":"Summary"},"dependsOn":[]},{"id":"step_2","description":"Write column headers to the new sheet","action":"writeValues","params":{"range":"Summary!A1:D1","values":[["Category","Value","Date","Notes"]]},"dependsOn":["step_1"]},{"id":"step_3","description":"Convert the header row and data area to an Excel Table","action":"createTable","params":{"range":"Summary!A1:D1","hasHeaders":true},"dependsOn":["step_2"]}],"preserveFormatting":true,"confidence":0.87,"warnings":[]}}'},

        # ── 15. Analyze / explain data → message ───────────────────────────
        {"role": "user", "content": "analyze my data / what insights can you find? / summarize what's in [[Sheet1!A1:C50]]"},
        {"role": "assistant", "content": '{"responseType":"message","message":"I can\'t read your spreadsheet data directly, but I can help you analyze it once you tell me what\'s in it!\\n\\nTry describing your data — for example:\\n- \\"Column A is product names, B is sales, C is region\\"\\n- \\"I have 50 rows of transaction data with date, amount, category\\"\\n\\nOr use **Ctrl+V** to paste a range reference into your message (select a range in Excel first), and tell me what you\'re looking for — trends, totals, outliers, etc.","plan":null}'},
    ]


def _build_chat_messages(request: ChatRequest) -> list[dict]:
    messages: list[dict] = [{"role": "system", "content": _build_chat_system_prompt()}]

    # Few-shot examples teach the model the expected response format and patterns
    messages.extend(_few_shot_examples())

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
