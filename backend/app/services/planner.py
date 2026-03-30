"""
LLM Planner Service

Sends the user's request to any LLM supported by LiteLLM and receives a
structured JSON execution plan.  The LLM is instructed to output ONLY a typed
JSON plan — never executable code.

Provider selection is driven entirely by config.py / .env:
  - Set LLM_MODEL to any LiteLLM model string (see config.py for examples)
  - Set LLM_API_KEY if the provider requires one
  - Set LLM_BASE_URL for local / self-hosted endpoints (Ollama, LiteLLM proxy…)

LiteLLM translates to the correct provider SDK automatically, so this file
stays provider-agnostic.
"""

from __future__ import annotations

import json
import uuid
from datetime import datetime, timezone
from pathlib import Path

import litellm

from ..config import settings
from ..models.plan import ExecutionPlan
from ..models.request import PlanRequest

# Silence LiteLLM's verbose success logs in production
litellm.success_callback = []

PROMPTS_DIR = Path(__file__).parent.parent / "prompts"


def _load_prompt(filename: str) -> str:
    path = PROMPTS_DIR / filename
    if path.exists():
        return path.read_text()
    return ""


SYSTEM_PROMPT = _load_prompt("planner_system.txt")

CAPABILITY_DESCRIPTIONS = {
    "readRange": "Read values from a cell range. Params: range (string), includeHeaders (bool, optional)",
    "writeValues": "Write a 2D array of values to a range. ONLY writes values, never formatting. Params: range (string), values (2D array)",
    "writeFormula": "Write a formula to a cell, optionally fill down. PREFER this over writeValues when a native Excel formula can express the operation. Params: cell (string), formula (string starting with =), fillDown (int, optional)",
    "matchRecords": "Lookup/match records between ranges using XLOOKUP/VLOOKUP or composite key matching. Params: lookupRange, sourceRange, returnColumns (array of 1-based ints), matchType ('exact'|'approximate'), outputRange, preferFormula (bool, default true). SPECIAL: when the user wants to write a constant string (like 'pass') to a column for matched rows, set writeValue='pass' instead of returnColumns — this automatically uses composite multi-column matching if lookupRange/sourceRange span multiple columns (e.g. C:D). NEVER use writeValues to write match results row-by-row.",
    "groupSum": "Sum values grouped by a column using SUMIF or computed aggregation. Params: dataRange, groupByColumn (1-based int), sumColumn (1-based int), outputRange, preferFormula (bool, default true)",
    "createTable": "Convert a range into an Excel Table. Params: range, tableName, hasHeaders (bool), style (optional string)",
    "applyFilter": "Apply filters to a table or range. Params: tableNameOrRange, columnIndex (0-based), criteria {filterOn, values, operator, value}",
    "sortRange": "Sort a range by columns. Params: range, sortFields [{columnIndex (0-based), ascending}], hasHeaders (bool)",
    "createPivot": "Create a PivotTable. Only sourceRange is required — everything else is auto-detected. rows/values fields accept EITHER header names ('Department') OR column range addresses ('Sheet2!A:A') — the handler resolves range addresses to header names automatically. Params: sourceRange (required), rows (optional list of field names or column refs), values (optional list of {field, summarizeBy} where field is a name or ref), columns (optional), filters (optional), destinationRange (optional — new sheet created if omitted), pivotName (optional)",
    "createChart": "Create a chart. Params: dataRange, chartType ('columnClustered'|'bar'|'line'|'pie'|'area'|'scatter'), title (optional), position (optional)",
    "addConditionalFormat": "Apply conditional formatting. Params: range, ruleType ('cellValue'|'colorScale'|'dataBar'|'iconSet'|'text'|'formula'), operator ('greaterThan'|'greaterThanOrEqualTo'|'lessThan'|'lessThanOrEqualTo'|'between'|'notBetween'|'equalTo'|'notEqualTo'), values, format {fillColor, fontColor, bold}, formula (for ruleType='formula': Excel formula string e.g. '=$D2=\"\"' to highlight blank rows, '=$B2>$C2' for cross-column compare)",
    "cleanupText": "Clean up text values. Params: range, operations ['trim'|'lowercase'|'uppercase'|'properCase'|'removeNonPrintable'|'normalizeWhitespace'], outputRange (optional)",
    "removeDuplicates": "Remove duplicate rows. Params: range, columnIndexes (0-based array, optional)",
    "freezePanes": "Freeze rows/columns at a cell. Params: cell (string), sheetName (optional)",
    "findReplace": "Find and replace text. Params: find, replace, range (optional), matchCase, matchEntireCell",
    "addValidation": "Add data validation. Params: range, validationType ('list'|'wholeNumber'|'decimal'|'date'|'textLength'|'custom'), listValues (comma list for static dropdown), formula ('=Sheet2!A:A' for dynamic range dropdown OR custom formula for validationType='custom'), operator ('between'|'notBetween'|'equalTo'|'notEqualTo'|'greaterThan'|'greaterThanOrEqualTo'|'lessThan'|'lessThanOrEqualTo'), min, max",
    "addSheet": "Add a new worksheet. Params: sheetName",
    "renameSheet": "Rename a worksheet. Params: sheetName, newName",
    "deleteSheet": "Delete a worksheet. Params: sheetName",
    "copySheet": "Copy a worksheet. Params: sheetName, newName (optional)",
    "protectSheet":    "Protect a worksheet. Params: sheetName, password (optional)",
    "autoFitColumns":    "Auto-fit column widths to their content. Params: range (optional — omit to fit all used columns), sheetName (optional)",
    "mergeCells":        "Merge cells in a range. Params: range (string), mergeType ('merge'=everything into one cell, 'mergeAcross'=each row separately)",
    "setNumberFormat":   "Apply a number format to a range. Params: range (string), format (e.g. '#,##0.00', '0%', 'dd/mm/yyyy', '$#,##0.00', 'General')",
    "insertDeleteRows":  "Insert or delete rows/columns. Params: range (determines which rows/columns and how many), shiftDirection ('down'=insert rows above, 'up'=delete rows, 'right'=insert columns left, 'left'=delete columns)",
    "addSparkline":      "Add sparkline mini-charts inside cells — ideal for dashboards showing trends. Params: dataRange (source data, one row per sparkline), locationRange (cells where sparklines appear), sparklineType ('line'|'column'|'winLoss', default 'line'), color (optional hex)",
    "formatCells":       "Format cell appearance — font, colors, borders, alignment. Params: range, bold (bool), italic (bool), underline (bool), strikethrough (bool), fontSize (int), fontFamily (string e.g. 'Calibri'), fontColor (hex), fillColor (hex), horizontalAlignment ('left'|'center'|'right'|'justify'), verticalAlignment ('top'|'middle'|'bottom'), wrapText (bool), borders ({style: 'thin'|'medium'|'thick'|'dashed'|'dotted'|'double'|'none', color: hex, edges: ['top'|'bottom'|'left'|'right'|'all'|'outside'|'inside']}). All params except range are optional — only set what you need to change.",
    "clearRange":        "Clear a range's contents, formatting, or both. Params: range, clearType ('contents'=values+formulas, 'formats'=only formatting, 'all'=everything)",
    "hideShow":          "Hide or unhide rows, columns, or entire sheets. Params: target ('rows'|'columns'|'sheet'), rangeOrName (row range e.g. '2:5', column range e.g. 'A:C', or sheet name), hide (bool: true=hide, false=unhide)",
    "addComment":        "Add a comment/note to a cell. Params: cell (string), content (string), author (optional string)",
    "addHyperlink":      "Insert a hyperlink in a cell. Params: cell (string), url (string), displayText (optional — defaults to url)",
    "groupRows":         "Group or ungroup rows/columns for outline collapsing. Params: range (row range e.g. '3:8' or column range e.g. 'B:E'), operation ('group'|'ungroup')",
    "setRowColSize":     "Set row height or column width manually. Params: range (row range e.g. '1:1' or column range e.g. 'A:C'), dimension ('rowHeight'|'columnWidth'), size (number — points for rows, characters for columns)",
    "copyPasteRange":    "Copy a range and paste to another location. Params: sourceRange, destinationRange, pasteType ('all'|'values'|'formats'|'formulas', default 'all')",
}


def build_system_prompt(relevant_actions: list[str] | None = None) -> str:
    """Return the planner system prompt (file-based or inline fallback)."""
    if relevant_actions:
        filtered = {k: v for k, v in CAPABILITY_DESCRIPTIONS.items() if k in relevant_actions}
    else:
        filtered = CAPABILITY_DESCRIPTIONS

    if SYSTEM_PROMPT:
        caps = "\n".join(f"- `{k}`: {v}" for k, v in filtered.items())
        return SYSTEM_PROMPT.replace("{CAPABILITIES}", caps)

    caps = "\n".join(f"  - {k}: {v}" for k, v in filtered.items())
    return f"""You are an Excel AI Copilot planner. Your job is to convert natural-language spreadsheet commands into a structured JSON execution plan.

CRITICAL RULES:
1. NEVER output executable code. Only output a JSON execution plan.
2. ALWAYS preserve existing formatting unless the user explicitly asks to change it. Set preserveFormatting: true by default.
3. PREFER native Excel formulas (writeFormula with preferFormula: true) over computed values (writeValues) whenever possible. Formulas auto-update and are auditable.
4. Each step MUST have a clear, user-friendly description.
5. Use the exact action names and parameter schemas defined below.
6. The plan must be valid JSON conforming to the ExecutionPlan schema.

AVAILABLE ACTIONS:
{caps}

OUTPUT SCHEMA:
{{
  "planId": "unique-id",
  "createdAt": "ISO timestamp",
  "userRequest": "original user message",
  "summary": "human-readable summary of what the plan does",
  "steps": [
    {{
      "id": "step_1",
      "description": "What this step does",
      "action": "actionName",
      "params": {{ ... action-specific params ... }},
      "dependsOn": ["step_id"]
    }}
  ],
  "preserveFormatting": true,
  "confidence": 0.0-1.0,
  "warnings": ["optional warnings"]
}}

When the user references ranges like [[Sheet1!A1:C20]], use those exact references in your plan.
When a formula can solve the problem, prefer writeFormula with the appropriate Excel formula.
For lookups, prefer XLOOKUP (Excel 365+) or VLOOKUP.
For grouped aggregations, prefer SUMIF/SUMIFS.
Always respond with ONLY the JSON plan, no other text."""


def build_user_message(request: PlanRequest) -> str:
    """Build the user message that goes to the LLM."""
    parts = [request.userMessage]

    if request.rangeTokens:
        refs = ", ".join(f"{t.sheetName}!{t.address}" for t in request.rangeTokens)
        parts.append(f"\nReferenced ranges: {refs}")

    if request.activeSheet:
        parts.append(f"\nActive sheet: {request.activeSheet}")

    if request.workbookName:
        parts.append(f"\nWorkbook: {request.workbookName}")

    if request.workbookPath:
        parts.append(f"\nWorkbook path: {request.workbookPath}")

    return "\n".join(parts)


def _litellm_kwargs() -> dict:
    """
    Build the keyword arguments for every litellm call.

    Only non-empty optional fields are included so that litellm's own
    per-provider defaults are not overridden with blank strings.
    """
    kwargs: dict = {
        "model": settings.llm_model,
        "max_tokens": settings.llm_max_tokens,
        "temperature": settings.llm_temperature,
        "timeout": 300,
    }
    if settings.llm_api_key:
        kwargs["api_key"] = settings.llm_api_key
    if settings.llm_base_url:
        kwargs["api_base"] = settings.llm_base_url
    if settings.llm_api_version:
        kwargs["api_version"] = settings.llm_api_version
    if settings.llm_json_mode:
        kwargs["response_format"] = {"type": "json_object"}
    # Qwen3 enables "thinking mode" by default — it generates thousands of
    # internal reasoning tokens before the actual response, making it appear
    # to hang.  Disable it: we want direct JSON output, not chain-of-thought.
    if "qwen3" in settings.llm_model.lower():
        kwargs["extra_body"] = {"think": False}
    return kwargs


def _build_messages(request: PlanRequest, relevant_actions: list[str] | None = None) -> list[dict]:
    """
    Assemble the message list.

    LiteLLM uses the OpenAI message format universally.  For providers that
    have a separate system-prompt field (Anthropic, Gemini), LiteLLM extracts
    it automatically from the first system-role message.
    """
    messages: list[dict] = [{"role": "system", "content": build_system_prompt(relevant_actions)}]

    if request.conversationHistory:
        for msg in request.conversationHistory[-6:]:
            messages.append({"role": msg.role, "content": msg.content})

    messages.append({"role": "user", "content": build_user_message(request)})
    return messages


async def generate_plan(request: PlanRequest) -> tuple[ExecutionPlan, str]:
    """
    Generate an execution plan from the user's request (non-streaming).

    Returns (plan, explanation) tuple.
    """
    from .capability_store import search_capabilities

    relevant_actions = search_capabilities(request.userMessage)

    response = await litellm.acompletion(
        messages=_build_messages(request, relevant_actions),
        **_litellm_kwargs(),
    )

    response_text: str = response.choices[0].message.content or ""
    plan_json = extract_json(response_text)
    _fill_defaults(plan_json, request)

    plan = ExecutionPlan(**plan_json)
    return plan, plan.summary



# ── Helpers ───────────────────────────────────────────────────────────────────

def _fill_defaults(plan_json: dict, request: PlanRequest) -> None:
    """Ensure required top-level fields are present."""
    if "planId" not in plan_json:
        plan_json["planId"] = str(uuid.uuid4())
    if "createdAt" not in plan_json:
        plan_json["createdAt"] = datetime.now(timezone.utc).isoformat()
    if "userRequest" not in plan_json:
        plan_json["userRequest"] = request.userMessage


def _clean_json_text(text: str) -> str:
    """
    Fix common LLM JSON generation issues before parsing:
    - trailing commas before } or ]
    - stray control characters
    """
    import re
    # Remove trailing commas before closing braces/brackets
    text = re.sub(r",\s*([}\]])", r"\1", text)
    # Remove control characters except standard whitespace
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", text)
    return text


def extract_json(text: str) -> dict:
    """
    Extract a JSON object from LLM response text.

    Handles:
    - Bare JSON (starts with {)
    - Markdown code fences (```json ... ```)
    - Trailing commas and minor formatting issues
    - Falls back to json-repair for deeply malformed output
    """
    import re

    text = text.strip()

    # Strip markdown fences first
    if "```json" in text:
        m = re.search(r"```json\s*(.*?)```", text, re.DOTALL)
        if m:
            text = m.group(1).strip()
    elif "```" in text:
        m = re.search(r"```\s*(.*?)```", text, re.DOTALL)
        if m:
            candidate = m.group(1).strip()
            if candidate.startswith("{"):
                text = candidate

    # Extract the outermost JSON object if surrounded by prose
    if not text.startswith("{"):
        try:
            first_brace = text.index("{")
            last_brace = text.rindex("}")
            text = text[first_brace : last_brace + 1]
        except ValueError:
            pass

    # First try: standard parse
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Second try: clean trailing commas etc. then parse
    cleaned = _clean_json_text(text)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        pass

    # Third try: json-repair (handles deeply malformed LLM output)
    try:
        from json_repair import repair_json
        repaired = repair_json(cleaned, return_objects=True)
        if isinstance(repaired, dict):
            return repaired
        return json.loads(repair_json(cleaned))
    except Exception:
        pass

    raise ValueError(f"Could not extract valid JSON from LLM response. First 200 chars: {text[:200]}")
