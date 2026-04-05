"""
AnalyticalPlanner: LLM-powered intent detection and tool-chain construction.

Takes a user message + sheet context and returns an AnalyticalPlan describing:
  - The detected intent (match_rows, aggregate, find_duplicates, …)
  - The recommended matching/analysis strategy
  - An ordered tool chain (list of OperationType) with parameters
  - Whether clarification is needed before proceeding

Uses the OpenAI SDK for the LLM call via the centralized llm_client module.
"""
from __future__ import annotations

import json
import logging
import re
from typing import Optional

from ..config import settings
from ..models.analytical_plan import (
    AnalyticalPlan,
    IntentType,
    SheetData,
)
from ..models.request import ConversationMessage

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# System prompt
# ---------------------------------------------------------------------------

_SYSTEM_PROMPT = """\
You are an analytical data-science assistant embedded in Excel AI Copilot.

Your job is to analyse the user's natural-language request and the available
sheet data, then produce a JSON AnalyticalPlan describing:
  1. The detected intent.
  2. The best strategy for matching/analysis.
  3. An ordered list of operations (tool_chain) with parameters.
  4. Whether you need clarification from the user.

SUPPORTED INTENTS (use exact string values):
  answer_question, ask_clarification, filter_rows, aggregate,
  group_and_summarize, match_rows, find_duplicates, clean_data,
  compare_sheets, semantic_lookup, profile_sheet

SUPPORTED STRATEGIES: exact, fuzzy, semantic, hybrid

SUPPORTED OPERATIONS (use exact string values):
  list_sheets, get_sheet_schema, preview_sheet, profile_columns,
  clean_columns, estimate_matchability, run_exact_match, run_fuzzy_match,
  run_semantic_match, run_hybrid_match, find_duplicates, aggregate_values,
  filter_rows, compare_sheets, explain_match_result

OUTPUT SCHEMA — respond with ONLY this JSON:
{
  "intent": "<IntentType>",
  "confidence": 0.0-1.0,
  "needs_clarification": false,
  "clarification_question": null,
  "selected_tool_chain": ["<OperationType>", ...],
  "parameters": {
    "left_sheet": "<sheet name>",
    "right_sheet": "<sheet name or null>",
    "left_columns": ["<col>", ...],
    "right_columns": ["<col>", ...],
    "strategy": "<StrategyType>",
    "threshold": 80,
    "group_by_columns": ["<col>", ...],
    "agg_column": "<col>",
    "agg_function": "sum"
  },
  "reasoning_summary": "<one sentence>"
}

RULES:
- Respond ONLY with the JSON object. No prose, no markdown fences.
- If you cannot determine the intent with confidence >= 0.5, set
  needs_clarification=true and write a clear clarification_question.
- For match_rows: include estimate_matchability + run_*_match in the chain.
- For aggregate / group_and_summarize: include aggregate_values.
- For find_duplicates: include find_duplicates in the chain.
- For compare_sheets: include compare_sheets.
- Always end the chain with explain_match_result when the intent involves rows.
"""


# ---------------------------------------------------------------------------
# AnalyticalPlanner
# ---------------------------------------------------------------------------

class AnalyticalPlanner:
    """
    Calls the LLM to produce an AnalyticalPlan from a user message and
    available sheet context.
    """

    def __init__(
        self,
        model: str | None = None,
        temperature: float | None = None,
        max_tokens: int | None = None,
    ) -> None:
        self._model = model or settings.llm_model
        self._temperature = temperature if temperature is not None else settings.llm_temperature
        self._max_tokens = max_tokens or settings.llm_max_tokens

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    async def plan(
        self,
        user_message: str,
        sheets: dict[str, SheetData],
        conversation_history: Optional[list[ConversationMessage]] = None,
    ) -> AnalyticalPlan:
        """
        Build an AnalyticalPlan for *user_message* given the available sheets.

        Raises ValueError only if the LLM returns completely unrecoverable
        output — callers should handle this and show a fallback.
        """
        messages = self._build_messages(user_message, sheets, conversation_history)

        try:
            from ..services.llm_client import acompletion
            raw_text = await acompletion(
                messages=messages,
                temperature=self._temperature,
                max_tokens=self._max_tokens,
            )
        except Exception as exc:
            logger.warning("LLM call failed in AnalyticalPlanner.plan: %s", exc)
            return self._fallback_plan(user_message, str(exc))

        logger.debug("AnalyticalPlanner raw LLM response: %.400s", raw_text)

        try:
            plan_dict = _extract_json(raw_text)
            plan = AnalyticalPlan(**plan_dict)
            return plan
        except Exception as exc:
            logger.warning("Failed to parse AnalyticalPlan JSON: %s", exc)
            return self._fallback_plan(user_message, str(exc))

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _build_messages(
        self,
        user_message: str,
        sheets: dict[str, SheetData],
        conversation_history: Optional[list[ConversationMessage]],
    ) -> list[dict]:
        messages: list[dict] = [{"role": "system", "content": _SYSTEM_PROMPT}]

        if conversation_history:
            for msg in conversation_history[-6:]:
                messages.append({"role": msg.role, "content": msg.content})

        # Build context block summarising available sheets
        context_parts = [f"User request: {user_message}", "", "Available sheets:"]
        for name, sheet in sheets.items():
            headers = sheet.header_row
            row_count = len(sheet.data_rows)
            preview = sheet.data_rows[:3] if sheet.data_rows else []
            context_parts.append(
                f"  - {name}: columns={headers}, rows={row_count}, preview={preview}"
            )

        messages.append({"role": "user", "content": "\n".join(context_parts)})
        return messages

    @staticmethod
    def _fallback_plan(user_message: str, reason: str) -> AnalyticalPlan:
        """Return a safe fallback plan that asks for clarification."""
        return AnalyticalPlan(
            intent=IntentType.ask_clarification,
            confidence=0.0,
            needs_clarification=True,
            clarification_question=(
                "I wasn't able to understand your request well enough to proceed. "
                "Could you describe more specifically what you'd like to analyse?"
            ),
            selected_tool_chain=[],
            parameters={},
            reasoning_summary=f"Fallback plan due to: {reason[:200]}",
        )


# ---------------------------------------------------------------------------
# JSON extraction (reuses logic from existing planner service)
# ---------------------------------------------------------------------------

def _extract_json(text: str) -> dict:
    """Extract the first JSON object from LLM output text."""
    text = text.strip()

    # Strip markdown fences
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

    # Find outermost {}
    if not text.startswith("{"):
        try:
            first = text.index("{")
            last = text.rindex("}")
            text = text[first: last + 1]
        except ValueError:
            pass

    # Clean trailing commas
    text = re.sub(r",\s*([}\]])", r"\1", text)

    # First try: standard parse
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Second try: json-repair
    try:
        from json_repair import repair_json
        repaired = repair_json(text, return_objects=True)
        if isinstance(repaired, dict):
            return repaired
    except Exception:
        pass

    raise ValueError(f"Could not parse JSON from LLM response: {text[:300]}")
