"""
Tests for the AnalyticalPlanner — JSON parsing, plan validation, fallback.
"""
from __future__ import annotations

import json
import pytest
from unittest.mock import AsyncMock, MagicMock, patch

from app.planner.planner import AnalyticalPlanner, _extract_json
from app.models.analytical_plan import (
    AnalyticalPlan,
    IntentType,
    OperationType,
    SheetData,
)


# ── Fixtures ──────────────────────────────────────────────────────────────────


@pytest.fixture()
def customers_sheet() -> SheetData:
    return SheetData(
        name="Customers",
        data=[
            ["CustomerID", "Name", "City", "Email"],
            ["C001", "Alice Johnson", "New York", "alice@example.com"],
            ["C002", "Bob Smith", "Boston", "bob@example.com"],
            ["C003", "Carol White", "Chicago", "carol@example.com"],
            ["C004", "Dan Brown", "New York", "dan@example.com"],
            ["C005", "Eve Davis", "Boston", "eve@example.com"],
        ],
    )


@pytest.fixture()
def orders_sheet() -> SheetData:
    return SheetData(
        name="Orders",
        data=[
            ["OrderID", "Customer", "Location", "Amount"],
            ["O001", "Alise Johnson", "New York", 250.0],
            ["O002", "Bob Smyth", "Boston", 180.0],
            ["O003", "Carol White", "Chicago", 320.0],
        ],
    )


@pytest.fixture()
def both_sheets(customers_sheet, orders_sheet) -> dict[str, SheetData]:
    return {"Customers": customers_sheet, "Orders": orders_sheet}


def _mock_llm_response(content: str):
    """Build a mock LiteLLM acompletion response."""
    msg = MagicMock()
    msg.content = content
    choice = MagicMock()
    choice.message = msg
    resp = MagicMock()
    resp.choices = [choice]
    return resp


_VALID_MATCH_JSON = json.dumps({
    "intent": "match_rows",
    "confidence": 0.88,
    "needs_clarification": False,
    "clarification_question": None,
    "selected_tool_chain": [
        "profile_columns",
        "estimate_matchability",
        "run_hybrid_match",
        "explain_match_result",
    ],
    "parameters": {
        "left_sheet": "Customers",
        "right_sheet": "Orders",
        "left_columns": ["Name", "City"],
        "right_columns": ["Customer", "Location"],
        "strategy": "hybrid",
    },
    "reasoning_summary": "Use fuzzy matching on Name and exact on City.",
})

_VALID_AGGREGATE_JSON = json.dumps({
    "intent": "aggregate",
    "confidence": 0.92,
    "needs_clarification": False,
    "clarification_question": None,
    "selected_tool_chain": ["profile_columns", "aggregate_values"],
    "parameters": {
        "sheet_name": "Customers",
        "group_by": ["City"],
        "metrics": [{"column": "CustomerID", "function": "count"}],
        "strategy": "exact",
    },
    "reasoning_summary": "Group by City and count customers per city.",
})

_VALID_DUPLICATES_JSON = json.dumps({
    "intent": "find_duplicates",
    "confidence": 0.95,
    "needs_clarification": False,
    "clarification_question": None,
    "selected_tool_chain": ["find_duplicates"],
    "parameters": {
        "sheet_name": "Customers",
        "key_columns": ["Name", "City"],
        "mode": "all",
    },
    "reasoning_summary": "Find rows where Name and City combination is duplicated.",
})

_CLARIFICATION_JSON = json.dumps({
    "intent": "match_rows",
    "confidence": 0.45,
    "needs_clarification": True,
    "clarification_question": "Which columns should I use to match the two sheets?",
    "selected_tool_chain": [],
    "parameters": {},
    "reasoning_summary": "Insufficient context to determine matching columns.",
})


# ── Tests ─────────────────────────────────────────────────────────────────────


@pytest.mark.asyncio
async def test_plan_parses_valid_match_rows_json(both_sheets):
    planner = AnalyticalPlanner()
    with patch("litellm.acompletion", new_callable=AsyncMock) as mock_llm:
        mock_llm.return_value = _mock_llm_response(_VALID_MATCH_JSON)
        plan = await planner.plan("Match customers with orders by name and city", both_sheets)

    assert plan.intent == IntentType.match_rows
    assert plan.confidence == pytest.approx(0.88)
    assert plan.needs_clarification is False
    assert OperationType.run_hybrid_match in plan.selected_tool_chain
    assert plan.parameters["left_sheet"] == "Customers"
    assert plan.parameters["right_sheet"] == "Orders"


@pytest.mark.asyncio
async def test_plan_parses_valid_aggregate_json(customers_sheet):
    planner = AnalyticalPlanner()
    sheets = {"Customers": customers_sheet}
    with patch("litellm.acompletion", new_callable=AsyncMock) as mock_llm:
        mock_llm.return_value = _mock_llm_response(_VALID_AGGREGATE_JSON)
        plan = await planner.plan("Group customers by city and count them", sheets)

    assert plan.intent == IntentType.aggregate
    assert OperationType.aggregate_values in plan.selected_tool_chain
    assert plan.parameters.get("group_by") == ["City"]


@pytest.mark.asyncio
async def test_plan_parses_valid_find_duplicates_json(customers_sheet):
    planner = AnalyticalPlanner()
    sheets = {"Customers": customers_sheet}
    with patch("litellm.acompletion", new_callable=AsyncMock) as mock_llm:
        mock_llm.return_value = _mock_llm_response(_VALID_DUPLICATES_JSON)
        plan = await planner.plan("Find duplicate customers", sheets)

    assert plan.intent == IntentType.find_duplicates
    assert OperationType.find_duplicates in plan.selected_tool_chain


@pytest.mark.asyncio
async def test_plan_returns_fallback_on_invalid_json(both_sheets):
    planner = AnalyticalPlanner()
    with patch("litellm.acompletion", new_callable=AsyncMock) as mock_llm:
        mock_llm.return_value = _mock_llm_response("This is not JSON at all, sorry!")
        plan = await planner.plan("Match rows", both_sheets)

    assert plan.needs_clarification is True
    assert plan.confidence == pytest.approx(0.0)
    assert plan.intent == IntentType.ask_clarification


@pytest.mark.asyncio
async def test_plan_returns_clarification_when_needed(both_sheets):
    planner = AnalyticalPlanner()
    with patch("litellm.acompletion", new_callable=AsyncMock) as mock_llm:
        mock_llm.return_value = _mock_llm_response(_CLARIFICATION_JSON)
        plan = await planner.plan("Do something with the data", both_sheets)

    assert plan.needs_clarification is True
    assert plan.clarification_question is not None
    assert len(plan.clarification_question) > 5


@pytest.mark.asyncio
async def test_plan_includes_correct_tool_chain_for_match(both_sheets):
    planner = AnalyticalPlanner()
    with patch("litellm.acompletion", new_callable=AsyncMock) as mock_llm:
        mock_llm.return_value = _mock_llm_response(_VALID_MATCH_JSON)
        plan = await planner.plan("Match by name and city", both_sheets)

    chain_values = [op.value for op in plan.selected_tool_chain]
    # A proper match plan should profile first, then estimate, then match
    assert "run_hybrid_match" in chain_values or any(
        op in chain_values for op in ["run_exact_match", "run_fuzzy_match", "run_semantic_match"]
    )


@pytest.mark.asyncio
async def test_plan_confidence_within_range(both_sheets):
    planner = AnalyticalPlanner()
    with patch("litellm.acompletion", new_callable=AsyncMock) as mock_llm:
        mock_llm.return_value = _mock_llm_response(_VALID_MATCH_JSON)
        plan = await planner.plan("match rows", both_sheets)

    assert 0.0 <= plan.confidence <= 1.0


@pytest.mark.asyncio
async def test_plan_with_multiple_sheets(customers_sheet, orders_sheet):
    sheets = {"Customers": customers_sheet, "Orders": orders_sheet}
    planner = AnalyticalPlanner()
    with patch("litellm.acompletion", new_callable=AsyncMock) as mock_llm:
        mock_llm.return_value = _mock_llm_response(_VALID_MATCH_JSON)
        plan = await planner.plan("Match Customers with Orders", sheets)

    # Both sheets should be referenced in parameters
    assert plan.parameters.get("left_sheet") in sheets or plan.parameters.get("right_sheet") in sheets


# ── _extract_json unit tests ──────────────────────────────────────────────────


def test_extract_json_parses_clean_json():
    raw = '{"intent": "match_rows", "confidence": 0.9}'
    result = _extract_json(raw)
    assert result["intent"] == "match_rows"


def test_extract_json_strips_markdown_fences():
    raw = '```json\n{"intent": "aggregate", "confidence": 0.8}\n```'
    result = _extract_json(raw)
    assert result["intent"] == "aggregate"


def test_extract_json_fixes_trailing_commas():
    raw = '{"intent": "match_rows", "confidence": 0.9,}'
    result = _extract_json(raw)
    assert result["confidence"] == pytest.approx(0.9)


def test_extract_json_raises_on_garbage():
    with pytest.raises((ValueError, Exception)):
        _extract_json("not json at all, no braces")
