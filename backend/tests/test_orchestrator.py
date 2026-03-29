"""
Tests for the Orchestrator: tool chaining, validation, and failure handling.
"""
from __future__ import annotations

import pytest
from unittest.mock import MagicMock, patch

from app.orchestrator.orchestrator import Orchestrator, OrchestratorResult
from app.orchestrator.execution_context import ExecutionContext
from app.orchestrator.validators import validate_plan, ValidationResult
from app.models.analytical_plan import (
    AnalyticalPlan,
    IntentType,
    OperationType,
    SheetData,
)
from app.models.tool_output import ToolOutput


# ── Fixtures ──────────────────────────────────────────────────────────────────


@pytest.fixture()
def customers() -> SheetData:
    return SheetData(
        name="Customers",
        data=[
            ["CustomerID", "Name", "City"],
            ["C001", "Alice Johnson", "New York"],
            ["C002", "Bob Smith", "Boston"],
            ["C003", "Carol White", "Chicago"],
        ],
    )


@pytest.fixture()
def orders() -> SheetData:
    return SheetData(
        name="Orders",
        data=[
            ["OrderID", "CustomerName", "Location"],
            ["O001", "Alice Johnson", "New York"],
            ["O002", "Bob Smyth", "Boston"],
        ],
    )


@pytest.fixture()
def sheets(customers, orders) -> dict[str, SheetData]:
    return {"Customers": customers, "Orders": orders}


def _make_plan(
    intent: IntentType,
    tool_chain: list[OperationType],
    parameters: dict | None = None,
    needs_clarification: bool = False,
    clarification_question: str | None = None,
) -> AnalyticalPlan:
    return AnalyticalPlan(
        intent=intent,
        confidence=0.9,
        needs_clarification=needs_clarification,
        clarification_question=clarification_question,
        selected_tool_chain=tool_chain,
        parameters=parameters or {},
        reasoning_summary="Test plan",
    )


# ── Validator tests ───────────────────────────────────────────────────────────


def test_validation_rejects_missing_sheet(sheets):
    plan = _make_plan(
        IntentType.match_rows,
        [OperationType.profile_columns],
        {"sheet_name": "NonExistent"},
    )
    result = validate_plan(plan, sheets)
    assert not result.valid
    assert any("NonExistent" in e for e in result.errors)


def test_validation_rejects_missing_column(customers, sheets):
    plan = _make_plan(
        IntentType.profile_sheet,
        [OperationType.profile_columns],
        {"sheet_name": "Customers", "columns": ["NonExistentColumn"]},
    )
    result = validate_plan(plan, sheets)
    assert not result.valid
    assert any("NonExistentColumn" in e for e in result.errors)


def test_validation_rejects_empty_tool_chain(sheets):
    plan = _make_plan(IntentType.match_rows, [], {})
    result = validate_plan(plan, sheets)
    assert not result.valid
    assert any("empty" in e.lower() for e in result.errors)


def test_validation_passes_valid_plan(sheets):
    plan = _make_plan(
        IntentType.profile_sheet,
        [OperationType.profile_columns],
        {"sheet_name": "Customers", "columns": ["Name", "City"]},
    )
    result = validate_plan(plan, sheets)
    assert result.valid
    assert not result.errors


def test_validation_rejects_out_of_range_threshold(sheets):
    plan = _make_plan(
        IntentType.match_rows,
        [OperationType.run_fuzzy_match],
        {"left_sheet": "Customers", "right_sheet": "Orders", "fuzzy_threshold": 1.5},
    )
    result = validate_plan(plan, sheets)
    assert not result.valid
    assert any("fuzzy_threshold" in e for e in result.errors)


def test_validation_warns_on_clarification_without_question(sheets):
    plan = AnalyticalPlan(
        intent=IntentType.ask_clarification,
        confidence=0.3,
        needs_clarification=True,
        clarification_question=None,  # Missing!
        selected_tool_chain=[OperationType.list_sheets],
        parameters={},
        reasoning_summary="test",
    )
    result = validate_plan(plan, sheets)
    assert not result.valid
    assert any("clarification_question" in e for e in result.errors)


# ── Orchestrator tests ────────────────────────────────────────────────────────


def test_orchestrator_executes_profile_columns(sheets):
    plan = _make_plan(
        IntentType.profile_sheet,
        [OperationType.profile_columns],
        {"sheet_name": "Customers", "columns": ["Name", "City"]},
    )
    orch = Orchestrator(sheets)
    result = orch.execute(plan)

    assert result.success
    assert result.context is not None
    assert "profile_columns" in result.context.tool_sequence


def test_orchestrator_executes_find_duplicates(sheets):
    plan = _make_plan(
        IntentType.find_duplicates,
        [OperationType.find_duplicates],
        {"sheet_name": "Customers", "key_columns": ["City"], "mode": "all"},
    )
    orch = Orchestrator(sheets)
    result = orch.execute(plan)

    assert result.success
    assert result.context is not None
    ctx_data = result.context.get_data("find_duplicates")
    assert ctx_data is not None


def test_orchestrator_executes_aggregate(sheets):
    plan = _make_plan(
        IntentType.aggregate,
        [OperationType.aggregate_values],
        {
            "sheet_name": "Customers",
            "group_by_columns": ["City"],
            "agg_column": "CustomerID",
            "agg_function": "count",
        },
    )
    orch = Orchestrator(sheets)
    result = orch.execute(plan)

    assert result.success


def test_orchestrator_stops_on_tool_failure(sheets):
    plan = _make_plan(
        IntentType.profile_sheet,
        [OperationType.profile_columns],
        {"sheet_name": "Customers", "columns": ["NonExistentColumn"]},
    )
    orch = Orchestrator(sheets)
    # profile_columns will warn but still succeed (just skip missing cols)
    # To force failure, reference a missing sheet
    plan2 = _make_plan(
        IntentType.profile_sheet,
        [OperationType.profile_columns],
        {"sheet_name": "MissingSheet", "columns": ["Name"]},
    )
    result = orch.execute(plan2)
    # Should fail at validation or tool dispatch
    assert not result.success


def test_orchestrator_returns_clarification_when_plan_requests_it(sheets):
    plan = _make_plan(
        IntentType.ask_clarification,
        [OperationType.list_sheets],
        {},
        needs_clarification=True,
        clarification_question="Which columns should I use?",
    )
    orch = Orchestrator(sheets)
    result = orch.execute(plan)

    assert result.needs_clarification
    assert result.clarification_question == "Which columns should I use?"


def test_orchestrator_profile_then_list_sheets(sheets):
    plan = _make_plan(
        IntentType.profile_sheet,
        [OperationType.list_sheets, OperationType.profile_columns],
        {"sheet_name": "Customers", "columns": ["Name"]},
    )
    orch = Orchestrator(sheets)
    result = orch.execute(plan)

    assert result.success
    assert "list_sheets" in result.context.tool_sequence
    assert "profile_columns" in result.context.tool_sequence


# ── ExecutionContext unit tests ───────────────────────────────────────────────


def test_execution_context_stores_and_retrieves():
    ctx = ExecutionContext()
    out = ToolOutput.ok("test_tool", data={"key": "value"})
    ctx.store("test_tool", out)

    assert ctx.has("test_tool")
    assert ctx.get("test_tool") is out
    assert ctx.get_data("test_tool") == {"key": "value"}
    assert ctx.get_data("test_tool", "key") == "value"
    assert ctx.get_data("test_tool", "missing") is None


def test_execution_context_last_result():
    ctx = ExecutionContext()
    ctx.store("tool_a", ToolOutput.ok("tool_a", data="first"))
    ctx.store("tool_b", ToolOutput.ok("tool_b", data="second"))

    last = ctx.last_result()
    assert last is not None
    assert last.data == "second"


def test_execution_context_all_warnings():
    ctx = ExecutionContext()
    ctx.store("tool_a", ToolOutput.ok("tool_a", warnings=["w1", "w2"]))
    ctx.store("tool_b", ToolOutput.ok("tool_b", warnings=["w3"]))

    warnings = ctx.all_warnings()
    assert len(warnings) == 3
    assert all("tool_a" in w or "tool_b" in w for w in warnings)


def test_execution_context_to_summary():
    ctx = ExecutionContext()
    ctx.log("started")
    ctx.store("my_tool", ToolOutput.ok("my_tool"))
    ctx.log("done")

    summary = ctx.to_summary()
    assert "my_tool" in summary["tool_sequence"]
    assert len(summary["execution_log"]) >= 2


def test_execution_context_does_not_duplicate_tool_sequence():
    ctx = ExecutionContext()
    ctx.store("tool_a", ToolOutput.ok("tool_a", data=1))
    ctx.store("tool_a", ToolOutput.ok("tool_a", data=2))  # overwrite

    assert ctx.tool_sequence.count("tool_a") == 1
    assert ctx.get_data("tool_a") == 2
