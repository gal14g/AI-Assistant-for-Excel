"""
Tests for plan validation logic.

These tests verify the validator catches schema errors, business rule
violations, and produces correct warnings — all without needing an LLM
or Office.js runtime.
"""

import pytest
from app.models.plan import ExecutionPlan, PlanStep, StepAction
from app.services.validator import validate_plan


def make_plan(steps: list[dict], **overrides) -> ExecutionPlan:
    """Helper to build a plan with defaults."""
    defaults = {
        "planId": "test-plan-1",
        "createdAt": "2025-01-01T00:00:00Z",
        "userRequest": "test request",
        "summary": "test summary",
        "steps": [PlanStep(**s) for s in steps],
        "preserveFormatting": True,
        "confidence": 0.9,
    }
    defaults.update(overrides)
    return ExecutionPlan(**defaults)


class TestBasicValidation:
    def test_valid_simple_plan(self):
        plan = make_plan([
            {
                "id": "step_1",
                "description": "Read data",
                "action": "readRange",
                "params": {"range": "A1:B10"},
            }
        ])
        result = validate_plan(plan)
        assert result.valid
        assert len(result.errors) == 0

    def test_valid_multi_step_plan(self):
        plan = make_plan([
            {
                "id": "step_1",
                "description": "Read data",
                "action": "readRange",
                "params": {"range": "A1:B10"},
            },
            {
                "id": "step_2",
                "description": "Write formula",
                "action": "writeFormula",
                "params": {"cell": "C1", "formula": "=SUM(A1:A10)"},
                "dependsOn": ["step_1"],
            },
        ])
        result = validate_plan(plan)
        assert result.valid

    def test_empty_steps_rejected(self):
        """Plans must have at least one step."""
        with pytest.raises(Exception):
            make_plan([])


class TestDuplicateIds:
    def test_duplicate_step_ids(self):
        plan = make_plan([
            {
                "id": "step_1",
                "description": "First",
                "action": "readRange",
                "params": {"range": "A1:B10"},
            },
            {
                "id": "step_1",
                "description": "Duplicate",
                "action": "readRange",
                "params": {"range": "C1:D10"},
            },
        ])
        result = validate_plan(plan)
        assert not result.valid
        assert any(e.code == "DUPLICATE_STEP_ID" for e in result.errors)


class TestDependencyValidation:
    def test_invalid_dependency(self):
        plan = make_plan([
            {
                "id": "step_1",
                "description": "Read",
                "action": "readRange",
                "params": {"range": "A1:B10"},
                "dependsOn": ["nonexistent"],
            },
        ])
        result = validate_plan(plan)
        assert not result.valid
        assert any(e.code == "INVALID_DEPENDENCY" for e in result.errors)

    def test_dependency_cycle(self):
        plan = make_plan([
            {
                "id": "step_1",
                "description": "First",
                "action": "readRange",
                "params": {"range": "A1:B10"},
                "dependsOn": ["step_2"],
            },
            {
                "id": "step_2",
                "description": "Second",
                "action": "readRange",
                "params": {"range": "C1:D10"},
                "dependsOn": ["step_1"],
            },
        ])
        result = validate_plan(plan)
        assert not result.valid
        assert any(e.code == "DEPENDENCY_CYCLE" for e in result.errors)


class TestParamValidation:
    def test_missing_range_param(self):
        plan = make_plan([
            {
                "id": "step_1",
                "description": "Read without range",
                "action": "readRange",
                "params": {},  # missing 'range'
            },
        ])
        result = validate_plan(plan)
        assert not result.valid
        assert any(e.code == "INVALID_PARAMS" for e in result.errors)

    def test_write_formula_missing_formula(self):
        plan = make_plan([
            {
                "id": "step_1",
                "description": "Write formula",
                "action": "writeFormula",
                "params": {"cell": "A1"},  # missing 'formula'
            },
        ])
        result = validate_plan(plan)
        assert not result.valid

    def test_valid_write_values_params(self):
        plan = make_plan([
            {
                "id": "step_1",
                "description": "Write values",
                "action": "writeValues",
                "params": {
                    "range": "A1:B2",
                    "values": [["a", "b"], ["c", "d"]],
                },
            },
        ])
        result = validate_plan(plan)
        assert result.valid


class TestFormattingSafety:
    def test_formatting_warning_on_create_table(self):
        plan = make_plan(
            [
                {
                    "id": "step_1",
                    "description": "Create table",
                    "action": "createTable",
                    "params": {"range": "A1:B10", "tableName": "Table1"},
                },
            ],
            preserveFormatting=True,
        )
        result = validate_plan(plan)
        assert result.valid  # warnings don't block
        assert any(w.code == "FORMAT_SAFETY_WARNING" for w in result.warnings)

    def test_no_formatting_warning_when_disabled(self):
        plan = make_plan(
            [
                {
                    "id": "step_1",
                    "description": "Create table",
                    "action": "createTable",
                    "params": {"range": "A1:B10", "tableName": "Table1"},
                },
            ],
            preserveFormatting=False,
        )
        result = validate_plan(plan)
        assert not any(w.code == "FORMAT_SAFETY_WARNING" for w in result.warnings)


class TestRangeSafety:
    def test_full_column_warning(self):
        plan = make_plan([
            {
                "id": "step_1",
                "description": "Read full column",
                "action": "readRange",
                "params": {"range": "A:A"},
            },
        ])
        result = validate_plan(plan)
        assert any(w.code == "LARGE_RANGE_WARNING" for w in result.warnings)


class TestMatchRecordsValidation:
    def test_valid_match_records(self):
        plan = make_plan([
            {
                "id": "step_1",
                "description": "Match records",
                "action": "matchRecords",
                "params": {
                    "lookupRange": "Sheet2!B:B",
                    "sourceRange": "Sheet1!A:A",
                    "returnColumns": [2],
                    "matchType": "exact",
                    "outputRange": "Sheet2!C:C",
                },
            },
        ])
        result = validate_plan(plan)
        assert result.valid

    def test_match_records_missing_fields(self):
        plan = make_plan([
            {
                "id": "step_1",
                "description": "Match records",
                "action": "matchRecords",
                "params": {
                    "lookupRange": "B:B",
                    # missing sourceRange, returnColumns, outputRange
                },
            },
        ])
        result = validate_plan(plan)
        assert not result.valid
