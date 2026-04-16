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


class TestComplexMultiStepPlan:
    """
    Stress-test the validator against the shape of plan the LLM actually
    produces for real data-analyst requests: 10+ steps, cross-step bindings
    via `{{step_N.outputRange}}` tokens, and mixed action categories
    (read → compute → summarise → chart → format).

    These tests don't exercise the binding resolver (that lives in the
    frontend executor + MCP bridge) — they only verify that the validator
    doesn't reject a legitimately long dependency chain.
    """

    def test_ten_step_plan_with_bindings(self):
        """10 steps, linear dependency chain, binding tokens in params."""
        plan = make_plan([
            {
                "id": "s1",
                "description": "Read raw transactions",
                "action": "readRange",
                "params": {"range": "Sheet1!A1:F5000"},
            },
            {
                "id": "s2",
                "description": "Clean header text",
                "action": "cleanupText",
                "params": {"range": "Sheet1!A1:F1", "operations": ["trim", "properCase"]},
                "dependsOn": ["s1"],
            },
            {
                "id": "s3",
                "description": "Normalise dates",
                "action": "normalizeDates",
                "params": {"range": "Sheet1!B2:B5000", "outputFormat": "yyyy-mm-dd"},
                "dependsOn": ["s2"],
            },
            {
                "id": "s4",
                "description": "Coerce numeric column",
                "action": "coerceDataType",
                "params": {"range": "Sheet1!E2:E5000", "targetType": "number"},
                "dependsOn": ["s3"],
            },
            {
                "id": "s5",
                "description": "Deduplicate by composite key",
                "action": "deduplicateAdvanced",
                "params": {
                    "range": "Sheet1!A1:F5000",
                    "keyColumns": [1, 2, 3],
                    "keepStrategy": "newest",
                    "dateColumn": 2,
                },
                "dependsOn": ["s4"],
            },
            {
                "id": "s6",
                "description": "Group totals by category",
                "action": "groupSum",
                "params": {
                    "dataRange": "Sheet1!A1:F5000",
                    "groupByColumn": 3,
                    "sumColumn": 5,
                    "outputRange": "Sheet2!A1",
                },
                "dependsOn": ["s5"],
            },
            {
                "id": "s7",
                "description": "Sort summary descending",
                "action": "sortRange",
                "params": {
                    "range": "Sheet2!A1:B200",
                    "sortFields": [{"columnIndex": 1, "ascending": False}],
                    "hasHeaders": True,
                },
                "dependsOn": ["s6"],
            },
            {
                "id": "s8",
                "description": "Add running total",
                "action": "runningTotal",
                "params": {
                    "sourceRange": "Sheet2!B2:B200",
                    "outputRange": "Sheet2!C2:C200",
                    "hasHeaders": False,
                },
                "dependsOn": ["s7"],
            },
            {
                "id": "s9",
                "description": "Highlight top earners",
                "action": "addConditionalFormat",
                "params": {
                    "range": "Sheet2!B2:B200",
                    "ruleType": "cellValue",
                    "operator": "greaterThan",
                    "values": [10000],
                    "format": {"fillColor": "#C6EFCE"},
                },
                "dependsOn": ["s8"],
            },
            {
                "id": "s10",
                "description": "Chart top categories",
                "action": "createChart",
                "params": {
                    "dataRange": "Sheet2!A1:B20",
                    "chartType": "bar",
                    "title": "Top categories by revenue",
                },
                "dependsOn": ["s9"],
            },
        ])
        result = validate_plan(plan)
        assert result.valid, [e.message for e in result.errors]

    def test_fan_out_fan_in_bindings(self):
        """
        Diamond dependency: s1 feeds s2 and s3 in parallel; s4 waits on both.
        Any regression in cycle detection that confuses "two edges from s1"
        with a back-edge will break this.
        """
        plan = make_plan([
            {
                "id": "s1",
                "description": "Read source",
                "action": "readRange",
                "params": {"range": "A1:C100"},
            },
            {
                "id": "s2",
                "description": "Sum by group",
                "action": "groupSum",
                "params": {
                    "dataRange": "A1:C100",
                    "groupByColumn": 1,
                    "sumColumn": 3,
                    "outputRange": "E1",
                },
                "dependsOn": ["s1"],
            },
            {
                "id": "s3",
                "description": "Frequency of groups",
                "action": "frequencyDistribution",
                "params": {"sourceRange": "A1:A100", "outputRange": "G1"},
                "dependsOn": ["s1"],
            },
            {
                "id": "s4",
                "description": "Join sums + frequencies",
                "action": "joinSheets",
                "params": {
                    "leftRange": "E1:F50",
                    "rightRange": "G1:H50",
                    "leftKeyColumn": 1,
                    "rightKeyColumn": 1,
                    "joinType": "left",
                    "outputRange": "J1",
                },
                "dependsOn": ["s2", "s3"],
            },
        ])
        result = validate_plan(plan)
        assert result.valid, [e.message for e in result.errors]


class TestHebrewContent:
    """
    The backend must pass Hebrew strings through validation untouched so the
    LLM's translated plan reaches the UI exactly as emitted. Pydantic is
    unicode-safe by default, but adding explicit coverage catches any future
    ASCII-only narrowing (e.g. someone slapping `re.ASCII` on an identifier
    regex).

    These tests do NOT hit the LLM — they assert that a plan object with
    Hebrew description / summary / sheet name survives Pydantic coercion
    and validator scrutiny.
    """

    def test_hebrew_description_and_summary(self):
        plan = make_plan(
            [
                {
                    "id": "step_1",
                    "description": "קרא את הטווח של הנתונים",
                    "action": "readRange",
                    "params": {"range": "גיליון1!A1:D100"},
                },
            ],
            summary="יצירת דוח מסכם",
            userRequest="צור לי טבלת ציר על עמודת מחלקה",
        )
        result = validate_plan(plan)
        # Hebrew sheet name in range should be accepted — Office.js supports
        # unicode sheet names and we should not narrow that server-side.
        assert result.valid, [e.message for e in result.errors]
        # Round-trip: the strings come back out unchanged.
        assert plan.summary == "יצירת דוח מסכם"
        assert plan.steps[0].description == "קרא את הטווח של הנתונים"

    def test_hebrew_sheet_name_in_multiple_params(self):
        """Hebrew sheet names in compound params (groupSum + createChart)."""
        plan = make_plan([
            {
                "id": "s1",
                "description": "סיכום לפי מחלקה",
                "action": "groupSum",
                "params": {
                    "dataRange": "מכירות!A1:E500",
                    "groupByColumn": 2,
                    "sumColumn": 5,
                    "outputRange": "סיכום!A1",
                },
            },
            {
                "id": "s2",
                "description": "גרף של הסיכום",
                "action": "createChart",
                "params": {
                    "dataRange": "סיכום!A1:B20",
                    "chartType": "bar",
                    "title": "מכירות לפי מחלקה",
                },
                "dependsOn": ["s1"],
            },
        ])
        result = validate_plan(plan)
        assert result.valid, [e.message for e in result.errors]

    def test_mixed_hebrew_english_request(self):
        """Real users switch languages mid-sentence — 'עשה pivot by department'."""
        plan = make_plan(
            [
                {
                    "id": "step_1",
                    "description": "צור pivot table",
                    "action": "createPivot",
                    "params": {"sourceRange": "Data!A1:F1000"},
                },
            ],
            userRequest="צור pivot by department and show top 10",
            summary="יצירת pivot לפי מחלקה",
        )
        result = validate_plan(plan)
        assert result.valid, [e.message for e in result.errors]
