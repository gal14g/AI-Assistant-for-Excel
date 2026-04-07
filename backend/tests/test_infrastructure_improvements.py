"""
Tests for the three infrastructure improvements:
  1. Larger snapshots — more sample rows sent to / shown by the LLM
  2. Step output binding — {{step_N.field}} syntax in system prompt
  3. Multi-turn plan refinement — execution context injection on failure

Also covers:
  - All 76 actions are registered
  - Hebrew values, merged cells, and date formats in snapshot injection
  - Realistic multi-step pipeline plans with Hebrew data
  - Failed-step refinement context construction
"""

from __future__ import annotations

import re
import pytest
from pydantic import ValidationError

from app.models.chat import (
    ChatRequest,
    SheetSnapshot,
    WorkbookSnapshot,
    ExecutionContext,
    StepExecutionResult,
)
from app.models.plan import ExecutionPlan, StepAction, ACTION_PARAM_MODELS
from app.services.chat_service import _build_user_content
from app.services.validator import validate_plan


# ═══════════════════════════════════════════════════════════════════════════════
# 1. LARGER SNAPSHOTS
# ═══════════════════════════════════════════════════════════════════════════════


class TestLargerSnapshots:
    """Verify that the system accepts and displays more sample rows than the
    old limit of 5 / max_length=10."""

    def test_accepts_20_sample_rows(self):
        """Backend model should accept up to 50 sample rows (was 10)."""
        rows = [[f"val_{i}", i, True] for i in range(20)]
        snap = SheetSnapshot(
            sheetName="Data",
            rowCount=500,
            columnCount=3,
            headers=["Name", "Amount", "Active"],
            sampleRows=rows,
            dtypes=["text", "number", "boolean"],
            anchorCell="A1",
            usedRangeAddress="Data!A1:C500",
        )
        assert len(snap.sampleRows) == 20

    def test_accepts_50_sample_rows(self):
        """Backend model max_length is now 50."""
        rows = [[f"item_{i}"] for i in range(50)]
        snap = SheetSnapshot(
            sheetName="Big",
            rowCount=1000,
            columnCount=1,
            headers=["Item"],
            sampleRows=rows,
            dtypes=["text"],
            anchorCell="A1",
            usedRangeAddress="Big!A1:A1000",
        )
        assert len(snap.sampleRows) == 50

    def test_shows_up_to_10_rows_in_prompt(self):
        """The LLM prompt should show up to 10 sample rows (was 3)."""
        rows = [[f"name_{i}", i * 100] for i in range(20)]
        req = ChatRequest(
            userMessage="analyze data",
            workbookSnapshot=WorkbookSnapshot(sheets=[
                SheetSnapshot(
                    sheetName="Sales",
                    rowCount=200,
                    columnCount=2,
                    headers=["Name", "Amount"],
                    sampleRows=rows,
                    dtypes=["text", "number"],
                    anchorCell="A1",
                    usedRangeAddress="Sales!A1:B200",
                ),
            ]),
        )
        content = _build_user_content(req)
        # Should show rows 2 through 11 (header is row 1, then 10 sample rows)
        assert "row 11:" in content  # 10th sample row
        # Should NOT show row 12 (11th sample row) — limited to 10
        assert "row 12:" not in content

    def test_snapshot_with_hebrew_headers_and_values(self):
        """Hebrew column names and cell values must survive snapshot injection."""
        req = ChatRequest(
            userMessage="סכם את העמודה",
            workbookSnapshot=WorkbookSnapshot(sheets=[
                SheetSnapshot(
                    sheetName="נתונים",
                    rowCount=50,
                    columnCount=3,
                    headers=["שם", "עיר", "סכום"],
                    sampleRows=[
                        ["דוד כהן", "תל אביב", 1500],
                        ["שרה לוי", "ירושלים", 2300],
                        ["משה ישראלי", "חיפה", 800],
                    ],
                    dtypes=["text", "text", "number"],
                    anchorCell="A1",
                    usedRangeAddress="נתונים!A1:C50",
                ),
            ]),
        )
        content = _build_user_content(req)
        assert "שם [text]" in content
        assert "סכום [number]" in content
        assert "נתונים" in content
        assert "דוד כהן" in content
        assert "תל אביב" in content

    def test_snapshot_with_dates_and_mixed_types(self):
        """Date columns and mixed-type data should be reported correctly."""
        req = ChatRequest(
            userMessage="normalize dates",
            workbookSnapshot=WorkbookSnapshot(sheets=[
                SheetSnapshot(
                    sheetName="Records",
                    rowCount=100,
                    columnCount=4,
                    headers=["ID", "Date", "Amount", "Notes"],
                    sampleRows=[
                        [1, "15/03/2026", 500, "first entry"],
                        [2, "2026-04-01", 1200, None],
                        [3, "01-Jan-25", 300, "old format"],
                        [4, 45000, 750, "Excel serial"],
                        [5, "05/13/2026", 900, "US format"],
                    ],
                    dtypes=["number", "date", "number", "mixed"],
                    anchorCell="A1",
                    usedRangeAddress="Records!A1:D100",
                ),
            ]),
        )
        content = _build_user_content(req)
        assert "Date [date]" in content
        assert "Notes [mixed]" in content
        assert "15/03/2026" in content

    def test_snapshot_with_offset_table_and_merged_cells(self):
        """Data starting at C5 (not A1) — common with merged-cell reports."""
        req = ChatRequest(
            userMessage="clean up this data",
            workbookSnapshot=WorkbookSnapshot(sheets=[
                SheetSnapshot(
                    sheetName="דוח מכירות",
                    rowCount=200,
                    columnCount=5,
                    headers=["תאריך", "מוצר", "כמות", "מחיר", "סה\"כ"],
                    sampleRows=[
                        ["01/04/2026", "מחשב נייד", 5, 3500, 17500],
                        ["02/04/2026", "מסך", 10, 1200, 12000],
                        ["03/04/2026", "מקלדת", 25, 150, 3750],
                    ],
                    dtypes=["date", "text", "number", "number", "number"],
                    anchorCell="C5",
                    usedRangeAddress="דוח מכירות!C5:G204",
                ),
            ]),
        )
        content = _build_user_content(req)
        assert "data starts at C5" in content
        assert "headers on row 5" in content
        # First data row should be row 6 (C5 is header)
        assert "row 6:" in content
        assert "מחשב נייד" in content
        assert "דוח מכירות!C5:G204" in content


# ═══════════════════════════════════════════════════════════════════════════════
# 2. STEP OUTPUT BINDING — system prompt teaches the syntax
# ═══════════════════════════════════════════════════════════════════════════════


class TestStepOutputBindingPrompt:
    """Verify the LLM system prompt includes step output binding documentation."""

    def _get_system_prompt(self) -> str:
        from app.services.chat_service import _build_chat_system_prompt
        return _build_chat_system_prompt()

    def test_prompt_includes_binding_syntax(self):
        prompt = self._get_system_prompt()
        assert "{{step_" in prompt
        assert "STEP OUTPUT BINDING" in prompt

    def test_prompt_explains_outputrange_binding(self):
        prompt = self._get_system_prompt()
        assert "outputRange" in prompt

    def test_prompt_explains_sheetname_binding(self):
        prompt = self._get_system_prompt()
        assert "sheetName" in prompt

    def test_prompt_mentions_dependson_requirement(self):
        """Binding only works if dependsOn is set correctly."""
        prompt = self._get_system_prompt()
        assert "dependsOn" in prompt


# ═══════════════════════════════════════════════════════════════════════════════
# 3. MULTI-TURN PLAN REFINEMENT
# ═══════════════════════════════════════════════════════════════════════════════


class TestExecutionContextModel:
    """Validate the ExecutionContext and StepExecutionResult Pydantic models."""

    def test_constructs_valid_execution_context(self):
        ctx = ExecutionContext(
            originalPlanId="plan-123",
            originalUserRequest="match customers",
            stepResults=[
                StepExecutionResult(stepId="step_1", status="success", message="Added sheet"),
                StepExecutionResult(
                    stepId="step_2", status="error",
                    message="Failed to write formula",
                    error="Formula error: #REF!",
                ),
            ],
            failedStepId="step_2",
            failedStepAction="writeFormula",
            failedStepError="Formula error: #REF!",
        )
        assert ctx.originalPlanId == "plan-123"
        assert len(ctx.stepResults) == 2
        assert ctx.failedStepId == "step_2"

    def test_step_result_status_validation(self):
        """Status must be one of the allowed literals."""
        sr = StepExecutionResult(stepId="s1", status="success", message="ok")
        assert sr.status == "success"

        sr2 = StepExecutionResult(stepId="s2", status="error", message="fail", error="boom")
        assert sr2.status == "error"

    def test_chat_request_accepts_execution_context(self):
        """ChatRequest should accept the optional executionContext field."""
        ctx = ExecutionContext(
            originalPlanId="p1",
            originalUserRequest="test",
            stepResults=[],
        )
        req = ChatRequest(
            userMessage="fix the plan",
            executionContext=ctx,
        )
        assert req.executionContext is not None
        assert req.executionContext.originalPlanId == "p1"


class TestRefinementPromptInjection:
    """Verify that execution context is injected into the LLM prompt."""

    def test_refinement_context_injected_into_prompt(self):
        ctx = ExecutionContext(
            originalPlanId="plan-abc",
            originalUserRequest="create dashboard from sales data",
            stepResults=[
                StepExecutionResult(stepId="step_1", status="success", message="Added sheet 'Dashboard'"),
                StepExecutionResult(stepId="step_2", status="success", message="Wrote 5 KPI formulas"),
                StepExecutionResult(
                    stepId="step_3", status="error",
                    message="Failed to create chart",
                    error="Invalid dataRange: Dashboard!A1:B0",
                ),
            ],
            failedStepId="step_3",
            failedStepAction="createChart",
            failedStepError="Invalid dataRange: Dashboard!A1:B0",
        )
        req = ChatRequest(
            userMessage="fix the failed step",
            executionContext=ctx,
        )
        content = _build_user_content(req)

        assert "PLAN REFINEMENT" in content
        assert "plan-abc" in content
        assert "create dashboard from sales data" in content
        # Successful steps shown
        assert "step_1" in content
        assert "success" in content
        # Failed step details shown
        assert "step_3" in content
        assert "createChart" in content
        assert "Invalid dataRange" in content
        # Instruction to fix
        assert "CORRECTED plan" in content
        assert "do NOT re-run" in content

    def test_no_refinement_section_without_context(self):
        req = ChatRequest(userMessage="normal message")
        content = _build_user_content(req)
        assert "PLAN REFINEMENT" not in content

    def test_refinement_with_hebrew_error_messages(self):
        """Hebrew error messages should be preserved in the refinement context."""
        ctx = ExecutionContext(
            originalPlanId="plan-heb",
            originalUserRequest="התאם נתונים מגיליון לקוחות",
            stepResults=[
                StepExecutionResult(
                    stepId="step_1", status="error",
                    message="שגיאה בטווח",
                    error="Range not found: גיליון1!A:A",
                ),
            ],
            failedStepId="step_1",
            failedStepAction="matchRecords",
            failedStepError="Range not found: גיליון1!A:A",
        )
        req = ChatRequest(
            userMessage="תקן את התוכנית",
            executionContext=ctx,
        )
        content = _build_user_content(req)
        assert "גיליון1!A:A" in content

    def test_refinement_with_multi_step_partial_success(self):
        """A 5-step pipeline where step 4 fails — first 3 succeeded."""
        ctx = ExecutionContext(
            originalPlanId="plan-multi",
            originalUserRequest="clean, match, format, chart, protect",
            stepResults=[
                StepExecutionResult(stepId="step_1", status="success", message="Cleaned 50 cells"),
                StepExecutionResult(stepId="step_2", status="success", message="Matched 45/50 rows"),
                StepExecutionResult(stepId="step_3", status="success", message="Applied formatting"),
                StepExecutionResult(
                    stepId="step_4", status="error",
                    message="Chart creation failed",
                    error="Source range is empty after filtering",
                ),
                StepExecutionResult(stepId="step_5", status="skipped", message="Skipped: depends on step_4"),
            ],
            failedStepId="step_4",
            failedStepAction="createChart",
            failedStepError="Source range is empty after filtering",
        )
        req = ChatRequest(
            userMessage="fix step 4",
            executionContext=ctx,
        )
        content = _build_user_content(req)

        # All 5 steps should appear in the context
        for i in range(1, 6):
            assert f"step_{i}" in content
        # The failed step error detail
        assert "Source range is empty after filtering" in content
        assert "step_5" in content  # skipped step


# ═══════════════════════════════════════════════════════════════════════════════
# 4. ALL 76 ACTIONS REGISTERED
# ═══════════════════════════════════════════════════════════════════════════════


class TestAll76ActionsRegistered:
    """Ensure all 76 StepAction enum values have param models."""

    def test_76_actions_in_enum(self):
        assert len(list(StepAction)) == 76

    def test_76_param_models(self):
        assert len(ACTION_PARAM_MODELS) == 76
        for action in StepAction:
            assert action in ACTION_PARAM_MODELS, f"{action} missing from ACTION_PARAM_MODELS"

    def test_new_25_actions_construct_plans(self):
        """Smoke test: each of the 25 new actions can construct a valid plan."""
        specimens = {
            "fuzzyMatch": {"lookupRange": "A:A", "sourceRange": "B:B", "outputRange": "C:C", "threshold": 0.7},
            "deleteRowsByCondition": {"range": "A1:D100", "column": 1, "condition": "blank"},
            "splitByGroup": {"dataRange": "A1:C50", "groupByColumn": 1},
            "lookupAll": {"lookupRange": "A:A", "sourceRange": "B:B", "returnColumn": 2, "outputRange": "C:C"},
            "regexReplace": {"range": "A:A", "pattern": r"\d+", "replacement": "#"},
            "coerceDataType": {"range": "A:A", "targetType": "number"},
            "normalizeDates": {"range": "A:A", "outputFormat": "yyyy-mm-dd"},
            "deduplicateAdvanced": {"range": "A1:D100", "keyColumns": [1, 2], "keepStrategy": "first"},
            "joinSheets": {"leftRange": "Sheet1!A1:C50", "rightRange": "Sheet2!A1:B50",
                           "leftKeyColumn": 1, "rightKeyColumn": 1, "joinType": "inner", "outputRange": "D1"},
            "frequencyDistribution": {"sourceRange": "A:A", "outputRange": "D1"},
            "runningTotal": {"sourceRange": "A:A", "outputRange": "B:B"},
            "rankColumn": {"sourceRange": "A:A", "outputRange": "B:B"},
            "topN": {"dataRange": "A1:D100", "valueColumn": 3, "n": 10, "outputRange": "F1"},
            "percentOfTotal": {"sourceRange": "A:A", "outputRange": "B:B"},
            "growthRate": {"sourceRange": "A:A", "outputRange": "B:B"},
            "consolidateAllSheets": {},
            "cloneSheetStructure": {"sourceSheet": "Template", "newSheetName": "Copy"},
            "addReportHeader": {"title": "Q1 Report"},
            "alternatingRowFormat": {"range": "A1:D50"},
            "quickFormat": {"range": "A1:D50"},
            "refreshPivot": {},
            "pivotCalculatedField": {"pivotName": "PivotTable1", "fieldName": "Margin", "formula": "=Revenue-Cost"},
            "addDropdownControl": {"cell": "A1", "listSource": "Options!A:A"},
            "conditionalFormula": {"range": "A:A", "conditionColumn": 1, "condition": "blank",
                                   "trueFormula": "=\"empty\"", "falseFormula": "=\"filled\"", "outputRange": "B:B"},
            "spillFormula": {"cell": "A1", "formula": "=UNIQUE(Sheet1!A:A)"},
        }
        for action_name, params in specimens.items():
            plan = ExecutionPlan(
                planId="test", createdAt="2026-04-07", userRequest="test", summary="test",
                confidence=0.9, preserveFormatting=True,
                steps=[{"id": "s1", "description": "test", "action": action_name, "params": params}],
            )
            assert plan.steps[0].action.value == action_name


# ═══════════════════════════════════════════════════════════════════════════════
# 5. MULTI-STEP PIPELINE VALIDATION — realistic complex plans
# ═══════════════════════════════════════════════════════════════════════════════


class TestMultiStepPipelineValidation:
    """Validate that realistic multi-step plans with dependencies pass validation."""

    def test_5_step_dashboard_pipeline(self):
        """addSheet → writeValues → writeFormula → createChart → autoFitColumns"""
        plan = ExecutionPlan(
            planId="dashboard-1",
            createdAt="2026-04-07T00:00:00Z",
            userRequest="create a sales dashboard",
            summary="Create dashboard sheet with KPIs and chart",
            confidence=0.92,
            preserveFormatting=False,
            steps=[
                {"id": "step_1", "description": "Create dashboard sheet",
                 "action": "addSheet", "params": {"sheetName": "Dashboard"}},
                {"id": "step_2", "description": "Write KPI labels",
                 "action": "writeValues",
                 "params": {"range": "Dashboard!A1:B3", "values": [["Metric", "Value"], ["Total Sales", ""], ["Avg Sale", ""]]},
                 "dependsOn": ["step_1"]},
                {"id": "step_3", "description": "Write SUMIF formula",
                 "action": "writeFormula",
                 "params": {"cell": "Dashboard!B2", "formula": "=SUM(Data!C:C)"},
                 "dependsOn": ["step_2"]},
                {"id": "step_4", "description": "Create bar chart",
                 "action": "createChart",
                 "params": {"dataRange": "Dashboard!A1:B3", "chartType": "bar", "title": "Sales KPIs"},
                 "dependsOn": ["step_3"]},
                {"id": "step_5", "description": "Auto-fit columns",
                 "action": "autoFitColumns",
                 "params": {"sheetName": "Dashboard"},
                 "dependsOn": ["step_4"]},
            ],
        )
        result = validate_plan(plan)
        assert result.valid, f"Errors: {[e.message for e in result.errors]}"

    def test_3_step_clean_match_format_pipeline(self):
        """cleanupText → matchRecords → addConditionalFormat"""
        plan = ExecutionPlan(
            planId="match-1",
            createdAt="2026-04-07T00:00:00Z",
            userRequest="clean names and match to other sheet",
            summary="Clean text, match records, highlight mismatches",
            confidence=0.88,
            preserveFormatting=False,
            steps=[
                {"id": "step_1", "description": "Clean names",
                 "action": "cleanupText",
                 "params": {"range": "Sheet1!A:A", "operations": ["trim", "properCase"]}},
                {"id": "step_2", "description": "Match records",
                 "action": "matchRecords",
                 "params": {
                     "lookupRange": "Sheet1!A:A", "sourceRange": "Sheet2!A:A",
                     "matchType": "contains", "outputRange": "Sheet1!D:D", "writeValue": "v",
                 },
                 "dependsOn": ["step_1"]},
                {"id": "step_3", "description": "Highlight unmatched",
                 "action": "addConditionalFormat",
                 "params": {
                     "range": "Sheet1!A2:D1000",
                     "ruleType": "formula", "formula": "=$D2=\"\"",
                     "format": {"fillColor": "#FFCCCC"},
                 },
                 "dependsOn": ["step_2"]},
            ],
        )
        result = validate_plan(plan)
        assert result.valid, f"Errors: {[e.message for e in result.errors]}"

    def test_hebrew_pipeline_with_dates(self):
        """Hebrew sheet names + date normalization + frequency distribution"""
        plan = ExecutionPlan(
            planId="hebrew-1",
            createdAt="2026-04-07T00:00:00Z",
            userRequest="נקה תאריכים וצור טבלת תדירות",
            summary="Normalize dates then create frequency table",
            confidence=0.9,
            preserveFormatting=True,
            steps=[
                {"id": "step_1", "description": "Normalize dates",
                 "action": "normalizeDates",
                 "params": {"range": "נתונים!B:B", "outputFormat": "dd/mm/yyyy"}},
                {"id": "step_2", "description": "Frequency distribution",
                 "action": "frequencyDistribution",
                 "params": {"sourceRange": "נתונים!C:C", "outputRange": "סיכום!A1"},
                 "dependsOn": ["step_1"]},
            ],
        )
        result = validate_plan(plan)
        assert result.valid, f"Errors: {[e.message for e in result.errors]}"

    def test_4_step_with_binding_syntax_in_params(self):
        """Plans with {{step_N.field}} binding tokens should still validate
        (the validator doesn't resolve bindings — that's the executor's job)."""
        plan = ExecutionPlan(
            planId="binding-1",
            createdAt="2026-04-07T00:00:00Z",
            userRequest="create sheet and chart its data",
            summary="Add sheet, write data, chart it",
            confidence=0.9,
            preserveFormatting=False,
            steps=[
                {"id": "step_1", "description": "Add sheet",
                 "action": "addSheet", "params": {"sheetName": "Report"}},
                {"id": "step_2", "description": "Write data",
                 "action": "writeValues",
                 "params": {"range": "Report!A1:B5", "values": [["Month", "Sales"], ["Jan", 100], ["Feb", 200], ["Mar", 150], ["Apr", 300]]},
                 "dependsOn": ["step_1"]},
                {"id": "step_3", "description": "Create chart from step_2 output",
                 "action": "createChart",
                 "params": {"dataRange": "{{step_2.outputRange}}", "chartType": "line", "title": "Trend"},
                 "dependsOn": ["step_2"]},
                {"id": "step_4", "description": "Auto-fit",
                 "action": "autoFitColumns", "params": {"sheetName": "Report"},
                 "dependsOn": ["step_3"]},
            ],
        )
        result = validate_plan(plan)
        # Binding tokens are strings — validator should pass them through
        assert result.valid, f"Errors: {[e.message for e in result.errors]}"


# ═══════════════════════════════════════════════════════════════════════════════
# 6. FAILED-STEP-THEN-FIX SCENARIO (end-to-end model test)
# ═══════════════════════════════════════════════════════════════════════════════


class TestFailedStepRefinementScenario:
    """Simulate a full fail-then-refine flow at the model layer."""

    def test_original_plan_then_refinement_request(self):
        """Build an original plan, simulate partial execution failure,
        then verify a refinement ChatRequest is correctly constructed."""

        # 1. Original plan: 4 steps
        original_plan = ExecutionPlan(
            planId="plan-orig-001",
            createdAt="2026-04-07T10:00:00Z",
            userRequest="match customers to orders and create summary",
            summary="Match, aggregate, format",
            confidence=0.9,
            preserveFormatting=True,
            steps=[
                {"id": "step_1", "description": "Clean customer names",
                 "action": "cleanupText",
                 "params": {"range": "Customers!B:B", "operations": ["trim", "properCase"]}},
                {"id": "step_2", "description": "Match to orders",
                 "action": "matchRecords",
                 "params": {
                     "lookupRange": "Customers!B:B", "sourceRange": "Orders!A:A",
                     "matchType": "exact", "outputRange": "Customers!E:E",
                 },
                 "dependsOn": ["step_1"]},
                {"id": "step_3", "description": "Count matches per city",
                 "action": "groupSum",
                 "params": {"dataRange": "Customers!A1:E100", "groupByColumn": 3, "sumColumn": 5, "outputRange": "Summary!A1"},
                 "dependsOn": ["step_2"]},
                {"id": "step_4", "description": "Create pie chart",
                 "action": "createChart",
                 "params": {"dataRange": "Summary!A1:B10", "chartType": "pie", "title": "Matches by City"},
                 "dependsOn": ["step_3"]},
            ],
        )

        # Validate original plan
        result = validate_plan(original_plan)
        assert result.valid

        # 2. Simulate execution: step_1 and step_2 succeed, step_3 fails
        execution_results = [
            StepExecutionResult(stepId="step_1", status="success", message="Cleaned 100 cells"),
            StepExecutionResult(stepId="step_2", status="success", message="Matched 85/100 records"),
            StepExecutionResult(
                stepId="step_3", status="error",
                message="GroupSum failed: column 5 has no numeric values",
                error="Column 5 contains text values, expected numbers",
            ),
            StepExecutionResult(stepId="step_4", status="skipped", message="Skipped: depends on step_3"),
        ]

        # 3. Build refinement request
        ctx = ExecutionContext(
            originalPlanId=original_plan.planId,
            originalUserRequest=original_plan.userRequest,
            stepResults=execution_results,
            failedStepId="step_3",
            failedStepAction="groupSum",
            failedStepError="Column 5 contains text values, expected numbers",
        )

        refinement_req = ChatRequest(
            userMessage="Fix the groupSum step — use COUNTIF instead of SUM",
            executionContext=ctx,
            workbookSnapshot=WorkbookSnapshot(sheets=[
                SheetSnapshot(
                    sheetName="Customers",
                    rowCount=100,
                    columnCount=5,
                    headers=["ID", "Name", "City", "Status", "MatchResult"],
                    sampleRows=[
                        ["C001", "Alice Johnson", "New York", "Active", "v"],
                        ["C002", "Bob Smith", "Boston", "Active", "v"],
                        ["C003", "Carol White", "Chicago", "Inactive", ""],
                    ],
                    dtypes=["text", "text", "text", "text", "text"],
                    anchorCell="A1",
                    usedRangeAddress="Customers!A1:E100",
                ),
            ]),
        )

        # 4. Verify the refinement prompt includes all relevant context
        content = _build_user_content(refinement_req)

        # Execution context
        assert "PLAN REFINEMENT" in content
        assert "plan-orig-001" in content
        assert "step_1" in content
        assert "Cleaned 100 cells" in content
        assert "step_3" in content
        assert "groupSum" in content
        assert "Column 5 contains text values" in content

        # Snapshot still present for grounding
        assert "Customers" in content
        assert "MatchResult" in content
        assert "Alice Johnson" in content

    def test_refinement_preserves_execution_order(self):
        """Step results in the refinement context must be in execution order."""
        results = [
            StepExecutionResult(stepId="step_1", status="success", message="ok"),
            StepExecutionResult(stepId="step_2", status="success", message="ok"),
            StepExecutionResult(stepId="step_3", status="error", message="fail", error="boom"),
        ]
        ctx = ExecutionContext(
            originalPlanId="p1",
            originalUserRequest="test",
            stepResults=results,
            failedStepId="step_3",
            failedStepAction="writeFormula",
            failedStepError="boom",
        )
        req = ChatRequest(userMessage="fix", executionContext=ctx)
        content = _build_user_content(req)

        # step_1 should appear before step_2 which should appear before step_3
        pos1 = content.index("step_1")
        pos2 = content.index("step_2")
        pos3 = content.index("step_3")
        assert pos1 < pos2 < pos3

    def test_corrected_plan_validates(self):
        """A corrected plan (starting from step_3) should still validate."""
        corrected = ExecutionPlan(
            planId="plan-fix-001",
            createdAt="2026-04-07T10:05:00Z",
            userRequest="fix the groupSum step",
            summary="Use COUNTIF instead of SUM for text column",
            confidence=0.85,
            preserveFormatting=True,
            steps=[
                {"id": "step_3", "description": "Count matches per city with COUNTIF",
                 "action": "writeFormula",
                 "params": {"cell": "Summary!B1", "formula": "=COUNTIF(Customers!E:E,\"v\")", "fillDown": 10}},
                {"id": "step_4", "description": "Create pie chart",
                 "action": "createChart",
                 "params": {"dataRange": "Summary!A1:B10", "chartType": "pie", "title": "Matches by City"},
                 "dependsOn": ["step_3"]},
            ],
        )
        result = validate_plan(corrected)
        assert result.valid, f"Errors: {[e.message for e in result.errors]}"
