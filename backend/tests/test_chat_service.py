"""
Tests for chat_service prompt construction.

Specifically verifies:
  - Workbook snapshot injection (headers, dtypes, anchor cell, offset tables)
  - All 51 StepAction values are constructible via ExecutionPlan
  - Retry prompt carries the previous failure reason
"""

from app.models.chat import ChatRequest, SheetSnapshot, WorkbookSnapshot
from app.models.plan import ExecutionPlan, StepAction, ACTION_PARAM_MODELS
from app.services.chat_service import _build_user_content, _build_retry_messages


class TestSnapshotInjection:
    def test_injects_headers_and_dtypes(self):
        req = ChatRequest(
            userMessage="sum the amounts",
            workbookSnapshot=WorkbookSnapshot(sheets=[
                SheetSnapshot(
                    sheetName="Data",
                    rowCount=100,
                    columnCount=3,
                    headers=["Region", "Product", "Amount"],
                    sampleRows=[["North", "Widget", 1200]],
                    dtypes=["text", "text", "number"],
                    anchorCell="A1",
                    usedRangeAddress="Data!A1:C100",
                ),
            ]),
        )
        content = _build_user_content(req)
        assert "Workbook data snapshot" in content
        assert "Region [text]" in content
        assert "Amount [number]" in content
        assert "100 rows × 3 cols" in content
        assert "data starts at A1" in content

    def test_offset_table_reports_absolute_rows(self):
        """When data starts at C5, sample rows should be labeled as sheet rows 6, 7..."""
        req = ChatRequest(
            userMessage="sum amounts",
            workbookSnapshot=WorkbookSnapshot(sheets=[
                SheetSnapshot(
                    sheetName="Report",
                    rowCount=36,
                    columnCount=4,
                    headers=["A", "B", "C", "D"],
                    sampleRows=[[1, 2, 3, 4], [5, 6, 7, 8]],
                    dtypes=["number", "number", "number", "number"],
                    anchorCell="C5",
                    usedRangeAddress="Report!C5:F40",
                ),
            ]),
        )
        content = _build_user_content(req)
        assert "data starts at C5" in content
        assert "headers on row 5" in content
        assert "Report!C5:F40" in content
        # First sample row is sheet-absolute row 6 (header is 5, +1)
        assert "row 6:" in content
        assert "row 7:" in content
        # Must NOT contain the old index-relative labels
        assert "row 2:" not in content

    def test_empty_sheet_noted(self):
        req = ChatRequest(
            userMessage="do something",
            workbookSnapshot=WorkbookSnapshot(sheets=[
                SheetSnapshot(
                    sheetName="Blank", rowCount=0, columnCount=0,
                    headers=[], sampleRows=[], dtypes=[],
                    anchorCell="A1", usedRangeAddress="",
                ),
            ]),
        )
        content = _build_user_content(req)
        assert "Blank: (empty)" in content

    def test_no_snapshot_no_section(self):
        req = ChatRequest(userMessage="hi")
        content = _build_user_content(req)
        assert "Workbook data snapshot" not in content


class TestAllActionsAccepted:
    def test_all_76_actions_have_param_models(self):
        assert len(list(StepAction)) == 76
        assert len(ACTION_PARAM_MODELS) == 76
        for action in StepAction:
            assert action in ACTION_PARAM_MODELS, f"{action} missing from ACTION_PARAM_MODELS"

    def test_new_capabilities_construct_plan(self):
        """Smoke test: each of the 17 newly-added actions must construct."""
        specimens = {
            "pageLayout": {"orientation": "landscape"},
            "insertPicture": {"imageBase64": "iVBORw"},
            "insertShape": {"shapeType": "rectangle", "left": 0, "top": 0, "width": 100, "height": 50},
            "insertTextBox": {"text": "hi", "left": 0, "top": 0, "width": 100, "height": 50},
            "addSlicer": {"sourceType": "table", "sourceName": "T1", "sourceField": "Region"},
            "splitColumn": {"sourceRange": "A:A", "delimiter": " ", "outputStartColumn": "B"},
            "unpivot": {"sourceRange": "A1:E10", "idColumns": 1, "outputRange": "G1"},
            "crossTabulate": {"sourceRange": "A1:C100", "rowField": 1, "columnField": 2, "valueField": 3, "aggregation": "sum", "outputRange": "E1"},
            "bulkFormula": {"formula": "=A2*2", "outputRange": "B:B", "dataRange": "A:A"},
            "compareSheets": {"rangeA": "Sheet1!A:Z", "rangeB": "Sheet2!A:Z"},
            "consolidateRanges": {"sourceRanges": ["Q1!A:C", "Q2!A:C"], "outputRange": "Summary!A1"},
            "extractPattern": {"sourceRange": "A:A", "pattern": "email", "outputRange": "B:B"},
            "categorize": {"sourceRange": "A:A", "outputRange": "B:B", "rules": [{"operator": "contains", "value": "VIP", "label": "Top"}]},
            "fillBlanks": {"range": "A:A"},
            "subtotals": {"dataRange": "A1:D100", "groupByColumn": 1, "subtotalColumns": [3, 4]},
            "transpose": {"sourceRange": "A1:C10", "outputRange": "E1"},
            "namedRange": {"operation": "create", "name": "SalesData", "range": "A1:D100"},
        }
        for action_name, params in specimens.items():
            plan = ExecutionPlan(
                planId="x", createdAt="2026", userRequest="u", summary="s",
                confidence=0.9, preserveFormatting=True,
                steps=[{"id": "s1", "description": "d", "action": action_name, "params": params}],
            )
            assert plan.steps[0].action.value == action_name


class TestRetryPromptFeedback:
    def test_retry_prompt_includes_failure_reason(self):
        req = ChatRequest(userMessage="test")
        msgs = _build_retry_messages(req, None, failure_reason="invalid action 'splitColum'")
        system = msgs[0]["content"]
        assert "previous attempt failed" in system
        assert "splitColum" in system

    def test_retry_prompt_without_reason(self):
        req = ChatRequest(userMessage="test")
        msgs = _build_retry_messages(req, None, failure_reason=None)
        system = msgs[0]["content"]
        assert "previous attempt failed" not in system
