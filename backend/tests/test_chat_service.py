"""
Tests for chat_service prompt construction.

Specifically verifies:
  - Workbook snapshot injection (headers, dtypes, anchor cell, offset tables)
  - All 51 StepAction values are constructible via ExecutionPlan
  - Retry prompt carries the previous failure reason
"""

from app.models.chat import ChatRequest, SheetSnapshot, WorkbookSnapshot
from app.models.plan import ExecutionPlan, StepAction, ACTION_PARAM_MODELS
from app.services.chat_service import (
    _build_user_content,
    _build_retry_messages,
    _validate_step_actions,
    _parse_response,
    _normalize_param_keys,
    chat_stream,
)
from app.services.validator import validate_plan
import pytest


def _make_plan(action: str, params: dict) -> ExecutionPlan:
    return ExecutionPlan(
        planId="x", createdAt="2026", userRequest="u", summary="s",
        confidence=0.9, preserveFormatting=True,
        steps=[{"id": "step_1", "description": "d", "action": action, "params": params}],
    )


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
    def test_all_actions_have_param_models(self):
        # Number grows as new capabilities land — the invariant that matters
        # is that every enum member has a param model and vice versa.
        assert len(ACTION_PARAM_MODELS) == len(list(StepAction))
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


class TestStepActionValidation:
    """Regression tests for hallucinated action names like 'manualStep'.

    Pydantic's enum validator emits a 76-item dump that the retry path
    can't make sense of. _validate_step_actions intercepts these earlier
    and produces a short, LLM-actionable failure_reason.
    """

    def test_valid_actions_pass(self):
        _validate_step_actions({"steps": [
            {"action": "writeValues"},
            {"action": "matchRecords"},
            {"action": "createChart"},
        ]})

    def test_empty_steps_pass(self):
        _validate_step_actions({"steps": []})
        _validate_step_actions({})

    def test_hallucinated_action_raises_actionable_error(self):
        with pytest.raises(ValueError) as exc:
            _validate_step_actions({"steps": [{"action": "manualStep"}]})
        msg = str(exc.value)
        assert "manualStep" in msg
        assert "Invalid action name" in msg
        assert "never invent" in msg.lower()

    def test_multiple_bad_actions_all_reported(self):
        with pytest.raises(ValueError) as exc:
            _validate_step_actions({"steps": [
                {"action": "writeValues"},
                {"action": "doIt"},
                {"action": "magicMove"},
            ]})
        msg = str(exc.value)
        assert "doIt" in msg
        assert "magicMove" in msg

    def test_close_match_suggestions_offered(self):
        # "writevalue" should suggest writeValues / writeValue
        with pytest.raises(ValueError) as exc:
            _validate_step_actions({"steps": [{"action": "writevalue"}]})
        msg = str(exc.value)
        assert "writeValues" in msg or "writeValue" in msg

    def test_non_string_action_skipped(self):
        # Non-string actions are left for Pydantic's normal validation path
        _validate_step_actions({"steps": [{"action": None}, {"action": 42}]})


class TestStepBindingValidation:
    """{{step_N.field}} bindings must reference real, earlier steps.

    Bindings to non-existent steps would otherwise reach the executor as
    literal "{{step_99.outputRange}}" strings and produce a cryptic
    Office.js "invalid range" error.
    """

    def test_valid_binding_passes(self):
        _validate_step_actions({"steps": [
            {"id": "step_1", "action": "addSheet", "params": {"sheetName": "X"}},
            {"id": "step_2", "action": "writeValues", "params": {
                "range": "{{step_1.sheetName}}!A1", "values": [["a"]],
            }},
        ]})

    def test_binding_to_missing_step_raises(self):
        with pytest.raises(ValueError) as exc:
            _validate_step_actions({"steps": [
                {"id": "step_1", "action": "writeValues", "params": {
                    "range": "{{step_99.outputRange}}", "values": [["a"]],
                }},
            ]})
        msg = str(exc.value)
        assert "step_99" in msg
        assert "binding" in msg.lower()

    def test_self_referencing_binding_raises(self):
        with pytest.raises(ValueError) as exc:
            _validate_step_actions({"steps": [
                {"id": "step_1", "action": "writeValues", "params": {
                    "range": "{{step_1.outputRange}}", "values": [["a"]],
                }},
            ]})
        assert "self-reference" in str(exc.value)

    def test_no_bindings_no_check(self):
        # Plain plans without bindings shouldn't even hit the binding validator
        _validate_step_actions({"steps": [
            {"id": "step_1", "action": "writeValues", "params": {"range": "A1", "values": [["a"]]}},
        ]})

    def test_binding_wrong_field_for_action_raises(self):
        """createChart produces 'chartName', not 'outputRange'. Should fail."""
        with pytest.raises(ValueError, match="chartName"):
            _validate_step_actions({"steps": [
                {"id": "step_1", "action": "createChart", "params": {"dataRange": "A1:B10", "chartType": "line"}},
                {"id": "step_2", "action": "writeValues", "params": {
                    "range": "{{step_1.outputRange}}", "values": [["a"]],
                }},
            ]})

    def test_binding_correct_field_accepted(self):
        """addSheet produces 'sheetName'. Binding to it should pass."""
        _validate_step_actions({"steps": [
            {"id": "step_1", "action": "addSheet", "params": {"sheetName": "NewSheet"}},
            {"id": "step_2", "action": "writeValues", "params": {
                "range": "{{step_1.sheetName}}!A1", "values": [["a"]],
            }},
        ]})

    def test_binding_to_no_output_action_raises(self):
        """freezePanes produces no outputs. Any binding to it should fail."""
        with pytest.raises(ValueError, match="none"):
            _validate_step_actions({"steps": [
                {"id": "step_1", "action": "freezePanes", "params": {"range": "A2"}},
                {"id": "step_2", "action": "writeValues", "params": {
                    "range": "{{step_1.range}}", "values": [["a"]],
                }},
            ]})

    def test_dependson_auto_inferred_from_bindings(self):
        """If step_2 binds to step_1 but has no dependsOn, it should be auto-added."""
        steps_data = {"steps": [
            {"id": "step_1", "action": "readRange", "params": {"range": "A1:B10"}},
            {"id": "step_2", "action": "createChart", "params": {
                "dataRange": "{{step_1.outputRange}}", "chartType": "line",
            }},
        ]}
        _validate_step_actions(steps_data)
        # dependsOn should have been auto-populated
        assert "step_1" in steps_data["steps"][1].get("dependsOn", [])


class TestParseResponseTypeGuards:
    """Defensive type guards for malformed LLM JSON shapes.

    Weaker models sometimes emit `plans` as a string, an array of strings,
    or `plan` as a non-dict. These should not crash with AttributeError —
    they should fall through to the parse-failure retry path with a clean
    error message.
    """

    def _req(self):
        return ChatRequest(userMessage="test")

    def test_plans_as_string_does_not_crash(self):
        # plans is a string, not a list — must not raise AttributeError
        result = _parse_response(
            '{"responseType":"plans","plans":"not a list","message":"hi"}',
            self._req(),
        )
        # Falls through to message path since plans isn't a list
        assert result.responseType == "message"

    def test_plans_array_with_non_dict_elements_skipped(self):
        # plans contains a string element — should be skipped, not crash
        with pytest.raises(ValueError):
            _parse_response(
                '{"responseType":"plans","plans":["just a string"]}',
                self._req(),
            )

    def test_plans_array_one_bad_one_good_keeps_good(self):
        # The good option should survive even if one option has a bad action
        good_step = '{"id":"s1","description":"d","action":"writeValues","params":{"range":"A1","values":[["x"]]}}'
        good_plan = '{"summary":"good","steps":[' + good_step + ']}'
        bad_step = '{"id":"s1","description":"d","action":"manualStep","params":{}}'
        bad_plan = '{"summary":"bad","steps":[' + bad_step + ']}'
        text = (
            '{"responseType":"plans","message":"hi","plans":['
            '{"optionLabel":"Option A","plan":' + bad_plan + '},'
            '{"optionLabel":"Option B","plan":' + good_plan + '}'
            ']}'
        )
        result = _parse_response(text, self._req())
        # Good option survived
        assert result.responseType == "plans"
        assert len(result.plans) == 1
        assert result.plans[0].plan.steps[0].action.value == "writeValues"

    def test_singular_plan_as_non_dict_does_not_crash(self):
        # plan is a string instead of a dict — must not crash
        result = _parse_response(
            '{"responseType":"plan","plan":"not a dict","message":"hi"}',
            self._req(),
        )
        assert result.responseType == "message"


class TestLiteralEnumTightening:
    """A1: bare-string control-vocab fields are now Literal[...].

    These Literals are enforced through ``validate_plan``, which runs each
    step's params through the matching ``ACTION_PARAM_MODELS`` entry.
    Hallucinated values like ``chartType="spaghetti"`` should produce an
    INVALID_PARAMS error. Free-form data fields (range, values, formula,
    dateFormat, etc.) and Hebrew strings remain unaffected.
    """

    def test_chart_type_rejects_garbage(self):
        # Direct param model rejection
        with pytest.raises(ValueError):
            ACTION_PARAM_MODELS[StepAction.createChart](
                dataRange="A1:B10", chartType="spaghetti"
            )
        # And via validate_plan
        plan = _make_plan("createChart", {"dataRange": "A1:B10", "chartType": "spaghetti"})
        result = validate_plan(plan)
        assert not result.valid
        assert any("chartType" in e.message or "spaghetti" in e.message for e in result.errors)

    def test_chart_type_accepts_all_supported(self):
        for ct in ["columnClustered", "columnStacked", "bar", "line", "pie", "area", "scatter", "combo"]:
            model = ACTION_PARAM_MODELS[StepAction.createChart](dataRange="A1:B10", chartType=ct)
            assert model.chartType == ct

    def test_named_range_operation_rejects_garbage(self):
        with pytest.raises(ValueError):
            ACTION_PARAM_MODELS[StepAction.namedRange](
                operation="manifest", name="X", range="A1"
            )

    def test_clear_range_type_rejects_garbage(self):
        with pytest.raises(ValueError):
            ACTION_PARAM_MODELS[StepAction.clearRange](
                range="A1:B10", clearType="everything"
            )

    def test_categorize_rule_operator_rejects_garbage(self):
        with pytest.raises(ValueError):
            ACTION_PARAM_MODELS[StepAction.categorize](
                sourceRange="A:A", outputRange="B:B",
                rules=[{"operator": "fuzzymatch", "value": "x", "label": "X"}],
            )

    def test_freeform_string_fields_still_accept_anything(self):
        """outputFormat must remain free-form (Hebrew, custom patterns, etc.)."""
        model = ACTION_PARAM_MODELS[StepAction.normalizeDates](
            range="A:A", outputFormat="יום dd חודש mm שנה yyyy",
        )
        assert "יום" in model.outputFormat

    def test_hebrew_range_addresses_unaffected(self):
        """Range strings carry sheet names that may be Hebrew — must pass through."""
        model = ACTION_PARAM_MODELS[StepAction.writeValues](
            range="נתונים!A1:B5", values=[["שלום", "world"]],
        )
        assert model.range == "נתונים!A1:B5"

    def test_number_strings_in_values_pass_through(self):
        """Cell data like '123' or '0042' must NOT be coerced or rejected."""
        model = ACTION_PARAM_MODELS[StepAction.writeValues](
            range="A1:A3", values=[["0042"], ["1.5e10"], ["١٢٣"]],  # last is Arabic-Indic digits
        )
        # Pydantic preserves the strings as-is
        assert model.values[0][0] == "0042"
        assert model.values[2][0] == "١٢٣"


class TestNormalizeParamKeys:
    """F1: snake_case → camelCase normalization, including a generic fallback."""

    def test_explicit_overrides_remap_semantically(self):
        # lookup_column → lookupRange (different name, not just case)
        out = _normalize_param_keys({"lookup_column": "A:A", "match_type": "exact"})
        assert out == {"lookupRange": "A:A", "matchType": "exact"}

    def test_generic_snake_to_camel_for_unknown_keys(self):
        # Keys not in the explicit map should still be converted
        out = _normalize_param_keys({"some_brand_new_param": 42, "another_one": "x"})
        assert out == {"someBrandNewParam": 42, "anotherOne": "x"}

    def test_already_camel_case_unchanged(self):
        out = _normalize_param_keys({"sourceRange": "A:A", "outputRange": "B:B"})
        assert out == {"sourceRange": "A:A", "outputRange": "B:B"}

    def test_recurses_into_nested_dicts(self):
        out = _normalize_param_keys({
            "outer_key": {"inner_snake_key": "v"},
        })
        assert out == {"outerKey": {"innerSnakeKey": "v"}}

    def test_recurses_into_list_of_dicts(self):
        # Mirrors how categorize/rules work
        out = _normalize_param_keys({
            "rules": [
                {"match_value": "x", "label_text": "X"},
                {"match_value": "y", "label_text": "Y"},
            ],
        })
        assert out == {
            "rules": [
                {"matchValue": "x", "labelText": "X"},
                {"matchValue": "y", "labelText": "Y"},
            ],
        }

    def test_cell_data_arrays_pass_through_untouched(self):
        # values is a 2D array of cell data — strings inside must NOT be mutated
        out = _normalize_param_keys({
            "range": "A1:B2",
            "values": [["snake_value", "another_one"], [1, 2]],
        })
        assert out == {
            "range": "A1:B2",
            "values": [["snake_value", "another_one"], [1, 2]],
        }

    def test_hebrew_keys_unchanged(self):
        # Hebrew chars don't match the snake_case regex — passthrough
        out = _normalize_param_keys({"שדה": "ערך"})
        assert out == {"שדה": "ערך"}


class TestChatStreamRobustness:
    """E1: chat_stream must always emit a terminal `done` event.

    Failures inside prompt construction or capability lookup must not
    silently close the SSE stream — the frontend depends on `done` to
    leave its "thinking" state.
    """

    @pytest.mark.asyncio
    async def test_done_event_emitted_when_capability_lookup_raises(self, monkeypatch):
        # Force search_capabilities to blow up
        from app.services import capability_store

        def boom(_msg):
            raise RuntimeError("capability index unavailable")

        monkeypatch.setattr(capability_store, "search_capabilities", boom)

        req = ChatRequest(userMessage="anything")
        events = []
        async for sse in chat_stream(req):
            events.append(sse)

        # Must have produced at least one event, and the last must be `done`
        assert events, "chat_stream produced no SSE events at all"
        assert '"type": "done"' in events[-1]

    @pytest.mark.asyncio
    async def test_done_event_emitted_when_prompt_build_raises(self, monkeypatch):
        # Make _build_chat_messages explode
        from app.services import chat_service as cs

        async def boom(*_a, **_kw):
            raise RuntimeError("prompt construction failed")

        monkeypatch.setattr(cs, "_build_chat_messages", boom)

        req = ChatRequest(userMessage="anything")
        events = []
        async for sse in chat_stream(req):
            events.append(sse)

        assert events
        assert '"type": "done"' in events[-1]
        # The fallback message should be present in the final result
        assert "couldn't process" in events[-1] or "Sorry" in events[-1]
