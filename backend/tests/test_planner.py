"""
Tests for the planner service utilities.

These test JSON extraction and message building without requiring
an actual LLM API call.
"""

import pytest
from app.services.planner import extract_json, build_user_message, build_system_prompt
from app.models.request import PlanRequest, RangeTokenRef


class TestExtractJson:
    def test_plain_json(self):
        text = '{"planId": "test", "steps": []}'
        result = extract_json(text)
        assert result["planId"] == "test"

    def test_json_in_markdown_fence(self):
        text = 'Here is the plan:\n```json\n{"planId": "test"}\n```\nDone.'
        result = extract_json(text)
        assert result["planId"] == "test"

    def test_json_in_generic_fence(self):
        text = 'Plan:\n```\n{"planId": "test"}\n```'
        result = extract_json(text)
        assert result["planId"] == "test"

    def test_json_with_surrounding_text(self):
        text = 'I will create a plan: {"planId": "test", "confidence": 0.9} as requested.'
        result = extract_json(text)
        assert result["planId"] == "test"
        assert result["confidence"] == 0.9

    def test_invalid_json_raises(self):
        with pytest.raises(Exception):
            extract_json("not json at all")

    def test_nested_json(self):
        text = '{"planId": "test", "steps": [{"id": "s1", "params": {"range": "A1:B5"}}]}'
        result = extract_json(text)
        assert result["steps"][0]["id"] == "s1"


class TestBuildUserMessage:
    def test_simple_message(self):
        req = PlanRequest(userMessage="Sum column A")
        msg = build_user_message(req)
        assert "Sum column A" in msg

    def test_message_with_range_tokens(self):
        req = PlanRequest(
            userMessage="Sum these values",
            rangeTokens=[RangeTokenRef(address="A1:A10", sheetName="Sheet1")],
        )
        msg = build_user_message(req)
        assert "Sheet1" in msg
        assert "A1:A10" in msg

    def test_message_with_active_sheet(self):
        req = PlanRequest(
            userMessage="Sort the data",
            activeSheet="Sales",
        )
        msg = build_user_message(req)
        assert "Sales" in msg


class TestSystemPrompt:
    def test_system_prompt_not_empty(self):
        prompt = build_system_prompt()
        assert len(prompt) > 100

    def test_system_prompt_contains_actions(self):
        prompt = build_system_prompt()
        assert "writeFormula" in prompt
        assert "readRange" in prompt
        assert "preserveFormatting" in prompt
