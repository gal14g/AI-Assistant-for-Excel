"""
Tests for the planner service utilities.

These test JSON extraction from LLM responses without requiring
an actual LLM API call.
"""

import pytest
from app.services.planner import extract_json


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
