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


class TestExtractJsonStressCases:
    """
    Real-world LLM output is messy. These cases cover the failure modes we've
    actually seen in production logs: trailing commas, truncated braces,
    Hebrew / RTL content inside strings, smart quotes from models that
    auto-correct, and JSON arrays wrapped in explanatory prose.

    When any of these regress, chat silently fails because extract_json()
    returns None and the planner raises a 500 — so we pin them here.
    """

    def test_trailing_comma_in_object(self):
        # json-repair fallback handles this
        text = '{"planId": "test", "confidence": 0.9,}'
        result = extract_json(text)
        assert result["planId"] == "test"

    def test_trailing_comma_in_array(self):
        text = '{"steps": [1, 2, 3,]}'
        result = extract_json(text)
        assert result["steps"] == [1, 2, 3]

    def test_hebrew_inside_strings(self):
        # Hebrew summary + step description must survive unchanged.
        text = (
            '{"planId": "p1", "summary": "יצירת טבלת ציר", '
            '"steps": [{"description": "מיון לפי עמודה"}]}'
        )
        result = extract_json(text)
        assert result["summary"] == "יצירת טבלת ציר"
        assert result["steps"][0]["description"] == "מיון לפי עמודה"

    def test_hebrew_prose_surrounds_json(self):
        # Models sometimes wrap output in the user's language.
        text = 'הנה התוכנית: {"planId": "test", "summary": "סיכום"} בסדר?'
        result = extract_json(text)
        assert result["summary"] == "סיכום"

    def test_nested_markdown_fence_with_json(self):
        # Explanatory prose BEFORE the fence used to confuse the brace-scan.
        text = (
            "Sure — here's what I'll do:\n"
            "```json\n"
            '{"planId": "x", "steps": [{"id": "s1", "action": "readRange"}]}\n'
            "```\n"
            "Let me know if you want changes."
        )
        result = extract_json(text)
        assert result["planId"] == "x"
        assert result["steps"][0]["action"] == "readRange"

    def test_deep_nesting(self):
        # 6-step plan with cross-step bindings — simulates complex multi-action
        # output. If brace-matching regresses, only the first object gets
        # parsed and `steps` comes back empty.
        text = (
            '{"planId": "complex", "steps": ['
            '{"id": "s1", "action": "readRange", "params": {"range": "A1:B10"}},'
            '{"id": "s2", "action": "groupSum", "dependsOn": ["s1"], '
            '"params": {"dataRange": "{{step_s1.range}}", "groupByColumn": 1}},'
            '{"id": "s3", "action": "createChart", "dependsOn": ["s2"], '
            '"params": {"dataRange": "{{step_s2.outputRange}}", "chartType": "bar"}}'
            "]}"
        )
        result = extract_json(text)
        assert len(result["steps"]) == 3
        assert result["steps"][1]["dependsOn"] == ["s1"]

    def test_unclosed_brace_recovered_by_json_repair(self):
        # Model cut off mid-stream — json-repair closes the object.
        text = '{"planId": "test", "summary": "incomplete"'
        result = extract_json(text)
        assert result["planId"] == "test"

    def test_smart_quotes_repaired(self):
        # Some models (especially Hebrew-tuned ones) emit curly quotes.
        text = '{\u201cplanId\u201d: \u201ctest\u201d}'
        result = extract_json(text)
        assert result.get("planId") == "test"

    def test_json_with_newlines_inside_strings(self):
        # Escaped newlines inside string values — legal JSON, but trips
        # naive regex-based extractors.
        text = r'{"summary": "line1\nline2", "steps": []}'
        result = extract_json(text)
        assert "line1" in result["summary"]
        assert "line2" in result["summary"]
