from __future__ import annotations

from src.agent.json_extraction import extract_json_from_text


def test_extracts_fenced_json_block():
    text = 'Here is the result:\n```json\n{"feature_id": "019", "warnings": []}\n```'
    extracted, warnings = extract_json_from_text(text)
    assert extracted == {"feature_id": "019"}
    assert warnings == []


def test_pulls_warnings_out_of_payload():
    text = '```json\n{"feature_id": "019", "warnings": ["a", "b"]}\n```'
    extracted, warnings = extract_json_from_text(text)
    assert "warnings" not in extracted
    assert warnings == ["a", "b"]


def test_falls_back_to_raw_json_without_fence():
    text = '{"feature_id": "019", "warnings": []}'
    extracted, warnings = extract_json_from_text(text)
    assert extracted == {"feature_id": "019"}
    assert warnings == []


def test_malformed_json_does_not_raise():
    text = "I couldn't finish extracting the data, sorry!"
    extracted, warnings = extract_json_from_text(text)
    assert extracted["_raw_response"] == text
    assert len(warnings) == 1


def test_non_object_json_does_not_raise():
    text = "```json\n[1, 2, 3]\n```"
    extracted, warnings = extract_json_from_text(text)
    assert extracted["_raw_response"] == text
    assert len(warnings) == 1


def test_non_list_warnings_coerced():
    text = '```json\n{"feature_id": "019", "warnings": "one problem"}\n```'
    extracted, warnings = extract_json_from_text(text)
    assert warnings == ["one problem"]
