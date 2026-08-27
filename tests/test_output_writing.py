from __future__ import annotations

import json

from src.main import write_output


def test_writes_utf8_with_real_unicode_characters(tmp_path):
    extracted = {
        "feature_id": "019",
        "comm_matrix": [{"logical signal name": "Slope_Assist_Enabled_Disabled", "marker": "〇"}],
    }

    output_path = write_output(extracted, tmp_path, "019")

    assert output_path == tmp_path / "019_extracted.json"
    raw = output_path.read_bytes()
    # The literal UTF-8 bytes for "〇" (U+3007), not a \uXXXX escape sequence.
    assert "〇".encode("utf-8") in raw
    assert b"\\u3007" not in raw


def test_output_is_valid_json_round_trip(tmp_path):
    extracted = {"feature_id": "019", "requirements": {"metadata": [], "functional_requirements": []}}
    output_path = write_output(extracted, tmp_path, "019")
    assert json.loads(output_path.read_text(encoding="utf-8")) == extracted


def test_creates_output_directory_if_missing(tmp_path):
    nested = tmp_path / "does" / "not" / "exist" / "yet"
    output_path = write_output({"feature_id": "019"}, nested, "019")
    assert output_path.exists()
