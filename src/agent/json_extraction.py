"""Pulls the final JSON object out of an LLM agent's closing message."""

from __future__ import annotations

import json
import re

_FENCE_RE = re.compile(r"```(?:json)?\s*(\{.*\})\s*```", re.DOTALL)


def extract_json_from_text(text: str) -> tuple[dict, list[str]]:
    """Parse the agent's final message into (extracted_dict, warnings).

    Tries a ```json fenced block first (what the system prompt asks for),
    then falls back to treating the whole message as JSON. If parsing fails
    entirely, the raw text is preserved under "_raw_response" and a warning
    is added rather than raising, so a malformed agent reply is still
    inspectable instead of crashing the pipeline.
    """
    candidate = text.strip()
    match = _FENCE_RE.search(text)
    if match:
        candidate = match.group(1)

    try:
        parsed = json.loads(candidate)
    except json.JSONDecodeError:
        return (
            {"_raw_response": text},
            ["Could not parse the agent's final reply as JSON; see _raw_response for the raw text"],
        )

    if not isinstance(parsed, dict):
        return (
            {"_raw_response": text},
            ["Agent's final reply parsed as JSON but was not a JSON object; see _raw_response"],
        )

    warnings = parsed.pop("warnings", [])
    if not isinstance(warnings, list):
        warnings = [str(warnings)]
    return parsed, warnings
