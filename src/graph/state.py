"""Shared state passed between LangGraph nodes."""

from __future__ import annotations

from typing import TypedDict


class GraphState(TypedDict):
    feature_id: str
    sysreq_path: str
    command_list_path: str
    config_path: str
    output_dir: str
    extracted: dict
    warnings: list[str]
