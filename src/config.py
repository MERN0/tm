"""Environment-driven configuration for the LLM + MCP agent."""

from __future__ import annotations

import os
from dataclasses import dataclass


@dataclass
class AgentConfig:
    openai_api_key: str
    openai_api_base: str | None
    openai_model: str
    mcp_url: str


def load_agent_config(
    *,
    model: str | None = None,
    mcp_url: str | None = None,
) -> AgentConfig:
    """Build the agent config from environment variables, with optional
    explicit overrides (e.g. from CLI flags)."""
    return AgentConfig(
        openai_api_key=os.environ.get("OPENAI_API_KEY", ""),
        openai_api_base=os.environ.get("OPENAI_API_BASE") or os.environ.get("OPENAI_BASE_URL"),
        openai_model=model or os.environ.get("OPENAI_MODEL", "gpt-4o"),
        mcp_url=mcp_url or os.environ.get("EXCEL_MCP_URL", "http://localhost:8017/mcp"),
    )
