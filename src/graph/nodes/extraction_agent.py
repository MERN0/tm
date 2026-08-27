"""Node 1: LLM-agent-driven data extraction.

Unlike the earlier fixed-logic implementation, this node does not parse
Excel files directly. It hands a ReAct-style agent the Excel MCP server's
tools (haris-musa/excel-mcp-server, running separately and reachable over
streamable-http) plus a system prompt encoding all the sheet traceability
knowledge, and lets the LLM itself navigate the workbooks, apply the
filtering/matching rules, and produce the final JSON.
"""

from __future__ import annotations

from langchain.agents import create_agent
from langchain_mcp_adapters.client import MultiServerMCPClient
from langchain_openai import ChatOpenAI

from src.agent.json_extraction import extract_json_from_text
from src.agent.prompts import SYSTEM_PROMPT, build_human_prompt
from src.config import load_agent_config
from src.graph.state import GraphState


def _tool_result_text(result) -> str:
    if isinstance(result, list) and result and isinstance(result[0], dict):
        return str(result[0].get("text", result))
    return str(result)


async def preflight_check_files(tools, filepaths: dict[str, str]) -> list[str]:
    """Directly call get_workbook_metadata for each workbook before handing
    control to the LLM agent.

    The MCP spec's reference filesystem server helps a model avoid bad paths
    via discovery tools like `list_allowed_directories`/`list_directory`,
    called before `read_file`. haris-musa/excel-mcp-server exposes no
    equivalent discovery tool, so this function does the same "discover
    before acting" check ourselves: if a given filepath can't be opened, we
    find out in one direct tool call instead of the LLM burning several
    reasoning turns on a mistyped or partial path (as previously happened
    with a bare "inputs" directory fragment) before giving up.
    """
    metadata_tool = next((t for t in tools if t.name == "get_workbook_metadata"), None)
    if metadata_tool is None:
        return []

    warnings: list[str] = []
    for label, filepath in filepaths.items():
        try:
            result = await metadata_tool.ainvoke({"filepath": filepath, "include_ranges": False})
        except Exception as exc:  # noqa: BLE001 - report as a warning either way
            warnings.append(f"{label} workbook '{filepath}' could not be opened: {exc}")
            continue
        text = _tool_result_text(result)
        if text.strip().lower().startswith("error"):
            warnings.append(f"{label} workbook '{filepath}' could not be opened: {text.strip()}")
    return warnings


async def extract_data(state: GraphState) -> GraphState:
    config = load_agent_config()

    mcp_client = MultiServerMCPClient(
        {
            "excel": {
                "url": config.mcp_url,
                "transport": "streamable_http",
            }
        }
    )
    tools = await mcp_client.get_tools()

    preflight_warnings = await preflight_check_files(
        tools,
        {
            "System Requirements": state["sysreq_path"],
            "Command List": state["command_list_path"],
            "Configuration File": state["config_path"],
        },
    )
    if preflight_warnings:
        return {**state, "extracted": {}, "warnings": preflight_warnings}

    model = ChatOpenAI(
        model=config.openai_model,
        api_key=config.openai_api_key,
        base_url=config.openai_api_base,
        temperature=0,
    )

    agent = create_agent(model, tools, system_prompt=SYSTEM_PROMPT)

    human_prompt = build_human_prompt(
        feature_id=state["feature_id"],
        sysreq_filename=state["sysreq_path"],
        command_list_filename=state["command_list_path"],
        config_filename=state["config_path"],
    )

    result = await agent.ainvoke({"messages": [{"role": "user", "content": human_prompt}]})
    final_message = result["messages"][-1]
    extracted, warnings = extract_json_from_text(final_message.content)

    return {**state, "extracted": extracted, "warnings": warnings}
