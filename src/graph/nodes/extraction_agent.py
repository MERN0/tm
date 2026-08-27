"""Node 1: LLM-agent-driven data extraction.

Unlike the earlier fixed-logic implementation, this node does not parse
Excel files directly. It hands a ReAct-style LangGraph agent the Excel MCP
server's tools (haris-musa/excel-mcp-server, running separately and reachable
over streamable-http) plus a system prompt encoding all the sheet
traceability knowledge, and lets the LLM itself navigate the workbooks,
apply the filtering/matching rules, and produce the final JSON.
"""

from __future__ import annotations

from langchain_mcp_adapters.client import MultiServerMCPClient
from langchain_openai import ChatOpenAI

# NOTE: create_react_agent is deprecated in favor of langchain.agents.create_agent
# as of LangGraph v1.0 (removal planned for v2.0). Not yet migrated to keep the
# dependency footprint minimal (would pull in the full `langchain` package);
# revisit when upgrading.
from langgraph.prebuilt import create_react_agent

from src.agent.json_extraction import extract_json_from_text
from src.agent.prompts import SYSTEM_PROMPT, build_human_prompt
from src.config import load_agent_config
from src.graph.state import GraphState


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

    model = ChatOpenAI(
        model=config.openai_model,
        api_key=config.openai_api_key,
        base_url=config.openai_api_base,
        temperature=0,
    )

    agent = create_react_agent(model, tools, prompt=SYSTEM_PROMPT)

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
