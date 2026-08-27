"""Validates the MCP tool-loading plumbing (and confirms our fixtures are
shaped the way the system prompt assumes) against a locally running
excel-mcp-server. This does NOT invoke an LLM -- it only exercises the tool
client, which is the part we can verify without OpenAI credentials.

Skips itself if no server is reachable at EXCEL_MCP_URL (default
http://localhost:8017/mcp), since that server is expected to run as a
separate process the developer starts explicitly, e.g.:

    EXCEL_FILES_PATH=tests/fixtures EXCEL_MCP_PORT=8017 excel-mcp-server streamable-http
"""

from __future__ import annotations

import json
import os

import pytest
from langchain.agents import create_agent
from langchain_mcp_adapters.client import MultiServerMCPClient
from langchain_openai import ChatOpenAI

from src.agent.prompts import SYSTEM_PROMPT
from src.graph.nodes.extraction_agent import _tool_result_text as _tool_text
from src.graph.nodes.extraction_agent import preflight_check_files
from tests.fixtures.build_fixtures import build_all

MCP_URL = os.environ.get("EXCEL_MCP_URL", "http://localhost:8017/mcp")


async def _get_tools():
    client = MultiServerMCPClient({"excel": {"url": MCP_URL, "transport": "streamable_http"}})
    try:
        return await client.get_tools()
    except Exception as exc:  # noqa: BLE001 - any connection failure means "skip"
        pytest.skip(f"No excel-mcp-server reachable at {MCP_URL}: {exc}")


@pytest.mark.asyncio
async def test_expected_tools_are_present():
    build_all()
    tools = await _get_tools()
    names = {t.name for t in tools}
    for expected in ("get_workbook_metadata", "read_data_from_excel", "get_merged_cells"):
        assert expected in names


@pytest.mark.asyncio
async def test_index_sheet_readable_and_shaped_as_expected():
    build_all()
    tools = await _get_tools()
    read_tool = next(t for t in tools if t.name == "read_data_from_excel")

    result = await read_tool.ainvoke({"filepath": "System Requirements.xlsx", "sheet_name": "Index"})
    payload = json.loads(_tool_text(result))
    values = {c["address"]: c["value"] for c in payload["cells"]}
    assert values["C1"] == "Feature Name"
    assert values["B3"] == "019"
    assert values["C3"] == "Slope Assist"


@pytest.mark.asyncio
async def test_comm_matrix_uses_ideographic_zero_marker():
    build_all()
    tools = await _get_tools()
    read_tool = next(t for t in tools if t.name == "read_data_from_excel")

    result = await read_tool.ainvoke(
        {"filepath": "System Requirements.xlsx", "sheet_name": "Master Comm Matrix (CAN)"}
    )
    payload = json.loads(_tool_text(result))
    values = {c["address"]: c["value"] for c in payload["cells"]}
    assert values["I1"] == "019"
    assert values["I2"] == "〇"
    assert values["I2"] != "O"


@pytest.mark.asyncio
async def test_model_input_mapping_merge_detected():
    build_all()
    tools = await _get_tools()
    merged_tool = next(t for t in tools if t.name == "get_merged_cells")

    result = await merged_tool.ainvoke(
        {"filepath": "TE_TMHC_Configuration_File.xlsx", "sheet_name": "Model_Input_Mapping"}
    )
    text = _tool_text(result)
    assert "B2:B3" in text


@pytest.mark.asyncio
async def test_preflight_check_passes_for_valid_filenames():
    build_all()
    tools = await _get_tools()
    warnings = await preflight_check_files(
        tools,
        {
            "System Requirements": "System Requirements.xlsx",
            "Command List": "TE_TMHC_Command_List.xlsx",
            "Configuration File": "TE_TMHC_Configuration_File.xlsx",
        },
    )
    assert warnings == []


@pytest.mark.asyncio
async def test_preflight_check_reports_bare_directory_path():
    # Regression test for the observed failure mode: the agent called a tool
    # with just "inputs" (a subfolder, not a filename) and got a confusing
    # "File not found: .../excel_files/inputs" error with no clear signal
    # about what was actually wrong. preflight_check_files should catch this
    # up front with a message naming which workbook and path failed.
    build_all()
    tools = await _get_tools()
    warnings = await preflight_check_files(tools, {"System Requirements": "inputs"})
    assert len(warnings) == 1
    assert "System Requirements" in warnings[0]
    assert "inputs" in warnings[0]


@pytest.mark.asyncio
async def test_create_agent_builds_with_mcp_tools_and_system_prompt():
    build_all()
    tools = await _get_tools()
    model = ChatOpenAI(model="gpt-4o", api_key="dummy-key-for-structural-test", temperature=0)
    agent = create_agent(model, tools, system_prompt=SYSTEM_PROMPT)
    assert hasattr(agent, "ainvoke")
