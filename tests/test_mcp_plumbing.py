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
from langchain_mcp_adapters.client import MultiServerMCPClient

from tests.fixtures.build_fixtures import build_all

MCP_URL = os.environ.get("EXCEL_MCP_URL", "http://localhost:8017/mcp")


async def _get_tools():
    client = MultiServerMCPClient({"excel": {"url": MCP_URL, "transport": "streamable_http"}})
    try:
        return await client.get_tools()
    except Exception as exc:  # noqa: BLE001 - any connection failure means "skip"
        pytest.skip(f"No excel-mcp-server reachable at {MCP_URL}: {exc}")


def _tool_text(result) -> str:
    # langchain-mcp-adapters tool results come back as a list of content
    # blocks; our tools only ever return a single text block.
    return result[0]["text"] if isinstance(result, list) else result


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
