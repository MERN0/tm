"""Assembles the LangGraph pipeline.

Currently a single node: an LLM agent (bound to the Excel MCP server's
tools) that performs the Node 1 data extraction. Future nodes (e.g. LLM-
driven test case generation using system/human prompts) can be appended here
without restructuring this module.
"""

from __future__ import annotations

from langgraph.graph import END, StateGraph

from src.graph.nodes.extraction_agent import extract_data
from src.graph.state import GraphState


def build_graph():
    graph = StateGraph(GraphState)
    graph.add_node("extract_data", extract_data)
    graph.set_entry_point("extract_data")
    graph.add_edge("extract_data", END)
    return graph.compile()


compiled_graph = build_graph()
