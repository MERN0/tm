"""CLI entrypoint for Node 1 (agentic data extraction).

The three workbook arguments are filenames (or paths relative to the Excel
MCP server's EXCEL_FILES_PATH) -- the server reads the files, not this
process, so they must be resolvable on whatever machine EXCEL_MCP_URL points
to.

Example:
    export OPENAI_API_KEY=...
    export OPENAI_API_BASE=...        # optional, for an OpenAI-compatible endpoint
    export OPENAI_MODEL=gpt-4o        # optional, defaults to gpt-4o
    export EXCEL_MCP_URL=http://localhost:8017/mcp   # optional, this is the default

    python -m src.main --feature 019 \\
        --sysreq "System Requirements.xlsx" \\
        --commands "TE_TMHC_Command_List.xlsx" \\
        --config "TE_TMHC_Configuration_File.xlsx" \\
        --output out/
"""

from __future__ import annotations

import argparse
import asyncio
import json
import os
import sys
from pathlib import Path

from src.graph.build_graph import compiled_graph


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Agentically extract traceability data for a system requirement feature."
    )
    parser.add_argument("--feature", required=True, help="Feature id, e.g. 019")
    parser.add_argument(
        "--sysreq", required=True, help="System Requirements workbook filename (as seen by the Excel MCP server)"
    )
    parser.add_argument(
        "--commands", required=True, help="Command List workbook filename (as seen by the Excel MCP server)"
    )
    parser.add_argument(
        "--config", required=True, help="Configuration File workbook filename (as seen by the Excel MCP server)"
    )
    parser.add_argument("--output", required=True, help="Output directory for the extracted JSON")
    parser.add_argument("--model", default=None, help="Override OPENAI_MODEL for this run")
    parser.add_argument("--mcp-url", default=None, help="Override EXCEL_MCP_URL for this run")
    return parser.parse_args(argv)


async def _run(args: argparse.Namespace) -> dict:
    initial_state = {
        "feature_id": args.feature,
        "sysreq_path": args.sysreq,
        "command_list_path": args.commands,
        "config_path": args.config,
        "output_dir": args.output,
        "extracted": {},
        "warnings": [],
    }
    return await compiled_graph.ainvoke(initial_state)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    if args.model:
        os.environ["OPENAI_MODEL"] = args.model
    if args.mcp_url:
        os.environ["EXCEL_MCP_URL"] = args.mcp_url

    result = asyncio.run(_run(args))

    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{args.feature}_extracted.json"
    output_path.write_text(json.dumps(result["extracted"], indent=2, default=str))

    print(f"Wrote {output_path}")
    for warning in result.get("warnings", []):
        print(f"WARNING: {warning}", file=sys.stderr)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
