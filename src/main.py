"""CLI entrypoint for Node 1 (data extraction).

Example:
    python -m src.main --feature 019 \\
        --sysreq "System Requirements.xlsx" \\
        --commands "TE_TMHC_Command_List.xlsx" \\
        --config "TE_TMHC_Configuration_File.xlsx" \\
        --output out/
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from src.graph.build_graph import compiled_graph


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extract traceability data for a system requirement feature.")
    parser.add_argument("--feature", required=True, help="Feature id, e.g. 019")
    parser.add_argument("--sysreq", required=True, help="Path to System Requirements.xlsx")
    parser.add_argument("--commands", required=True, help="Path to TE_TMHC_Command_List.xlsx")
    parser.add_argument("--config", required=True, help="Path to TE_TMHC_Configuration_File.xlsx")
    parser.add_argument("--output", required=True, help="Output directory for the extracted JSON")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    initial_state = {
        "feature_id": args.feature,
        "sysreq_path": args.sysreq,
        "command_list_path": args.commands,
        "config_path": args.config,
        "output_dir": args.output,
        "extracted": {},
        "warnings": [],
    }

    result = compiled_graph.invoke(initial_state)

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
