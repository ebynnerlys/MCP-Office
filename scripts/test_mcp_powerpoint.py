from __future__ import annotations

import argparse
import asyncio
import json
from pathlib import Path
from typing import Any

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Test the local Office AI MCP server with a PowerPoint file.")
    parser.add_argument("path", help="Path to the .ppt or .pptx file to inspect.")
    parser.add_argument(
        "--slide-index",
        type=int,
        default=1,
        help="Slide number to extract text from after listing slides.",
    )
    return parser


def result_to_json_payload(result: Any) -> Any:
    if hasattr(result, "model_dump"):
        return result.model_dump(by_alias=True)
    return result


async def run_test(presentation_path: Path, slide_index: int) -> None:
    workspace = Path(__file__).resolve().parents[1]
    python_executable = workspace / ".venv" / "Scripts" / "python.exe"

    server_params = StdioServerParameters(
        command=str(python_executable),
        args=["-m", "office_ai_mcp.powerpoint_server", "--transport", "stdio"],
        cwd=str(workspace),
        env={"PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"},
    )

    async with stdio_client(server_params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()

            tools = await session.list_tools()
            print("TOOLS")
            print(json.dumps([tool.name for tool in tools.tools], ensure_ascii=False, indent=2))

            slides = await session.call_tool("ppt_list_slides", {"path": str(presentation_path)})
            print("\nPPT_LIST_SLIDES")
            print(json.dumps(result_to_json_payload(slides), ensure_ascii=False, indent=2))

            slide_text = await session.call_tool(
                "ppt_get_slide_text",
                {"path": str(presentation_path), "slide_index": slide_index},
            )
            print(f"\nPPT_GET_SLIDE_TEXT slide={slide_index}")
            print(json.dumps(result_to_json_payload(slide_text), ensure_ascii=False, indent=2))


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    presentation_path = Path(args.path).expanduser().resolve()
    asyncio.run(run_test(presentation_path, args.slide_index))


if __name__ == "__main__":
    main()