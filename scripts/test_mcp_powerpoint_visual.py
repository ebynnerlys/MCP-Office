from __future__ import annotations

import argparse
import asyncio
import json
from pathlib import Path
from typing import Any

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Exercise visual PowerPoint editing tools exposed by the local Office AI MCP server."
    )
    parser.add_argument("path", help="Path to the PowerPoint file to modify.")
    parser.add_argument("--slide-index", type=int, default=1, help="Slide index to update.")
    parser.add_argument("--title-shape-index", type=int, default=1, help="Shape index for title/fill/border.")
    parser.add_argument("--body-shape-index", type=int, default=2, help="Shape index for text formatting.")
    parser.add_argument(
        "--table-shape-index",
        type=int,
        default=None,
        help="Optional table shape index for per-cell animation examples.",
    )
    parser.add_argument(
        "--table-row-index",
        type=int,
        default=1,
        help="Table row to animate when --table-shape-index is provided.",
    )
    parser.add_argument(
        "--table-column-index",
        type=int,
        default=1,
        help="Table column to animate when --table-shape-index is provided.",
    )
    return parser


def result_to_json_payload(result: Any) -> Any:
    if hasattr(result, "model_dump"):
        return result.model_dump(by_alias=True)
    return result


async def run_test(
    presentation_path: Path,
    slide_index: int,
    title_shape_index: int,
    body_shape_index: int,
    table_shape_index: int | None,
    table_row_index: int,
    table_column_index: int,
) -> None:
    workspace = Path(__file__).resolve().parents[1]
    python_executable = workspace / ".venv" / "Scripts" / "python.exe"

    server_params = StdioServerParameters(
        command=str(python_executable),
        args=["-m", "office_ai_mcp.powerpoint_server", "--transport", "stdio"],
        cwd=str(workspace),
        env={"PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"},
    )

    operations = [
        (
            "ppt_set_slide_transition",
            {
                "path": str(presentation_path),
                "slide_index": slide_index,
                "effect": "fade",
                "speed": "fast",
                "advance_on_click": True,
                "create_backup": True,
            },
        ),
        (
            "ppt_set_shape_text_style",
            {
                "path": str(presentation_path),
                "slide_index": slide_index,
                "shape_index": body_shape_index,
                "font_name": "Aptos",
                "font_size": 22,
                "bold": True,
                "color": "#0B5FFF",
                "alignment": "center",
                "create_backup": True,
            },
        ),
        (
            "ppt_set_shape_fill",
            {
                "path": str(presentation_path),
                "slide_index": slide_index,
                "shape_index": title_shape_index,
                "color": "#FFF4CC",
                "transparency": 0.05,
                "create_backup": True,
            },
        ),
        (
            "ppt_set_shape_line",
            {
                "path": str(presentation_path),
                "slide_index": slide_index,
                "shape_index": title_shape_index,
                "color": "#FF7A00",
                "weight": 2.5,
                "visible": True,
                "create_backup": True,
            },
        ),
        (
            "ppt_set_slide_background",
            {
                "path": str(presentation_path),
                "slide_index": slide_index,
                "color": "#F3F8FF",
                "follow_master": False,
                "create_backup": True,
            },
        ),
        (
            "ppt_add_shape_animation",
            {
                "path": str(presentation_path),
                "slide_index": slide_index,
                "shape_index": title_shape_index,
                "effect": "bounce",
                "trigger": "on_click",
                "duration_seconds": 1.0,
                "delay_seconds": 0.0,
                "create_backup": True,
            },
        ),
        (
            "ppt_add_element_animation",
            {
                "path": str(presentation_path),
                "slide_index": slide_index,
                "shape_index": title_shape_index,
                "effect": "appear",
                "trigger": "with_previous",
                "target_kind": "text",
                "animation_level": "all_text_levels",
                "duration_seconds": 0.6,
                "delay_seconds": 0.0,
                "create_backup": True,
            },
        ),
        (
            "ppt_add_element_animation",
            {
                "path": str(presentation_path),
                "slide_index": slide_index,
                "shape_index": body_shape_index,
                "effect": "wipe",
                "trigger": "after_previous",
                "target_kind": "text",
                "animation_level": "all_text_levels",
                "duration_seconds": 0.8,
                "delay_seconds": 0.2,
                "create_backup": True,
            },
        ),
    ]

    if table_shape_index is not None:
        operations.append(
            (
                "ppt_add_element_animation",
                {
                    "path": str(presentation_path),
                    "slide_index": slide_index,
                    "shape_index": table_shape_index,
                    "effect": "zoom",
                    "trigger": "after_previous",
                    "target_kind": "table_cell",
                    "row_index": table_row_index,
                    "column_index": table_column_index,
                    "duration_seconds": 0.7,
                    "delay_seconds": 0.1,
                    "create_backup": True,
                },
            )
        )

    async with stdio_client(server_params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()

            for tool_name, arguments in operations:
                result = await session.call_tool(tool_name, arguments)
                print(f"\n{tool_name}")
                print(json.dumps(result_to_json_payload(result), ensure_ascii=False, indent=2)[:4000])

            shapes = await session.call_tool(
                "ppt_get_slide_shapes",
                {"path": str(presentation_path), "slide_index": slide_index},
            )
            print("\nppt_get_slide_shapes")
            print(json.dumps(result_to_json_payload(shapes), ensure_ascii=False, indent=2)[:6000])

            animations = await session.call_tool(
                "ppt_get_slide_animations",
                {"path": str(presentation_path), "slide_index": slide_index},
            )
            print("\nppt_get_slide_animations")
            print(json.dumps(result_to_json_payload(animations), ensure_ascii=False, indent=2)[:6000])

            cleared = await session.call_tool(
                "ppt_clear_slide_animations",
                {"path": str(presentation_path), "slide_index": slide_index, "create_backup": True},
            )
            print("\nppt_clear_slide_animations")
            print(json.dumps(result_to_json_payload(cleared), ensure_ascii=False, indent=2))


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    presentation_path = Path(args.path).expanduser().resolve()
    asyncio.run(
        run_test(
            presentation_path,
            slide_index=args.slide_index,
            title_shape_index=args.title_shape_index,
            body_shape_index=args.body_shape_index,
            table_shape_index=args.table_shape_index,
            table_row_index=args.table_row_index,
            table_column_index=args.table_column_index,
        )
    )


if __name__ == "__main__":
    main()