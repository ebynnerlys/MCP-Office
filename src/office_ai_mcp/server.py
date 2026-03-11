from __future__ import annotations

import argparse
from collections.abc import Callable
from typing import Literal

from mcp.server.fastmcp import FastMCP

from office_ai_mcp.config import Settings, get_settings
from office_ai_mcp.utils.logging import configure_logging

TransportName = Literal["stdio", "sse", "streamable-http"]
ToolRegistrar = Callable[[FastMCP, Settings], None]


def create_server_with_registrar(
    server_name: str,
    registrar: ToolRegistrar,
    settings: Settings | None = None,
) -> FastMCP:
    app_settings = settings or get_settings()
    configure_logging(app_settings.log_level)
    mcp = FastMCP(server_name, json_response=True)
    registrar(mcp, app_settings)
    return mcp


def build_parser(description: str = "Office AI MCP server") -> argparse.ArgumentParser:
    settings = get_settings()
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument(
        "--transport",
        choices=["stdio", "sse", "streamable-http"],
        default=settings.default_transport,
        help="MCP transport to expose.",
    )
    parser.add_argument("--host", default=settings.host, help="Host for HTTP based transports.")
    parser.add_argument("--port", type=int, default=settings.port, help="Port for HTTP transports.")
    return parser


def run_server(create_mcp: Callable[[], FastMCP], description: str = "Office AI MCP server") -> None:
    parser = build_parser(description)
    args = parser.parse_args()
    transport: TransportName = args.transport
    mcp = create_mcp()

    if transport == "stdio":
        mcp.run(transport=transport)
        return

    mcp.run(transport=transport, host=args.host, port=args.port)
