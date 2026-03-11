from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from office_ai_mcp.config import Settings, get_settings
from office_ai_mcp.server import create_server_with_registrar, run_server
from office_ai_mcp.tools.registry import register_word_stack


def create_word_server(settings: Settings | None = None) -> FastMCP:
    app_settings = settings or get_settings()
    return create_server_with_registrar("Office AI Word MCP", register_word_stack, app_settings)


def main() -> None:
    run_server(create_word_server, description="Office AI Word MCP server")


if __name__ == "__main__":
    main()