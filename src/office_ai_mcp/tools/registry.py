from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from office_ai_mcp.config import Settings
from office_ai_mcp.services.excel_service import ExcelService
from office_ai_mcp.services.powerpoint_service import PowerPointService
from office_ai_mcp.services.word_service import WordService
from office_ai_mcp.tools.excel_tools import register_excel_tools
from office_ai_mcp.tools.powerpoint_tools import register_powerpoint_tools
from office_ai_mcp.tools.system_tools import register_system_tools
from office_ai_mcp.tools.word_tools import register_word_tools


def register_word_stack(mcp: FastMCP, settings: Settings) -> None:
    register_system_tools(mcp, settings)
    register_word_tools(mcp, WordService(settings))


def register_excel_stack(mcp: FastMCP, settings: Settings) -> None:
    register_system_tools(mcp, settings)
    register_excel_tools(mcp, ExcelService(settings))


def register_powerpoint_stack(mcp: FastMCP, settings: Settings) -> None:
    register_system_tools(mcp, settings)
    register_powerpoint_tools(mcp, PowerPointService(settings))
