from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from office_ai_mcp.config import Settings
from office_ai_mcp.utils.backups import create_backup
from office_ai_mcp.utils.paths import ensure_directory, validate_file_path


def register_system_tools(mcp: FastMCP, settings: Settings) -> None:
    @mcp.tool(
        name="server_status",
        description="Return the current MCP server configuration, runtime defaults, and working directories.",
    )
    def server_status() -> dict[str, object]:
        """Return basic server metadata and configured working paths."""
        ensure_directory(settings.backup_dir)
        ensure_directory(settings.temp_dir)
        return {
            "name": settings.project_name,
            "version": settings.version,
            "default_transport": settings.default_transport,
            "allowed_roots": settings.allowed_roots,
            "backup_dir": settings.backup_dir,
            "temp_dir": settings.temp_dir,
        }

    @mcp.tool(
        name="create_working_backup",
        description="Create a timestamped backup copy of an Office document before editing it.",
    )
    def create_working_backup(path: str) -> dict[str, str]:
        """Create a backup of an existing file inside the configured backup directory."""
        source = validate_file_path(path, allowed_roots=settings.allowed_roots, must_exist=True)
        backup_path = create_backup(source, settings.backup_dir)
        return {"source_path": str(source), "backup_path": str(backup_path)}
