from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from office_ai_mcp.models.requests import ExportPdfRequest, ReplaceTextRequest
from office_ai_mcp.services.word_service import WordService


def register_word_tools(mcp: FastMCP, service: WordService) -> None:
    @mcp.tool(
        name="word_get_structure",
        description="Inspect a Word document and return headings, paragraph counts, tables, comments, and track-changes status.",
    )
    def word_get_structure(path: str) -> dict[str, object]:
        """Inspect a Word document and summarize its visible structure."""
        return service.get_structure(path).model_dump()

    @mcp.tool(
        name="word_replace_text",
        description="Replace text across a Word document, optionally creating a backup before saving changes.",
    )
    def word_replace_text(path: str, find: str, replace: str, create_backup: bool = True) -> dict[str, object]:
        """Replace occurrences of one string with another in a Word document."""
        request = ReplaceTextRequest(path=path, find_text=find, replace_text=replace, create_backup=create_backup)
        return service.replace_text(
            path=request.path,
            find_text=request.find_text,
            replace_text=request.replace_text,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="word_export_pdf",
        description="Export a Word document to PDF without modifying the source file.",
    )
    def word_export_pdf(path: str, out_path: str) -> dict[str, object]:
        """Export a Word document to a PDF file."""
        request = ExportPdfRequest(path=path, out_path=out_path)
        return service.export_pdf(path=request.path, out_path=request.out_path).model_dump()
