from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from office_ai_mcp.models.requests import (
    DocumentPathRequest,
    ExportPdfRequest,
    ReplaceTextRequest,
    WordDocumentPropertiesRequest,
)
from office_ai_mcp.services.word_service import WordService


def register_word_tools(mcp: FastMCP, service: WordService) -> None:
    @mcp.tool(
        name="word_get_document_properties",
        description="Read built-in and custom document properties from a Word file, including timestamps and total editing time when available.",
    )
    def word_get_document_properties(path: str) -> dict[str, object]:
        """Inspect built-in and custom Word document properties."""
        request = DocumentPathRequest(path=path, create_backup=False)
        return service.get_document_properties(path=request.path).model_dump()

    @mcp.tool(
        name="word_set_document_properties",
        description="Update built-in Word document properties such as author, title, subject, keywords, comments, category, company, manager, last author, timestamps, total editing time, and revision number.",
    )
    def word_set_document_properties(
        path: str,
        author: str | None = None,
        title: str | None = None,
        subject: str | None = None,
        keywords: str | None = None,
        comments: str | None = None,
        category: str | None = None,
        company: str | None = None,
        manager: str | None = None,
        last_author: str | None = None,
        creation_date: str | None = None,
        last_save_time: str | None = None,
        total_editing_time: int | None = None,
        revision_number: str | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Write selected built-in document properties."""
        request = WordDocumentPropertiesRequest(
            path=path,
            author=author,
            title=title,
            subject=subject,
            keywords=keywords,
            comments=comments,
            category=category,
            company=company,
            manager=manager,
            last_author=last_author,
            creation_date=creation_date,
            last_save_time=last_save_time,
            total_editing_time=total_editing_time,
            revision_number=revision_number,
            create_backup=create_backup,
        )
        return service.set_document_properties(
            path=request.path,
            author=request.author,
            title=request.title,
            subject=request.subject,
            keywords=request.keywords,
            comments=request.comments,
            category=request.category,
            company=request.company,
            manager=request.manager,
            last_author=request.last_author,
            creation_date=request.creation_date,
            last_save_time=request.last_save_time,
            total_editing_time=request.total_editing_time,
            revision_number=request.revision_number,
            create_backup=request.create_backup,
        ).model_dump()

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
