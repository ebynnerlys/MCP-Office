from __future__ import annotations

from contextlib import suppress

from tenacity import retry, stop_after_attempt, wait_fixed

from office_ai_mcp.config import Settings
from office_ai_mcp.models.responses import OperationResult, SectionSummary, WordStructureResult
from office_ai_mcp.services.base import OfficeService
from office_ai_mcp.utils.com_cleanup import office_application


class WordService(OfficeService):
    allowed_suffixes = (".doc", ".docx", ".docm")

    def __init__(self, settings: Settings) -> None:
        super().__init__(settings)

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def get_structure(self, path: str) -> WordStructureResult:
        source = self.resolve_document_path(path)
        with office_application("Word.Application", visible=self.settings.office_visible) as word:
            document = None
            try:
                document = word.Documents.Open(str(source), ReadOnly=True)
                headings: list[SectionSummary] = []

                for index in range(1, int(document.Paragraphs.Count) + 1):
                    if len(headings) >= 25:
                        break

                    paragraph = document.Paragraphs(index)
                    text = str(paragraph.Range.Text).replace("\r", " ").strip()
                    if not text:
                        continue

                    style_name = ""
                    with suppress(Exception):
                        style_name = str(paragraph.Range.Style)

                    normalized_style = style_name.lower()
                    if normalized_style.startswith("heading") or normalized_style.startswith("titulo"):
                        headings.append(
                            SectionSummary(index=index, text=text, style_name=style_name or None)
                        )

                return WordStructureResult(
                    file_path=str(source),
                    paragraph_count=int(document.Paragraphs.Count),
                    table_count=int(document.Tables.Count),
                    comment_count=int(document.Comments.Count),
                    track_changes_enabled=bool(document.TrackRevisions),
                    headings=headings,
                )
            finally:
                if document is not None:
                    with suppress(Exception):
                        document.Close(SaveChanges=False)

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def replace_text(self, path: str, find_text: str, replace_text: str, create_backup: bool) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with office_application("Word.Application", visible=self.settings.office_visible) as word:
            document = None
            try:
                document = word.Documents.Open(str(source), ReadOnly=False)
                current_text = str(document.Content.Text)
                replacement_count = current_text.count(find_text)
                document.Content.Find.Execute(FindText=find_text, ReplaceWith=replace_text, Replace=2)
                document.Save()
                return OperationResult(
                    message="Word text replacement completed",
                    file_path=str(source),
                    backup_path=backup_path,
                    details={
                        "find_text": find_text,
                        "replace_text": replace_text,
                        "replacement_count": replacement_count,
                    },
                )
            finally:
                if document is not None:
                    with suppress(Exception):
                        document.Close(SaveChanges=False)

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def export_pdf(self, path: str, out_path: str) -> OperationResult:
        source = self.resolve_document_path(path)
        target = self.resolve_output_path(out_path, allowed_suffixes=(".pdf",))

        with office_application("Word.Application", visible=self.settings.office_visible) as word:
            document = None
            try:
                document = word.Documents.Open(str(source), ReadOnly=True)
                document.ExportAsFixedFormat(str(target), 17)
                return OperationResult(
                    message="Word document exported to PDF",
                    file_path=str(source),
                    details={"out_path": str(target)},
                )
            finally:
                if document is not None:
                    with suppress(Exception):
                        document.Close(SaveChanges=False)
