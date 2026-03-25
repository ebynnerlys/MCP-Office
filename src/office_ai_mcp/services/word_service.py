from __future__ import annotations

from contextlib import suppress
from datetime import datetime

from tenacity import retry, stop_after_attempt, wait_fixed

from office_ai_mcp.config import Settings
from office_ai_mcp.models.responses import (
    DocumentPropertiesResult,
    OperationResult,
    SectionSummary,
    WordStructureResult,
)
from office_ai_mcp.services.base import OfficeService
from office_ai_mcp.utils.com_cleanup import office_application

BUILTIN_DOCUMENT_PROPERTY_NAMES = {
    "author": "Author",
    "title": "Title",
    "subject": "Subject",
    "keywords": "Keywords",
    "comments": "Comments",
    "category": "Category",
    "company": "Company",
    "manager": "Manager",
    "last_author": "Last Author",
    "creation_date": "Creation Date",
    "last_save_time": "Last Save Time",
    "total_editing_time": "Total Editing Time",
    "revision_number": "Revision Number",
}

WRITABLE_BUILTIN_DOCUMENT_PROPERTY_NAMES = {
    "author": "Author",
    "title": "Title",
    "subject": "Subject",
    "keywords": "Keywords",
    "comments": "Comments",
    "category": "Category",
    "company": "Company",
    "manager": "Manager",
    "last_author": "Last Author",
    "creation_date": "Creation Date",
    "last_save_time": "Last Save Time",
    "total_editing_time": "Total Editing Time",
    "revision_number": "Revision Number",
}


def normalize_office_value(value: object) -> object:
    if isinstance(value, tuple):
        return [normalize_office_value(item) for item in value]
    if isinstance(value, list):
        return [normalize_office_value(item) for item in value]
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    try:
        return [normalize_office_value(item) for item in value]  # type: ignore[arg-type]
    except Exception:
        return str(value)


def coerce_document_property_value(key: str, value: object) -> object:
    if value is None:
        return value
    if key not in {"creation_date", "last_save_time"}:
        return value
    if not isinstance(value, str):
        return value
    normalized = value[:-1] + "+00:00" if value.endswith("Z") else value
    return datetime.fromisoformat(normalized)


class WordService(OfficeService):
    allowed_suffixes = (".doc", ".docx", ".docm")

    def __init__(self, settings: Settings) -> None:
        super().__init__(settings)

    def _get_document_property(self, properties: object, property_name: str) -> object | None:
        with suppress(Exception):
            return normalize_office_value(properties(property_name).Value)
        with suppress(Exception):
            return normalize_office_value(properties.Item(property_name).Value)
        return None

    def _set_document_property(self, properties: object, property_name: str, value: object) -> None:
        with suppress(Exception):
            properties(property_name).Value = value
            return
        properties.Item(property_name).Value = value

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
    def get_document_properties(self, path: str) -> DocumentPropertiesResult:
        source = self.resolve_document_path(path)

        with office_application("Word.Application", visible=self.settings.office_visible) as word:
            document = None
            try:
                document = word.Documents.Open(str(source), ReadOnly=True)
                built_in = {
                    key: self._get_document_property(document.BuiltInDocumentProperties, property_name)
                    for key, property_name in BUILTIN_DOCUMENT_PROPERTY_NAMES.items()
                }
                custom: dict[str, object] = {}
                with suppress(Exception):
                    properties = document.CustomDocumentProperties
                    for index in range(1, int(properties.Count) + 1):
                        item = properties(index)
                        custom[str(item.Name)] = normalize_office_value(item.Value)

                return DocumentPropertiesResult(file_path=str(source), built_in=built_in, custom=custom)
            finally:
                if document is not None:
                    with suppress(Exception):
                        document.Close(SaveChanges=False)

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_document_properties(
        self,
        path: str,
        *,
        author: str | None,
        title: str | None,
        subject: str | None,
        keywords: str | None,
        comments: str | None,
        category: str | None,
        company: str | None,
        manager: str | None,
        last_author: str | None,
        creation_date: str | None,
        last_save_time: str | None,
        total_editing_time: int | None,
        revision_number: str | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)
        updates = {
            "author": author,
            "title": title,
            "subject": subject,
            "keywords": keywords,
            "comments": comments,
            "category": category,
            "company": company,
            "manager": manager,
            "last_author": last_author,
            "creation_date": creation_date,
            "last_save_time": last_save_time,
            "total_editing_time": total_editing_time,
            "revision_number": revision_number,
        }

        with office_application("Word.Application", visible=self.settings.office_visible) as word:
            document = None
            try:
                document = word.Documents.Open(str(source), ReadOnly=False)
                properties = document.BuiltInDocumentProperties
                applied_updates: dict[str, object] = {}
                for key, value in updates.items():
                    if value is None:
                        continue
                    property_name = WRITABLE_BUILTIN_DOCUMENT_PROPERTY_NAMES[key]
                    self._set_document_property(properties, property_name, coerce_document_property_value(key, value))
                    applied_updates[key] = value

                document.Save()
                return OperationResult(
                    message="Word document properties updated",
                    file_path=str(source),
                    backup_path=backup_path,
                    details={"updated_properties": applied_updates},
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
