from __future__ import annotations

from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

import pytest

from office_ai_mcp.config import Settings
from office_ai_mcp.models.requests import WordDocumentPropertiesRequest
from office_ai_mcp.services.word_service import WordService


def test_word_document_properties_request_requires_at_least_one_change() -> None:
    with pytest.raises(ValueError):
        WordDocumentPropertiesRequest(path="demo.docx")


def test_word_document_properties_request_accepts_author_update() -> None:
    request = WordDocumentPropertiesRequest(path="demo.docx", author="Ada Lovelace")
    assert request.author == "Ada Lovelace"


def test_word_document_properties_request_accepts_editing_stats() -> None:
    request = WordDocumentPropertiesRequest(
        path="demo.docx",
        total_editing_time=15,
        creation_date="2026-03-25T10:00:00",
    )
    assert request.total_editing_time == 15
    assert request.creation_date == "2026-03-25T10:00:00"


def test_word_get_document_properties_reads_extended_builtin_metadata(monkeypatch: pytest.MonkeyPatch) -> None:
    class FakeProperty:
        def __init__(self, value: object) -> None:
            self.Value = value

    class FakeProperties:
        def __init__(self, values: dict[str, object]) -> None:
            self._values = values

        def __call__(self, name: str) -> FakeProperty:
            if name not in self._values:
                raise KeyError(name)
            return FakeProperty(self._values[name])

        def Item(self, name: str) -> FakeProperty:
            return self(name)

    class FakeCustomProperties:
        Count = 1

        def __call__(self, index: int) -> object:
            if index != 1:
                raise IndexError(index)
            return type("CustomProperty", (), {"Name": "Client", "Value": "Contoso"})()

    class FakeDocument:
        BuiltInDocumentProperties = FakeProperties(
            {
                "Author": "Ada",
                "Last Author": "Grace",
                "Creation Date": "2026-03-01",
                "Last Save Time": "2026-03-02",
                "Total Editing Time": 42,
                "Revision Number": "7",
            }
        )
        CustomDocumentProperties = FakeCustomProperties()

        def Close(self, SaveChanges: bool = False) -> None:
            return None

    class FakeDocuments:
        def Open(self, path: str, ReadOnly: bool = True) -> FakeDocument:
            assert path == "demo.docx"
            assert ReadOnly is True
            return FakeDocument()

    class FakeWordApplication:
        Documents = FakeDocuments()

    service = WordService(Settings(allowed_roots=[]))
    monkeypatch.setattr(service, "resolve_document_path", lambda path: Path("demo.docx"))

    @contextmanager
    def fake_office_application(prog_id: str, visible: bool = False):
        assert prog_id == "Word.Application"
        yield FakeWordApplication()

    monkeypatch.setattr("office_ai_mcp.services.word_service.office_application", fake_office_application)

    result = service.get_document_properties("demo.docx")

    assert result.built_in["author"] == "Ada"
    assert result.built_in["last_author"] == "Grace"
    assert result.built_in["creation_date"] == "2026-03-01"
    assert result.built_in["last_save_time"] == "2026-03-02"
    assert result.built_in["total_editing_time"] == 42
    assert result.built_in["revision_number"] == "7"
    assert result.custom == {"Client": "Contoso"}


def test_word_set_document_properties_writes_extended_builtin_metadata(monkeypatch: pytest.MonkeyPatch) -> None:
    class FakeProperty:
        def __init__(self) -> None:
            self.Value = None

    class FakeProperties:
        def __init__(self) -> None:
            self.items: dict[str, FakeProperty] = {}

        def __call__(self, name: str) -> FakeProperty:
            return self.items.setdefault(name, FakeProperty())

        def Item(self, name: str) -> FakeProperty:
            return self(name)

    class FakeDocument:
        def __init__(self) -> None:
            self.BuiltInDocumentProperties = FakeProperties()
            self.saved = False

        def Save(self) -> None:
            self.saved = True

        def Close(self, SaveChanges: bool = False) -> None:
            return None

    fake_document = FakeDocument()

    class FakeDocuments:
        def Open(self, path: str, ReadOnly: bool = False) -> FakeDocument:
            assert path == "demo.docx"
            assert ReadOnly is False
            return fake_document

    class FakeWordApplication:
        Documents = FakeDocuments()

    service = WordService(Settings(allowed_roots=[]))
    monkeypatch.setattr(service, "resolve_document_path", lambda path: Path("demo.docx"))
    monkeypatch.setattr(service, "maybe_create_backup", lambda source, create_backup: None)

    @contextmanager
    def fake_office_application(prog_id: str, visible: bool = False):
        assert prog_id == "Word.Application"
        yield FakeWordApplication()

    monkeypatch.setattr("office_ai_mcp.services.word_service.office_application", fake_office_application)

    result = service.set_document_properties(
        "demo.docx",
        author=None,
        title=None,
        subject=None,
        keywords=None,
        comments=None,
        category=None,
        company=None,
        manager=None,
        last_author="Synthetic Last Author",
        creation_date="2026-03-25T10:00:00",
        last_save_time="2026-03-25T11:00:00",
        total_editing_time=123,
        revision_number="99",
        create_backup=False,
    )

    assert fake_document.saved is True
    assert fake_document.BuiltInDocumentProperties("Last Author").Value == "Synthetic Last Author"
    assert fake_document.BuiltInDocumentProperties("Creation Date").Value == datetime(2026, 3, 25, 10, 0, 0)
    assert fake_document.BuiltInDocumentProperties("Last Save Time").Value == datetime(2026, 3, 25, 11, 0, 0)
    assert fake_document.BuiltInDocumentProperties("Total Editing Time").Value == 123
    assert fake_document.BuiltInDocumentProperties("Revision Number").Value == "99"
    assert result.details["updated_properties"]["revision_number"] == "99"