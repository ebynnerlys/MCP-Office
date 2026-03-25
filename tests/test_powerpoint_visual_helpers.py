import pytest
from pathlib import Path
from contextlib import contextmanager
from datetime import datetime

from office_ai_mcp.services.powerpoint_service import (
    STYLE_PRESETS,
    PowerPointService,
    alias_for_value,
    office_color_to_hex,
    parse_office_color,
    resolve_style_preset,
    resolve_named_constant_alias,
)
from office_ai_mcp.config import Settings


def test_parse_office_color_roundtrip_hex() -> None:
    assert office_color_to_hex(parse_office_color("#FF6600")) == "#FF6600"


def test_parse_office_color_roundtrip_csv() -> None:
    assert office_color_to_hex(parse_office_color("255, 102, 0")) == "#FF6600"


def test_parse_office_color_rejects_invalid_text() -> None:
    with pytest.raises(ValueError):
        parse_office_color("orange")


def test_resolve_named_constant_alias_handles_alias_and_numeric() -> None:
    aliases = {"fade": "ppEffectFadeSmoothly"}
    assert resolve_named_constant_alias("fade", aliases) == "ppEffectFadeSmoothly"
    assert resolve_named_constant_alias("12", aliases) == 12


def test_resolve_style_preset_returns_expected_preset() -> None:
    assert resolve_style_preset("executive") == STYLE_PRESETS["executive"]


def test_alias_for_value_returns_matching_name() -> None:
    assert alias_for_value(2, {"medium": 2, "fast": 3}) == "medium"


def test_get_document_properties_reads_extended_builtin_metadata(monkeypatch: pytest.MonkeyPatch) -> None:
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

    class FakePresentation:
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

    service = PowerPointService(Settings(allowed_roots=[]))

    monkeypatch.setattr(service, "resolve_document_path", lambda path: Path("demo.pptx"))

    @contextmanager
    def fake_open(source: Path, *, read_only: bool):
        assert source == Path("demo.pptx")
        assert read_only is True
        yield FakePresentation()

    monkeypatch.setattr(service, "_open_presentation", fake_open)

    result = service.get_document_properties("demo.pptx")

    assert result.built_in["author"] == "Ada"
    assert result.built_in["last_author"] == "Grace"
    assert result.built_in["creation_date"] == "2026-03-01"
    assert result.built_in["last_save_time"] == "2026-03-02"
    assert result.built_in["total_editing_time"] == 42
    assert result.built_in["revision_number"] == "7"
    assert result.custom == {"Client": "Contoso"}


def test_set_document_properties_writes_extended_builtin_metadata(monkeypatch: pytest.MonkeyPatch) -> None:
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

    class FakePresentation:
        def __init__(self) -> None:
            self.BuiltInDocumentProperties = FakeProperties()
            self.saved = False

        def Save(self) -> None:
            self.saved = True

    fake_presentation = FakePresentation()
    service = PowerPointService(Settings(allowed_roots=[]))

    monkeypatch.setattr(service, "resolve_document_path", lambda path: Path("demo.pptx"))
    monkeypatch.setattr(service, "maybe_create_backup", lambda source, create_backup: None)

    @contextmanager
    def fake_open(source: Path, *, read_only: bool):
        assert source == Path("demo.pptx")
        assert read_only is False
        yield fake_presentation

    monkeypatch.setattr(service, "_open_presentation", fake_open)

    result = service.set_document_properties(
        "demo.pptx",
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

    assert fake_presentation.saved is True
    assert fake_presentation.BuiltInDocumentProperties("Last Author").Value == "Synthetic Last Author"
    assert fake_presentation.BuiltInDocumentProperties("Creation Date").Value == datetime(2026, 3, 25, 10, 0, 0)
    assert fake_presentation.BuiltInDocumentProperties("Last Save Time").Value == datetime(2026, 3, 25, 11, 0, 0)
    assert fake_presentation.BuiltInDocumentProperties("Total Editing Time").Value == 123
    assert fake_presentation.BuiltInDocumentProperties("Revision Number").Value == "99"
    assert result.details["updated_properties"]["total_editing_time"] == 123