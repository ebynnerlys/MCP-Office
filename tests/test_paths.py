from pathlib import Path

import pytest

from office_ai_mcp.utils.paths import ensure_directory, resolve_workspace_path, validate_file_path


def test_resolve_workspace_path_returns_absolute(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.chdir(tmp_path)
    resolved = resolve_workspace_path("demo/file.docx")
    assert resolved == (tmp_path / "demo" / "file.docx").resolve()


def test_validate_file_path_rejects_invalid_suffix(tmp_path: Path) -> None:
    target = tmp_path / "notes.txt"
    target.write_text("hello", encoding="utf-8")

    with pytest.raises(ValueError):
        validate_file_path(target, allowed_suffixes=(".docx",))


def test_validate_file_path_honors_allowed_roots(tmp_path: Path) -> None:
    allowed_root = tmp_path / "allowed"
    blocked_root = tmp_path / "blocked"
    ensure_directory(allowed_root)
    ensure_directory(blocked_root)

    allowed_file = allowed_root / "demo.docx"
    blocked_file = blocked_root / "demo.docx"
    allowed_file.write_text("ok", encoding="utf-8")
    blocked_file.write_text("nope", encoding="utf-8")

    assert validate_file_path(allowed_file, allowed_roots=[str(allowed_root)]) == allowed_file.resolve()

    with pytest.raises(ValueError):
        validate_file_path(blocked_file, allowed_roots=[str(allowed_root)])
