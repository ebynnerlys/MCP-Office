from pathlib import Path

from office_ai_mcp.utils.backups import create_backup


def test_create_backup_copies_source_file(tmp_path: Path) -> None:
    source = tmp_path / "sample.docx"
    source.write_text("content", encoding="utf-8")

    backup = create_backup(source, tmp_path / "backups")

    assert backup.exists()
    assert backup.read_text(encoding="utf-8") == "content"
    assert backup != source