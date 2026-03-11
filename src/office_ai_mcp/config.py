from __future__ import annotations

import os
from functools import lru_cache

from dotenv import load_dotenv
from pydantic import BaseModel, Field


def _split_csv_env(value: str | None) -> list[str]:
    if not value:
        return []
    return [item.strip() for item in value.split(";") if item.strip()]


def _parse_bool(value: str | None, default: bool) -> bool:
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}


class Settings(BaseModel):
    project_name: str = Field(default="Office AI MCP")
    version: str = Field(default="0.1.0")
    log_level: str = Field(default="INFO")
    office_visible: bool = Field(default=False)
    allowed_roots: list[str] = Field(default_factory=list)
    backup_dir: str = Field(default="./backups")
    temp_dir: str = Field(default="./tmp")
    default_transport: str = Field(default="stdio")
    host: str = Field(default="127.0.0.1")
    port: int = Field(default=8000)


@lru_cache(maxsize=1)
def get_settings() -> Settings:
    load_dotenv()
    return Settings(
        project_name=os.getenv("OFFICE_AI_PROJECT_NAME", "Office AI MCP"),
        version=os.getenv("OFFICE_AI_VERSION", "0.1.0"),
        log_level=os.getenv("OFFICE_AI_LOG_LEVEL", "INFO"),
        office_visible=_parse_bool(os.getenv("OFFICE_AI_VISIBLE"), False),
        allowed_roots=_split_csv_env(os.getenv("OFFICE_AI_ALLOWED_ROOTS")),
        backup_dir=os.getenv("OFFICE_AI_BACKUP_DIR", "./backups"),
        temp_dir=os.getenv("OFFICE_AI_TEMP_DIR", "./tmp"),
        default_transport=os.getenv("OFFICE_AI_DEFAULT_TRANSPORT", "stdio"),
        host=os.getenv("OFFICE_AI_HOST", "127.0.0.1"),
        port=int(os.getenv("OFFICE_AI_PORT", "8000")),
    )
