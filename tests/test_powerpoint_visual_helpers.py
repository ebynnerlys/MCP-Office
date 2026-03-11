import pytest

from office_ai_mcp.services.powerpoint_service import (
    STYLE_PRESETS,
    alias_for_value,
    office_color_to_hex,
    parse_office_color,
    resolve_style_preset,
    resolve_named_constant_alias,
)


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