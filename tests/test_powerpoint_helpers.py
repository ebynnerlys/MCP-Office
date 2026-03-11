import pytest

from office_ai_mcp.models.requests import PowerPointAnimationRequest, PowerPointExportSlideImagesRequest, PowerPointNotesRequest
from office_ai_mcp.services.powerpoint_service import (
    resolve_animation_level,
    resolve_animation_target,
    resolve_chart_legend_position,
    resolve_connector_type,
    resolve_chart_type,
    resolve_export_image_format,
    resolve_shape_type,
    resolve_slide_layout,
    resolve_smartart_layout_identifier,
)


def test_resolve_slide_layout_accepts_named_layout() -> None:
    assert resolve_slide_layout("title_and_text") == 2


def test_resolve_slide_layout_accepts_numeric_layout() -> None:
    assert resolve_slide_layout("12") == 12


def test_resolve_export_image_format_accepts_png() -> None:
    assert resolve_export_image_format("png") == ("PNG", ".png")


def test_resolve_chart_type_accepts_alias() -> None:
    assert resolve_chart_type("column_clustered") == 51


def test_resolve_smartart_layout_identifier_accepts_alias() -> None:
    assert resolve_smartart_layout_identifier("basic_process") == "urn:microsoft.com/office/officeart/2005/8/layout/process1"


def test_resolve_shape_type_accepts_alias() -> None:
    assert resolve_shape_type("rounded_rectangle") == 5


def test_resolve_connector_type_accepts_alias() -> None:
    assert resolve_connector_type("elbow") == 2


def test_resolve_chart_legend_position_accepts_alias() -> None:
    assert resolve_chart_legend_position("right") == -4152


def test_resolve_animation_level_accepts_text_alias() -> None:
    assert resolve_animation_level("all_text_levels") == 1


def test_resolve_animation_target_accepts_cell_alias() -> None:
    assert resolve_animation_target("cell") == "table_cell"


def test_export_slide_images_request_requires_both_dimensions() -> None:
    with pytest.raises(ValueError):
        PowerPointExportSlideImagesRequest(path="demo.pptx", out_dir="out", width=1280)


def test_animation_request_requires_coordinates_for_table_cell() -> None:
    with pytest.raises(ValueError):
        PowerPointAnimationRequest(
            path="demo.pptx",
            slide_index=1,
            shape_index=1,
            target_kind="table_cell",
        )


def test_animation_request_rejects_text_level_for_table_targets() -> None:
    with pytest.raises(ValueError):
        PowerPointAnimationRequest(
            path="demo.pptx",
            slide_index=1,
            shape_index=1,
            target_kind="table_cell",
            row_index=1,
            column_index=1,
            animation_level="all_text_levels",
        )


def test_notes_request_allows_empty_text_to_clear_notes() -> None:
    request = PowerPointNotesRequest(path="demo.pptx", slide_index=1, text="", append=False)
    assert request.text == ""
    assert request.append is False