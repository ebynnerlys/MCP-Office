import pytest

from office_ai_mcp.models.requests import (
    PowerPointAddChartSeriesRequest,
    PowerPointAnimationRequest,
    PowerPointAutofitRequest,
    PowerPointBuiltinThemeRequest,
    PowerPointBackgroundGradientRequest,
    PowerPointBulletStyleRequest,
    PowerPointChartAxisScaleRequest,
    PowerPointChartColorsRequest,
    PowerPointChartDataLabelsRequest,
    PowerPointChartExportDataRequest,
    PowerPointChartGridlinesRequest,
    PowerPointChartLinkRequest,
    PowerPointCreatePresentationRequest,
    PowerPointCropImageRequest,
    PowerPointDesignIdeasRequest,
    PowerPointDocumentPropertiesRequest,
    PowerPointDeleteChartSeriesRequest,
    PowerPointExportSlideImagesRequest,
    PowerPointFillPlaceholderRequest,
    PowerPointImageFormatRequest,
    PowerPointInsertBulletsRequest,
    PowerPointMediaPlaybackRequest,
    PowerPointMediaRequest,
    PowerPointMediaTrimRequest,
    PowerPointMasterColorsRequest,
    PowerPointMasterFontsRequest,
    PowerPointParagraphSpacingRequest,
    PowerPointReplacePlaceholderRequest,
    PowerPointProofingLanguageRequest,
    PowerPointShape3DRequest,
    PowerPointShapeGlowRequest,
    PowerPointShapeCropRequest,
    PowerPointShapeMergeRequest,
    PowerPointNotesRequest,
    PowerPointReplaceImageRequest,
    PowerPointSearchRequest,
    PowerPointShapeCollectionRequest,
    PowerPointShapeReflectionRequest,
    PowerPointShapeSearchRequest,
    PowerPointShapeSelectorRequest,
    PowerPointShapeShadowRequest,
    PowerPointShapeSizeRequest,
    PowerPointShapeSoftEdgesRequest,
    PowerPointShapeTextRunsRequest,
    PowerPointSlideLayoutRequest,
    PowerPointSpellcheckPresentationRequest,
    PowerPointSpellcheckSlideRequest,
    PowerPointThemeVariantRequest,
    PowerPointSlideTitleRequest,
    PowerPointTableCellRequest,
    PowerPointTableCellStyleRequest,
    PowerPointTableColumnRequest,
    PowerPointTableColumnStyleRequest,
    PowerPointTableFromCsvRequest,
    PowerPointTableFromExcelRangeRequest,
    PowerPointTableMergeCellsRequest,
    PowerPointTableRowRequest,
    PowerPointTableRowStyleRequest,
    PowerPointTableSortRequest,
    PowerPointTableSplitCellsRequest,
    PowerPointTableStyleRequest,
    PowerPointTextboxMarginsRequest,
    PowerPointTextDirectionRequest,
    PowerPointTextGradientRequest,
    PowerPointTextRangeStyleRequest,
    PowerPointTranslateTextRequest,
)
from office_ai_mcp.services.powerpoint_service import (
    resolve_animation_level,
    resolve_animation_target,
    resolve_chart_axis_kind,
    resolve_chart_data_label_position,
    resolve_chart_legend_position,
    resolve_chart_plot_by,
    resolve_connector_type,
    resolve_chart_type,
    resolve_export_image_format,
    resolve_gradient_style,
    resolve_proofing_language,
    resolve_shape_alignment,
    resolve_shape_distribution,
    resolve_shape_flip,
    resolve_shape_merge_mode,
    resolve_shape_type,
    resolve_shape_resize_mode,
    resolve_shape_z_order,
    resolve_slide_layout,
    resolve_smartart_layout_identifier,
    resolve_text_autofit_mode,
    resolve_text_direction,
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


def test_resolve_shape_alignment_accepts_middle() -> None:
    assert resolve_shape_alignment("middle") == 4


def test_resolve_shape_distribution_accepts_vertical() -> None:
    assert resolve_shape_distribution("vertical") == 1


def test_resolve_shape_flip_accepts_horizontal() -> None:
    assert resolve_shape_flip("horizontal") == 0


def test_resolve_shape_z_order_accepts_bring_forward() -> None:
    assert resolve_shape_z_order("bring_forward") == 2


def test_resolve_shape_resize_mode_accepts_both() -> None:
    assert resolve_shape_resize_mode("both") == "both"


def test_resolve_shape_merge_mode_accepts_union() -> None:
    assert resolve_shape_merge_mode("union") == 2


def test_resolve_chart_legend_position_accepts_alias() -> None:
    assert resolve_chart_legend_position("right") == -4152


def test_resolve_text_direction_accepts_vertical() -> None:
    assert resolve_text_direction("vertical") == "msoTextOrientationVertical"


def test_resolve_gradient_style_accepts_horizontal() -> None:
    assert resolve_gradient_style("horizontal") == "msoGradientHorizontal"


def test_resolve_text_autofit_mode_accepts_alias() -> None:
    assert resolve_text_autofit_mode("shape_to_fit_text") == "ppAutoSizeShapeToFitText"


def test_resolve_proofing_language_accepts_alias() -> None:
    assert resolve_proofing_language("spanish") == "msoLanguageIDSpanish"


def test_resolve_chart_axis_kind_accepts_alias() -> None:
    assert resolve_chart_axis_kind("value") == 2


def test_resolve_chart_data_label_position_accepts_alias() -> None:
    assert resolve_chart_data_label_position("outside_end") == 3


def test_resolve_chart_plot_by_accepts_alias() -> None:
    assert resolve_chart_plot_by("columns") == 2


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


def test_create_presentation_request_defaults_to_title_layout() -> None:
    request = PowerPointCreatePresentationRequest(path="demo.pptx")
    assert request.layout == "title"
    assert request.create_backup is False


def test_search_request_rejects_blank_query() -> None:
    with pytest.raises(ValueError):
        PowerPointSearchRequest(path="demo.pptx", query="   ")


def test_slide_layout_request_requires_layout_text() -> None:
    with pytest.raises(ValueError):
        PowerPointSlideLayoutRequest(path="demo.pptx", slide_index=1, layout="")


def test_master_fonts_request_requires_one_change() -> None:
    with pytest.raises(ValueError):
        PowerPointMasterFontsRequest(path="demo.pptx", master_index=1)


def test_master_colors_request_requires_one_change() -> None:
    with pytest.raises(ValueError):
        PowerPointMasterColorsRequest(path="demo.pptx", master_index=1)


def test_builtin_theme_request_rejects_blank_name() -> None:
    with pytest.raises(ValueError):
        PowerPointBuiltinThemeRequest(path="demo.pptx", theme_name="   ")


def test_design_ideas_request_accepts_optional_slide_and_fallback() -> None:
    request = PowerPointDesignIdeasRequest(path="demo.pptx", slide_index=2, fallback_preset="executive")
    assert request.slide_index == 2
    assert request.fallback_preset == "executive"


def test_theme_variant_request_rejects_blank_variant() -> None:
    with pytest.raises(ValueError):
        PowerPointThemeVariantRequest(path="demo.pptx", master_index=1, variant=" ")


def test_fill_placeholder_request_requires_exactly_one_selector() -> None:
    with pytest.raises(ValueError):
        PowerPointFillPlaceholderRequest(path="demo.pptx", slide_index=1, text="Hola")


def test_fill_placeholder_request_accepts_placeholder_type_selector() -> None:
    request = PowerPointFillPlaceholderRequest(path="demo.pptx", slide_index=1, placeholder_type=2, text="Hola")
    assert request.placeholder_type == 2
    assert request.placeholder_occurrence == 1


def test_text_gradient_request_accepts_defaults() -> None:
    request = PowerPointTextGradientRequest(
        path="demo.pptx",
        slide_index=1,
        shape_index=2,
        start_color="#004E92",
        end_color="#4FC3F7",
    )
    assert request.style == "horizontal"
    assert request.variant == 1


def test_background_gradient_request_restricts_variant_range() -> None:
    with pytest.raises(ValueError):
        PowerPointBackgroundGradientRequest(
            path="demo.pptx",
            slide_index=1,
            start_color="#004E92",
            end_color="#4FC3F7",
            variant=5,
        )


def test_replace_placeholder_request_requires_image_path_for_image() -> None:
    with pytest.raises(ValueError):
        PowerPointReplacePlaceholderRequest(
            path="demo.pptx",
            slide_index=1,
            shape_index=2,
            replacement_kind="image",
        )


def test_slide_title_request_allows_empty_title() -> None:
    request = PowerPointSlideTitleRequest(path="demo.pptx", slide_index=1, title="")
    assert request.title == ""


def test_spellcheck_slide_request_accepts_notes_flag() -> None:
    request = PowerPointSpellcheckSlideRequest(path="demo.pptx", slide_index=1, include_notes=True)
    assert request.include_notes is True


def test_spellcheck_presentation_request_accepts_notes_flag() -> None:
    request = PowerPointSpellcheckPresentationRequest(path="demo.pptx", include_notes=True)
    assert request.include_notes is True


def test_translate_text_request_accepts_languages() -> None:
    request = PowerPointTranslateTextRequest(
        path="demo.pptx",
        slide_index=1,
        shape_index=2,
        target_language="english_us",
        source_language="spanish",
    )
    assert request.target_language == "english_us"
    assert request.source_language == "spanish"


def test_shape_text_runs_request_accepts_shape_index() -> None:
    request = PowerPointShapeTextRunsRequest(path="demo.pptx", slide_index=1, shape_index=2)
    assert request.shape_index == 2


def test_text_range_style_request_requires_one_change() -> None:
    with pytest.raises(ValueError):
        PowerPointTextRangeStyleRequest(path="demo.pptx", slide_index=1, shape_index=2, start=1, length=3)


def test_insert_bullets_request_rejects_blank_items() -> None:
    with pytest.raises(ValueError):
        PowerPointInsertBulletsRequest(
            path="demo.pptx",
            slide_index=1,
            shape_index=2,
            items=["Hola", "   "],
        )


def test_bullet_style_request_requires_one_change() -> None:
    with pytest.raises(ValueError):
        PowerPointBulletStyleRequest(path="demo.pptx", slide_index=1, shape_index=2)


def test_paragraph_spacing_request_requires_one_change() -> None:
    with pytest.raises(ValueError):
        PowerPointParagraphSpacingRequest(path="demo.pptx", slide_index=1, shape_index=2)


def test_textbox_margins_request_requires_one_change() -> None:
    with pytest.raises(ValueError):
        PowerPointTextboxMarginsRequest(path="demo.pptx", slide_index=1, shape_index=2)


def test_text_direction_request_accepts_direction() -> None:
    request = PowerPointTextDirectionRequest(
        path="demo.pptx",
        slide_index=1,
        shape_index=2,
        direction="vertical",
    )
    assert request.direction == "vertical"


def test_autofit_request_accepts_mode_and_word_wrap() -> None:
    request = PowerPointAutofitRequest(
        path="demo.pptx",
        slide_index=1,
        shape_index=2,
        mode="none",
        word_wrap=True,
    )
    assert request.mode == "none"
    assert request.word_wrap is True


def test_proofing_language_request_accepts_language() -> None:
    request = PowerPointProofingLanguageRequest(
        path="demo.pptx",
        slide_index=1,
        shape_index=2,
        language="spanish",
    )
    assert request.language == "spanish"


def test_document_properties_request_requires_at_least_one_change() -> None:
    with pytest.raises(ValueError):
        PowerPointDocumentPropertiesRequest(path="demo.pptx")


def test_document_properties_request_accepts_author_update() -> None:
    request = PowerPointDocumentPropertiesRequest(path="demo.pptx", author="Ada Lovelace")
    assert request.author == "Ada Lovelace"


def test_shape_selector_request_requires_exactly_one_selector() -> None:
    with pytest.raises(ValueError):
        PowerPointShapeSelectorRequest(path="demo.pptx", slide_index=1, shape_index=1, shape_name="Box")


def test_shape_search_request_requires_at_least_one_filter() -> None:
    with pytest.raises(ValueError):
        PowerPointShapeSearchRequest(path="demo.pptx", slide_index=1)


def test_shape_collection_request_rejects_duplicates() -> None:
    with pytest.raises(ValueError):
        PowerPointShapeCollectionRequest(path="demo.pptx", slide_index=1, shape_indexes=[1, 1])


def test_shape_size_request_requires_one_dimension() -> None:
    with pytest.raises(ValueError):
        PowerPointShapeSizeRequest(path="demo.pptx", slide_index=1, shape_index=1)


def test_shape_shadow_request_requires_one_change() -> None:
    with pytest.raises(ValueError):
        PowerPointShapeShadowRequest(path="demo.pptx", slide_index=1, shape_index=1)


def test_shape_glow_request_requires_one_change() -> None:
    with pytest.raises(ValueError):
        PowerPointShapeGlowRequest(path="demo.pptx", slide_index=1, shape_index=1)


def test_shape_reflection_request_requires_one_change() -> None:
    with pytest.raises(ValueError):
        PowerPointShapeReflectionRequest(path="demo.pptx", slide_index=1, shape_index=1)


def test_shape_soft_edges_request_accepts_radius() -> None:
    request = PowerPointShapeSoftEdgesRequest(path="demo.pptx", slide_index=1, shape_index=1, radius=4.5)
    assert request.radius == 4.5


def test_shape_3d_request_requires_one_change() -> None:
    with pytest.raises(ValueError):
        PowerPointShape3DRequest(path="demo.pptx", slide_index=1, shape_index=1)


def test_shape_merge_request_requires_primary_in_shape_indexes() -> None:
    with pytest.raises(ValueError):
        PowerPointShapeMergeRequest(
            path="demo.pptx",
            slide_index=1,
            shape_indexes=[1, 2],
            mode="union",
            primary_shape_index=3,
        )


def test_shape_crop_request_accepts_shape_index() -> None:
    request = PowerPointShapeCropRequest(path="demo.pptx", slide_index=1, shape_index=2)
    assert request.shape_index == 2


def test_replace_image_request_requires_image_path() -> None:
    with pytest.raises(ValueError):
        PowerPointReplaceImageRequest(
            path="demo.pptx",
            slide_index=1,
            shape_index=2,
            image_path="   ",
        )


def test_crop_image_request_requires_one_crop_value() -> None:
    with pytest.raises(ValueError):
        PowerPointCropImageRequest(path="demo.pptx", slide_index=1, shape_index=2)


def test_image_format_request_accepts_midpoint_value() -> None:
    request = PowerPointImageFormatRequest(
        path="demo.pptx",
        slide_index=1,
        shape_index=2,
        value=0.5,
    )
    assert request.value == 0.5


def test_media_request_requires_embedded_or_linked_media() -> None:
    with pytest.raises(ValueError):
        PowerPointMediaRequest(
            path="demo.pptx",
            slide_index=1,
            media_path="demo.mp4",
            link_to_file=False,
            save_with_document=False,
        )


def test_media_playback_request_requires_one_change() -> None:
    with pytest.raises(ValueError):
        PowerPointMediaPlaybackRequest(path="demo.pptx", slide_index=1, shape_index=2)


def test_media_trim_request_rejects_inverted_points() -> None:
    with pytest.raises(ValueError):
        PowerPointMediaTrimRequest(
            path="demo.pptx",
            slide_index=1,
            shape_index=2,
            start_point=5000,
            end_point=4000,
        )


def test_media_trim_request_accepts_single_start_point() -> None:
    request = PowerPointMediaTrimRequest(
        path="demo.pptx",
        slide_index=1,
        shape_index=2,
        start_point=1000,
    )
    assert request.start_point == 1000


def test_table_row_request_allows_append_when_index_missing() -> None:
    request = PowerPointTableRowRequest(path="demo.pptx", slide_index=1, shape_index=2)
    assert request.row_index is None


def test_table_column_request_accepts_specific_index() -> None:
    request = PowerPointTableColumnRequest(path="demo.pptx", slide_index=1, shape_index=2, column_index=3)
    assert request.column_index == 3


def test_table_merge_cells_request_accepts_target_coordinates() -> None:
    request = PowerPointTableMergeCellsRequest(
        path="demo.pptx",
        slide_index=1,
        shape_index=2,
        row_index=1,
        column_index=1,
        merge_to_row_index=1,
        merge_to_column_index=2,
    )
    assert request.merge_to_column_index == 2


def test_table_split_cells_request_rejects_noop_split() -> None:
    with pytest.raises(ValueError):
        PowerPointTableSplitCellsRequest(
            path="demo.pptx",
            slide_index=1,
            shape_index=2,
            row_index=1,
            column_index=1,
            num_rows=1,
            num_columns=1,
        )


def test_table_cell_request_accepts_text() -> None:
    request = PowerPointTableCellRequest(
        path="demo.pptx",
        slide_index=1,
        shape_index=2,
        row_index=1,
        column_index=1,
        text="Hola",
    )
    assert request.text == "Hola"


def test_table_style_request_requires_one_change() -> None:
    with pytest.raises(ValueError):
        PowerPointTableStyleRequest(path="demo.pptx", slide_index=1, shape_index=2)


def test_table_style_request_accepts_banding_toggle() -> None:
    request = PowerPointTableStyleRequest(
        path="demo.pptx",
        slide_index=1,
        shape_index=2,
        horiz_banding=True,
    )
    assert request.horiz_banding is True


def test_table_cell_style_request_requires_one_change() -> None:
    with pytest.raises(ValueError):
        PowerPointTableCellStyleRequest(
            path="demo.pptx",
            slide_index=1,
            shape_index=2,
            row_index=1,
            column_index=1,
        )


def test_table_row_style_request_accepts_fill_color() -> None:
    request = PowerPointTableRowStyleRequest(
        path="demo.pptx",
        slide_index=1,
        shape_index=2,
        row_index=1,
        fill_color="#FF6600",
    )
    assert request.fill_color == "#FF6600"


def test_table_column_style_request_accepts_alignment() -> None:
    request = PowerPointTableColumnStyleRequest(
        path="demo.pptx",
        slide_index=1,
        shape_index=2,
        column_index=1,
        alignment="center",
    )
    assert request.alignment == "center"


def test_table_from_csv_request_requires_csv_path() -> None:
    with pytest.raises(ValueError):
        PowerPointTableFromCsvRequest(
            path="demo.pptx",
            slide_index=1,
            csv_path="   ",
        )


def test_table_from_csv_request_requires_single_character_delimiter() -> None:
    with pytest.raises(ValueError):
        PowerPointTableFromCsvRequest(
            path="demo.pptx",
            slide_index=1,
            csv_path="demo.csv",
            delimiter=";;",
        )


def test_table_from_csv_request_accepts_semicolon_delimiter() -> None:
    request = PowerPointTableFromCsvRequest(
        path="demo.pptx",
        slide_index=1,
        csv_path="demo.csv",
        delimiter=";",
    )
    assert request.delimiter == ";"


def test_table_sort_request_accepts_header_and_descending() -> None:
    request = PowerPointTableSortRequest(
        path="demo.pptx",
        slide_index=1,
        shape_index=2,
        column_index=1,
        descending=True,
        has_header=True,
    )
    assert request.descending is True
    assert request.has_header is True


def test_table_from_excel_range_request_requires_excel_path() -> None:
    with pytest.raises(ValueError):
        PowerPointTableFromExcelRangeRequest(
            path="demo.pptx",
            slide_index=1,
            excel_path=" ",
            sheet="Sheet1",
            cell_range="A1:B2",
        )


def test_table_from_excel_range_request_requires_sheet_name() -> None:
    with pytest.raises(ValueError):
        PowerPointTableFromExcelRangeRequest(
            path="demo.pptx",
            slide_index=1,
            excel_path="demo.xlsx",
            sheet=" ",
            cell_range="A1:B2",
        )


def test_table_from_excel_range_request_accepts_refresh_shape_index() -> None:
    request = PowerPointTableFromExcelRangeRequest(
        path="demo.pptx",
        slide_index=1,
        excel_path="demo.xlsx",
        sheet="Sheet1",
        cell_range="A1:B2",
        shape_index=3,
    )
    assert request.shape_index == 3


def test_chart_axis_scale_request_requires_one_change() -> None:
    with pytest.raises(ValueError):
        PowerPointChartAxisScaleRequest(path="demo.pptx", slide_index=1, shape_index=2)


def test_chart_axis_scale_request_rejects_inverted_range() -> None:
    with pytest.raises(ValueError):
        PowerPointChartAxisScaleRequest(
            path="demo.pptx",
            slide_index=1,
            shape_index=2,
            minimum_scale=10,
            maximum_scale=5,
        )


def test_add_chart_series_request_rejects_mismatched_categories() -> None:
    with pytest.raises(ValueError):
        PowerPointAddChartSeriesRequest(
            path="demo.pptx",
            slide_index=1,
            shape_index=2,
            name="Revenue",
            values=[1, 2],
            categories=["Jan"],
        )


def test_delete_chart_series_request_accepts_series_index() -> None:
    request = PowerPointDeleteChartSeriesRequest(
        path="demo.pptx",
        slide_index=1,
        shape_index=2,
        series_index=1,
    )
    assert request.series_index == 1


def test_chart_data_labels_request_requires_one_change() -> None:
    with pytest.raises(ValueError):
        PowerPointChartDataLabelsRequest(path="demo.pptx", slide_index=1, shape_index=2)


def test_chart_gridlines_request_requires_one_change() -> None:
    with pytest.raises(ValueError):
        PowerPointChartGridlinesRequest(path="demo.pptx", slide_index=1, shape_index=2)


def test_chart_colors_request_requires_one_change() -> None:
    with pytest.raises(ValueError):
        PowerPointChartColorsRequest(path="demo.pptx", slide_index=1, shape_index=2)


def test_chart_colors_request_rejects_palette_with_series_index() -> None:
    with pytest.raises(ValueError):
        PowerPointChartColorsRequest(
            path="demo.pptx",
            slide_index=1,
            shape_index=2,
            series_index=1,
            series_fill_colors=["#FF0000"],
        )


def test_chart_export_data_request_accepts_csv_output() -> None:
    request = PowerPointChartExportDataRequest(
        path="demo.pptx",
        slide_index=1,
        shape_index=2,
        out_path="chart.csv",
        export_format="csv",
    )
    assert request.out_path == "chart.csv"
    assert request.export_format == "csv"


def test_chart_link_request_requires_excel_path() -> None:
    with pytest.raises(ValueError):
        PowerPointChartLinkRequest(
            path="demo.pptx",
            slide_index=1,
            shape_index=2,
            excel_path=" ",
            sheet="Sheet1",
            cell_range="A1:D5",
        )


def test_chart_link_request_accepts_plot_by() -> None:
    request = PowerPointChartLinkRequest(
        path="demo.pptx",
        slide_index=1,
        shape_index=2,
        excel_path="demo.xlsx",
        sheet="Sheet1",
        cell_range="A1:D5",
        plot_by="rows",
    )
    assert request.plot_by == "rows"