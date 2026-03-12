from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from office_ai_mcp.models.requests import (
    DocumentPathRequest,
    ExportPdfRequest,
    PowerPointAddSlideRequest,
    PowerPointAddChartSeriesRequest,
    PowerPointAddChartRequest,
    PowerPointAddConnectorRequest,
    PowerPointAddShapeRequest,
    PowerPointAddSmartArtRequest,
    PowerPointAddTableRequest,
    PowerPointAnimationRequest,
    PowerPointAutofitRequest,
    PowerPointBackgroundRequest,
    PowerPointBackgroundGradientRequest,
    PowerPointBuiltinThemeRequest,
    PowerPointBulletStyleRequest,
    PowerPointChartLayoutRequest,
    PowerPointChartAxisScaleRequest,
    PowerPointChartLinkRequest,
    PowerPointChartColorsRequest,
    PowerPointChartDataRequest,
    PowerPointChartDataLabelsRequest,
    PowerPointChartExportDataRequest,
    PowerPointChartSeriesOrderRequest,
    PowerPointChartSeriesStyleRequest,
    PowerPointChartTypeChangeRequest,
    PowerPointChartTitleRequest,
    PowerPointConnectShapesRequest,
    PowerPointCreatePresentationRequest,
    PowerPointCropImageRequest,
    PowerPointDesignIdeasRequest,
    PowerPointDocumentPropertiesRequest,
    PowerPointDeleteChartSeriesRequest,
    PowerPointExportSlideImagesRequest,
    PowerPointFillPlaceholderRequest,
    PowerPointImageFormatRequest,
    PowerPointImageRequest,
    PowerPointInsertBulletsRequest,
    PowerPointMediaPlaybackRequest,
    PowerPointMediaRequest,
    PowerPointMediaTrimRequest,
    PowerPointMasterBackgroundRequest,
    PowerPointMasterColorsRequest,
    PowerPointMasterFontsRequest,
    PowerPointMasterRequest,
    PowerPointNotesRequest,
    PowerPointOptionalMasterRequest,
    PowerPointParagraphSpacingRequest,
    PowerPointPresetRequest,
    PowerPointProofingLanguageRequest,
    PowerPointReplacePlaceholderRequest,
    PowerPointReplaceImageRequest,
    PowerPointShapeAspectRatioRequest,
    PowerPointShapeCollectionRequest,
    PowerPointShapeFlipRequest,
    PowerPointSearchRequest,
    PowerPointShape3DRequest,
    PowerPointShapeCropRequest,
    PowerPointShapeRequest,
    PowerPointShapeGlowRequest,
    PowerPointShapeMergeRequest,
    PowerPointShapeNameRequest,
    PowerPointShapePositionRequest,
    PowerPointShapeReflectionRequest,
    PowerPointShapeRotateRequest,
    PowerPointShapeSearchRequest,
    PowerPointShapeSelectorRequest,
    PowerPointShapeShadowRequest,
    PowerPointShapeTextRunsRequest,
    PowerPointShapesAlignRequest,
    PowerPointShapesDistributeRequest,
    PowerPointShapesResizeRequest,
    PowerPointShapeFillRequest,
    PowerPointShapeLineRequest,
    PowerPointShapeSizeRequest,
    PowerPointShapeSoftEdgesRequest,
    PowerPointTextDirectionRequest,
    PowerPointTextboxMarginsRequest,
    PowerPointTranslateTextRequest,
    PowerPointShapeVisibilityRequest,
    PowerPointSlideLayoutRequest,
    PowerPointSlideNameRequest,
    PowerPointSlidePositionRequest,
    PowerPointSlideRequest,
    PowerPointSpellcheckPresentationRequest,
    PowerPointSpellcheckSlideRequest,
    PowerPointSlideTitleRequest,
    PowerPointSmartArtNodeRequest,
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
    PowerPointTextReplaceRequest,
    PowerPointTextGradientRequest,
    PowerPointTextRangeStyleRequest,
    PowerPointTextStyleRequest,
    PowerPointThemeRequest,
    PowerPointThemeVariantRequest,
    PowerPointTransitionRequest,
    ReplaceTextRequest,
    SaveAsRequest,
)
from office_ai_mcp.services.powerpoint_service import PowerPointService


def register_powerpoint_tools(mcp: FastMCP, service: PowerPointService) -> None:
    @mcp.tool(
        name="ppt_create_presentation",
        description="Create a new PowerPoint presentation file from scratch with an initial slide.",
    )
    def ppt_create_presentation(
        path: str,
        layout: str = "title",
        title: str | None = None,
        body_text: str | None = None,
    ) -> dict[str, object]:
        """Create a new PowerPoint presentation on disk."""
        request = PowerPointCreatePresentationRequest(
            path=path,
            layout=layout,
            title=title,
            body_text=body_text,
        )
        return service.create_presentation(
            path=request.path,
            layout=request.layout,
            title=request.title,
            body_text=request.body_text,
        ).model_dump()

    @mcp.tool(
        name="ppt_list_slides",
        description="List slides in a PowerPoint presentation with titles and shape counts.",
    )
    def ppt_list_slides(path: str) -> dict[str, object]:
        """Summarize the slides present in a PowerPoint file."""
        return service.list_slides(path).model_dump()

    @mcp.tool(
        name="ppt_save",
        description="Save changes in the current PowerPoint presentation.",
    )
    def ppt_save(path: str, create_backup: bool = False) -> dict[str, object]:
        """Persist the current presentation to disk."""
        request = DocumentPathRequest(path=path, create_backup=create_backup)
        return service.save(path=request.path, create_backup=request.create_backup).model_dump()

    @mcp.tool(
        name="ppt_save_copy",
        description="Save a copy of the presentation without changing the active source path.",
    )
    def ppt_save_copy(path: str, out_path: str) -> dict[str, object]:
        """Write a copy of the presentation to another PowerPoint file."""
        request = SaveAsRequest(path=path, out_path=out_path)
        return service.save_copy(path=request.path, out_path=request.out_path).model_dump()

    @mcp.tool(
        name="ppt_get_document_properties",
        description="Read built-in and custom document properties from a PowerPoint file.",
    )
    def ppt_get_document_properties(path: str) -> dict[str, object]:
        """Inspect built-in and custom PowerPoint document properties."""
        request = DocumentPathRequest(path=path, create_backup=False)
        return service.get_document_properties(path=request.path).model_dump()

    @mcp.tool(
        name="ppt_set_document_properties",
        description="Update built-in PowerPoint document properties such as author, title, subject, and keywords.",
    )
    def ppt_set_document_properties(
        path: str,
        author: str | None = None,
        title: str | None = None,
        subject: str | None = None,
        keywords: str | None = None,
        comments: str | None = None,
        category: str | None = None,
        company: str | None = None,
        manager: str | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Write selected built-in document properties."""
        request = PowerPointDocumentPropertiesRequest(
            path=path,
            author=author,
            title=title,
            subject=subject,
            keywords=keywords,
            comments=comments,
            category=category,
            company=company,
            manager=manager,
            create_backup=create_backup,
        )
        return service.set_document_properties(
            path=request.path,
            author=request.author,
            title=request.title,
            subject=request.subject,
            keywords=request.keywords,
            comments=request.comments,
            category=request.category,
            company=request.company,
            manager=request.manager,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_get_file_links",
        description="Inspect external file links referenced by the PowerPoint presentation.",
    )
    def ppt_get_file_links(path: str) -> dict[str, object]:
        """Return external links detected in the presentation file."""
        request = DocumentPathRequest(path=path, create_backup=False)
        return service.get_file_links(path=request.path).model_dump()

    @mcp.tool(
        name="ppt_duplicate_slide",
        description="Duplicate an existing slide and insert the copy immediately after it.",
    )
    def ppt_duplicate_slide(path: str, slide_index: int, create_backup: bool = True) -> dict[str, object]:
        """Duplicate one slide in a presentation."""
        request = PowerPointSlideRequest(path=path, slide_index=slide_index, create_backup=create_backup)
        return service.duplicate_slide(
            path=request.path,
            slide_index=request.slide_index,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_delete_slide",
        description="Delete a slide from the PowerPoint presentation.",
    )
    def ppt_delete_slide(path: str, slide_index: int, create_backup: bool = True) -> dict[str, object]:
        """Remove one slide from a presentation."""
        request = PowerPointSlideRequest(path=path, slide_index=slide_index, create_backup=create_backup)
        return service.delete_slide(
            path=request.path,
            slide_index=request.slide_index,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_move_slide",
        description="Change the order of a slide by moving it to another position.",
    )
    def ppt_move_slide(
        path: str,
        slide_index: int,
        position: int,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Move one slide to a new 1-based position."""
        request = PowerPointSlidePositionRequest(
            path=path,
            slide_index=slide_index,
            position=position,
            create_backup=create_backup,
        )
        return service.move_slide(
            path=request.path,
            slide_index=request.slide_index,
            position=request.position,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_hide_slide",
        description="Hide a slide so it is skipped during slideshow playback.",
    )
    def ppt_hide_slide(path: str, slide_index: int, create_backup: bool = True) -> dict[str, object]:
        """Mark one slide as hidden."""
        request = PowerPointSlideRequest(path=path, slide_index=slide_index, create_backup=create_backup)
        return service.hide_slide(
            path=request.path,
            slide_index=request.slide_index,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_unhide_slide",
        description="Unhide a slide so it is shown again during slideshow playback.",
    )
    def ppt_unhide_slide(path: str, slide_index: int, create_backup: bool = True) -> dict[str, object]:
        """Clear the hidden state of one slide."""
        request = PowerPointSlideRequest(path=path, slide_index=slide_index, create_backup=create_backup)
        return service.unhide_slide(
            path=request.path,
            slide_index=request.slide_index,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_slide_name",
        description="Assign or update the internal PowerPoint name of a slide.",
    )
    def ppt_set_slide_name(
        path: str,
        slide_index: int,
        name: str,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Rename one slide for more stable automation flows."""
        request = PowerPointSlideNameRequest(
            path=path,
            slide_index=slide_index,
            name=name,
            create_backup=create_backup,
        )
        return service.set_slide_name(
            path=request.path,
            slide_index=request.slide_index,
            name=request.name,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_get_slide_metadata",
        description="Read extended metadata for a slide, including id, name, layout, section, and hidden state.",
    )
    def ppt_get_slide_metadata(path: str, slide_index: int) -> dict[str, object]:
        """Return rich metadata about one slide."""
        request = PowerPointSlideRequest(path=path, slide_index=slide_index, create_backup=False)
        return service.get_slide_metadata(path=request.path, slide_index=request.slide_index).model_dump()

    @mcp.tool(
        name="ppt_get_slide_summary_extended",
        description="Return an extended slide summary with text, notes, shapes, tables, charts, SmartArt, and animations.",
    )
    def ppt_get_slide_summary_extended(path: str, slide_index: int) -> dict[str, object]:
        """Collect a detailed slide summary in one call."""
        request = PowerPointSlideRequest(path=path, slide_index=slide_index, create_backup=False)
        return service.get_slide_summary_extended(path=request.path, slide_index=request.slide_index).model_dump()

    @mcp.tool(
        name="ppt_list_masters",
        description="List presentation slide masters and their associated theme names and layout counts.",
    )
    def ppt_list_masters(path: str) -> dict[str, object]:
        """Enumerate the slide masters available in the presentation."""
        request = DocumentPathRequest(path=path, create_backup=False)
        return service.list_masters(path=request.path).model_dump()

    @mcp.tool(
        name="ppt_get_master_details",
        description="Read one slide master with theme colors, fonts, layouts, placeholders, and variants.",
    )
    def ppt_get_master_details(path: str, master_index: int) -> dict[str, object]:
        """Return detailed theme information for one slide master."""
        request = PowerPointMasterRequest(path=path, master_index=master_index, create_backup=False)
        return service.get_master_details(path=request.path, master_index=request.master_index).model_dump()

    @mcp.tool(
        name="ppt_list_layouts",
        description="List the layouts available in the presentation across its slide masters.",
    )
    def ppt_list_layouts(path: str) -> dict[str, object]:
        """Enumerate slide layouts exposed by the presentation."""
        request = DocumentPathRequest(path=path, create_backup=False)
        return service.list_layouts(path=request.path).model_dump()

    @mcp.tool(
        name="ppt_get_slide_layout",
        description="Inspect the current layout of a specific slide.",
    )
    def ppt_get_slide_layout(path: str, slide_index: int) -> dict[str, object]:
        """Return layout details for one slide."""
        request = PowerPointSlideRequest(path=path, slide_index=slide_index, create_backup=False)
        return service.get_slide_layout(path=request.path, slide_index=request.slide_index).model_dump()

    @mcp.tool(
        name="ppt_apply_layout",
        description="Apply a named or numeric slide layout to an existing slide.",
    )
    def ppt_apply_layout(
        path: str,
        slide_index: int,
        layout: str,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Change the layout of one slide."""
        request = PowerPointSlideLayoutRequest(
            path=path,
            slide_index=slide_index,
            layout=layout,
            create_backup=create_backup,
        )
        return service.apply_layout(
            path=request.path,
            slide_index=request.slide_index,
            layout=request.layout,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_master_background",
        description="Update the background color of a specific slide master.",
    )
    def ppt_set_master_background(
        path: str,
        master_index: int,
        color: str,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Set the background fill color on one slide master."""
        request = PowerPointMasterBackgroundRequest(
            path=path,
            master_index=master_index,
            color=color,
            create_backup=create_backup,
        )
        return service.set_master_background(
            path=request.path,
            master_index=request.master_index,
            color=request.color,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_master_fonts",
        description="Configure title and body fonts for a specific slide master and its layouts.",
    )
    def ppt_set_master_fonts(
        path: str,
        master_index: int,
        title_font_name: str | None = None,
        title_font_size: float | None = None,
        body_font_name: str | None = None,
        body_font_size: float | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Update theme typography on one slide master."""
        request = PowerPointMasterFontsRequest(
            path=path,
            master_index=master_index,
            title_font_name=title_font_name,
            title_font_size=title_font_size,
            body_font_name=body_font_name,
            body_font_size=body_font_size,
            create_backup=create_backup,
        )
        return service.set_master_fonts(
            path=request.path,
            master_index=request.master_index,
            title_font_name=request.title_font_name,
            title_font_size=request.title_font_size,
            body_font_name=request.body_font_name,
            body_font_size=request.body_font_size,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_master_colors",
        description="Update the main color palette of a specific slide master.",
    )
    def ppt_set_master_colors(
        path: str,
        master_index: int,
        background_color: str | None = None,
        title_text_color: str | None = None,
        body_text_color: str | None = None,
        accent_color: str | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Update background, title, body, and accent colors on one master."""
        request = PowerPointMasterColorsRequest(
            path=path,
            master_index=master_index,
            background_color=background_color,
            title_text_color=title_text_color,
            body_text_color=body_text_color,
            accent_color=accent_color,
            create_backup=create_backup,
        )
        return service.set_master_colors(
            path=request.path,
            master_index=request.master_index,
            background_color=request.background_color,
            title_text_color=request.title_text_color,
            body_text_color=request.body_text_color,
            accent_color=request.accent_color,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_reset_slide_to_layout",
        description="Reset a slide to its current layout, restoring placeholders where possible.",
    )
    def ppt_reset_slide_to_layout(
        path: str,
        slide_index: int,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Reset one slide back to its assigned layout."""
        request = PowerPointSlideRequest(path=path, slide_index=slide_index, create_backup=create_backup)
        return service.reset_slide_to_layout(
            path=request.path,
            slide_index=request.slide_index,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_list_placeholders",
        description="Enumerate the placeholders present on a specific slide.",
    )
    def ppt_list_placeholders(path: str, slide_index: int) -> dict[str, object]:
        """List placeholder shapes available on one slide."""
        request = PowerPointSlideRequest(path=path, slide_index=slide_index, create_backup=False)
        return service.list_placeholders(path=request.path, slide_index=request.slide_index).model_dump()

    @mcp.tool(
        name="ppt_fill_placeholder",
        description="Fill a placeholder selected by shape index, name, or placeholder type.",
    )
    def ppt_fill_placeholder(
        path: str,
        slide_index: int,
        text: str,
        shape_index: int | None = None,
        shape_name: str | None = None,
        placeholder_type: int | None = None,
        placeholder_occurrence: int = 1,
        text_color: str | None = None,
        font_name: str | None = None,
        font_size: float | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Write content into a placeholder and optionally style its text."""
        request = PowerPointFillPlaceholderRequest(
            path=path,
            slide_index=slide_index,
            text=text,
            shape_index=shape_index,
            shape_name=shape_name,
            placeholder_type=placeholder_type,
            placeholder_occurrence=placeholder_occurrence,
            text_color=text_color,
            font_name=font_name,
            font_size=font_size,
            create_backup=create_backup,
        )
        return service.fill_placeholder(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            shape_name=request.shape_name,
            placeholder_type=request.placeholder_type,
            placeholder_occurrence=request.placeholder_occurrence,
            text=request.text,
            text_color=request.text_color,
            font_name=request.font_name,
            font_size=request.font_size,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_replace_placeholder_with_shape",
        description="Replace a placeholder with a textbox, image, or shape while preserving its geometry.",
    )
    def ppt_replace_placeholder_with_shape(
        path: str,
        slide_index: int,
        shape_index: int | None = None,
        shape_name: str | None = None,
        placeholder_type: int | None = None,
        placeholder_occurrence: int = 1,
        replacement_kind: str = "textbox",
        text: str | None = None,
        image_path: str | None = None,
        shape_type: str = "rectangle",
        fill_color: str | None = None,
        line_color: str | None = None,
        text_color: str | None = None,
        font_name: str | None = None,
        font_size: float | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Swap one placeholder for a final content shape."""
        request = PowerPointReplacePlaceholderRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            shape_name=shape_name,
            placeholder_type=placeholder_type,
            placeholder_occurrence=placeholder_occurrence,
            replacement_kind=replacement_kind,
            text=text,
            image_path=image_path,
            shape_type=shape_type,
            fill_color=fill_color,
            line_color=line_color,
            text_color=text_color,
            font_name=font_name,
            font_size=font_size,
            create_backup=create_backup,
        )
        return service.replace_placeholder_with_shape(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            shape_name=request.shape_name,
            placeholder_type=request.placeholder_type,
            placeholder_occurrence=request.placeholder_occurrence,
            replacement_kind=request.replacement_kind,
            text=request.text,
            image_path=request.image_path,
            shape_type=request.shape_type,
            fill_color=request.fill_color,
            line_color=request.line_color,
            text_color=request.text_color,
            font_name=request.font_name,
            font_size=request.font_size,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_restore_placeholder",
        description="Restore deleted placeholders on a slide by resetting it to its assigned layout.",
    )
    def ppt_restore_placeholder(
        path: str,
        slide_index: int,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Restore missing placeholders on one slide."""
        request = PowerPointSlideRequest(path=path, slide_index=slide_index, create_backup=create_backup)
        return service.restore_placeholder(
            path=request.path,
            slide_index=request.slide_index,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_find_text",
        description="Search visible slide text across the entire presentation.",
    )
    def ppt_find_text(path: str, query: str) -> dict[str, object]:
        """Find matching text across all slide shapes."""
        request = PowerPointSearchRequest(path=path, query=query)
        return service.find_text(path=request.path, query=request.query).model_dump()

    @mcp.tool(
        name="ppt_replace_text_all",
        description="Replace text across the whole presentation, not only within a single slide.",
    )
    def ppt_replace_text_all(
        path: str,
        find: str,
        replace: str,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Replace slide text everywhere in the presentation."""
        request = ReplaceTextRequest(
            path=path,
            find_text=find,
            replace_text=replace,
            create_backup=create_backup,
        )
        return service.replace_text_all(
            path=request.path,
            find_text=request.find_text,
            replace_text=request.replace_text,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_slide_title",
        description="Update only the title of a specific slide.",
    )
    def ppt_set_slide_title(
        path: str,
        slide_index: int,
        title: str,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Change the title text of one slide."""
        request = PowerPointSlideTitleRequest(
            path=path,
            slide_index=slide_index,
            title=title,
            create_backup=create_backup,
        )
        return service.set_slide_title(
            path=request.path,
            slide_index=request.slide_index,
            title=request.title,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_get_shape_text_runs",
        description="Return formatted text runs for a specific text shape.",
    )
    def ppt_get_shape_text_runs(path: str, slide_index: int, shape_index: int) -> dict[str, object]:
        """Inspect text runs within one shape."""
        request = PowerPointShapeTextRunsRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            create_backup=False,
        )
        return service.get_shape_text_runs(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_text_range_style",
        description="Apply text formatting to a character range within a text shape.",
    )
    def ppt_set_text_range_style(
        path: str,
        slide_index: int,
        shape_index: int,
        start: int,
        length: int,
        text: str | None = None,
        font_name: str | None = None,
        font_size: float | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        color: str | None = None,
        alignment: str | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Format a specific character range inside a shape."""
        request = PowerPointTextRangeStyleRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            start=start,
            length=length,
            text=text,
            font_name=font_name,
            font_size=font_size,
            bold=bold,
            italic=italic,
            underline=underline,
            color=color,
            alignment=alignment,
            create_backup=create_backup,
        )
        return service.set_text_range_style(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            start=request.start,
            length=request.length,
            text=request.text,
            font_name=request.font_name,
            font_size=request.font_size,
            bold=request.bold,
            italic=request.italic,
            underline=request.underline,
            color=request.color,
            alignment=request.alignment,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_insert_bullets",
        description="Create or append a bulleted list inside a text shape.",
    )
    def ppt_insert_bullets(
        path: str,
        slide_index: int,
        shape_index: int,
        items: list[str],
        level: int = 1,
        append: bool = False,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Insert bullet list items into one shape."""
        request = PowerPointInsertBulletsRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            items=items,
            level=level,
            append=append,
            create_backup=create_backup,
        )
        return service.insert_bullets(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            items=request.items,
            level=request.level,
            append=request.append,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_bullet_style",
        description="Adjust bullet visibility, level, symbol, and indentation for one paragraph or the whole text box.",
    )
    def ppt_set_bullet_style(
        path: str,
        slide_index: int,
        shape_index: int,
        paragraph_index: int | None = None,
        visible: bool | None = None,
        level: int | None = None,
        bullet_character: str | None = None,
        font_name: str | None = None,
        color: str | None = None,
        relative_size: float | None = None,
        left_margin: float | None = None,
        first_margin: float | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Configure bullet styling and indentation."""
        request = PowerPointBulletStyleRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            paragraph_index=paragraph_index,
            visible=visible,
            level=level,
            bullet_character=bullet_character,
            font_name=font_name,
            color=color,
            relative_size=relative_size,
            left_margin=left_margin,
            first_margin=first_margin,
            create_backup=create_backup,
        )
        return service.set_bullet_style(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            paragraph_index=request.paragraph_index,
            visible=request.visible,
            level=request.level,
            bullet_character=request.bullet_character,
            font_name=request.font_name,
            color=request.color,
            relative_size=request.relative_size,
            left_margin=request.left_margin,
            first_margin=request.first_margin,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_paragraph_spacing",
        description="Set paragraph spacing before, after, and within lines for a text shape.",
    )
    def ppt_set_paragraph_spacing(
        path: str,
        slide_index: int,
        shape_index: int,
        paragraph_index: int | None = None,
        space_before: float | None = None,
        space_after: float | None = None,
        space_within: float | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Adjust paragraph spacing on one paragraph or the full text box."""
        request = PowerPointParagraphSpacingRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            paragraph_index=paragraph_index,
            space_before=space_before,
            space_after=space_after,
            space_within=space_within,
            create_backup=create_backup,
        )
        return service.set_paragraph_spacing(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            paragraph_index=request.paragraph_index,
            space_before=request.space_before,
            space_after=request.space_after,
            space_within=request.space_within,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_textbox_margins",
        description="Update the internal margins of a text box shape.",
    )
    def ppt_set_textbox_margins(
        path: str,
        slide_index: int,
        shape_index: int,
        margin_left: float | None = None,
        margin_right: float | None = None,
        margin_top: float | None = None,
        margin_bottom: float | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Change internal text margins for one shape."""
        request = PowerPointTextboxMarginsRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            margin_left=margin_left,
            margin_right=margin_right,
            margin_top=margin_top,
            margin_bottom=margin_bottom,
            create_backup=create_backup,
        )
        return service.set_textbox_margins(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            margin_left=request.margin_left,
            margin_right=request.margin_right,
            margin_top=request.margin_top,
            margin_bottom=request.margin_bottom,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_text_direction",
        description="Set the text orientation inside a shape.",
    )
    def ppt_set_text_direction(
        path: str,
        slide_index: int,
        shape_index: int,
        direction: str,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Change the text orientation of one shape."""
        request = PowerPointTextDirectionRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            direction=direction,
            create_backup=create_backup,
        )
        return service.set_text_direction(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            direction=request.direction,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_autofit",
        description="Configure automatic text fitting for a text shape.",
    )
    def ppt_set_autofit(
        path: str,
        slide_index: int,
        shape_index: int,
        mode: str = "shape_to_fit_text",
        word_wrap: bool | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Set text autofit behavior on one shape."""
        request = PowerPointAutofitRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            mode=mode,
            word_wrap=word_wrap,
            create_backup=create_backup,
        )
        return service.set_autofit(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            mode=request.mode,
            word_wrap=request.word_wrap,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_proofing_language",
        description="Set the proofing language used for spelling and grammar on a text shape.",
    )
    def ppt_set_proofing_language(
        path: str,
        slide_index: int,
        shape_index: int,
        language: str,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Assign a proofing language to the text of one shape."""
        request = PowerPointProofingLanguageRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            language=language,
            create_backup=create_backup,
        )
        return service.set_proofing_language(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            language=request.language,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_spellcheck_slide",
        description="Check the spelling of text on a single slide, optionally including presenter notes.",
    )
    def ppt_spellcheck_slide(
        path: str,
        slide_index: int,
        include_notes: bool = False,
    ) -> dict[str, object]:
        """Return misspelled terms detected on one slide."""
        request = PowerPointSpellcheckSlideRequest(
            path=path,
            slide_index=slide_index,
            include_notes=include_notes,
            create_backup=False,
        )
        return service.spellcheck_slide(
            path=request.path,
            slide_index=request.slide_index,
            include_notes=request.include_notes,
        ).model_dump()

    @mcp.tool(
        name="ppt_spellcheck_presentation",
        description="Check the spelling of text across the whole presentation, optionally including presenter notes.",
    )
    def ppt_spellcheck_presentation(
        path: str,
        include_notes: bool = False,
    ) -> dict[str, object]:
        """Return misspelled terms detected across the presentation."""
        request = PowerPointSpellcheckPresentationRequest(
            path=path,
            include_notes=include_notes,
            create_backup=False,
        )
        return service.spellcheck_presentation(
            path=request.path,
            include_notes=request.include_notes,
        ).model_dump()

    @mcp.tool(
        name="ppt_translate_text",
        description="Launch PowerPoint's built-in translation UI for the text in a shape.",
    )
    def ppt_translate_text(
        path: str,
        slide_index: int,
        shape_index: int,
        target_language: str,
        source_language: str | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Open the built-in translation experience for one text shape."""
        request = PowerPointTranslateTextRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            target_language=target_language,
            source_language=source_language,
            create_backup=create_backup,
        )
        return service.translate_text(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            target_language=request.target_language,
            source_language=request.source_language,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_get_presenter_notes_all",
        description="Extract presenter notes for every slide in the presentation.",
    )
    def ppt_get_presenter_notes_all(path: str) -> dict[str, object]:
        """Return notes text for all slides."""
        request = DocumentPathRequest(path=path, create_backup=False)
        return service.get_presenter_notes_all(path=request.path).model_dump()

    @mcp.tool(
        name="ppt_find_in_notes",
        description="Search for text inside presenter notes across the whole presentation.",
    )
    def ppt_find_in_notes(path: str, query: str) -> dict[str, object]:
        """Find matching text inside slide notes."""
        request = PowerPointSearchRequest(path=path, query=query)
        return service.find_in_notes(path=request.path, query=request.query).model_dump()

    @mcp.tool(
        name="ppt_replace_notes_text",
        description="Replace text inside presenter notes across the whole presentation.",
    )
    def ppt_replace_notes_text(
        path: str,
        find: str,
        replace: str,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Replace matching text inside notes pages."""
        request = ReplaceTextRequest(
            path=path,
            find_text=find,
            replace_text=replace,
            create_backup=create_backup,
        )
        return service.replace_notes_text(
            path=request.path,
            find_text=request.find_text,
            replace_text=request.replace_text,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_get_slide_shapes",
        description="Inspect all shapes on a slide, including type, position, size, and text previews.",
    )
    def ppt_get_slide_shapes(path: str, slide_index: int) -> dict[str, object]:
        """Return structural information about the shapes on one slide."""
        request = PowerPointSlideRequest(path=path, slide_index=slide_index, create_backup=False)
        return service.get_slide_shapes(path=request.path, slide_index=request.slide_index).model_dump()

    @mcp.tool(
        name="ppt_duplicate_shape",
        description="Duplicate an individual shape on a slide.",
    )
    def ppt_duplicate_shape(
        path: str,
        slide_index: int,
        shape_index: int,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Duplicate one shape and keep the copy on the same slide."""
        request = PowerPointShapeRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            create_backup=create_backup,
        )
        return service.duplicate_shape(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_delete_shape",
        description="Delete a shape by index or exact name from a slide.",
    )
    def ppt_delete_shape(
        path: str,
        slide_index: int,
        shape_index: int | None = None,
        shape_name: str | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Remove one shape from a slide."""
        request = PowerPointShapeSelectorRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            shape_name=shape_name,
            create_backup=create_backup,
        )
        return service.delete_shape(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            shape_name=request.shape_name,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_rename_shape",
        description="Rename a shape to make future automation more robust.",
    )
    def ppt_rename_shape(
        path: str,
        slide_index: int,
        shape_index: int,
        name: str,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Assign a new name to one shape."""
        request = PowerPointShapeNameRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            name=name,
            create_backup=create_backup,
        )
        return service.rename_shape(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            name=request.name,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_find_shapes",
        description="Find shapes on a slide by name, visible text, or shape type.",
    )
    def ppt_find_shapes(
        path: str,
        slide_index: int,
        shape_name_contains: str | None = None,
        text_contains: str | None = None,
        shape_type: str | None = None,
    ) -> dict[str, object]:
        """Search shapes on one slide using partial filters."""
        request = PowerPointShapeSearchRequest(
            path=path,
            slide_index=slide_index,
            shape_name_contains=shape_name_contains,
            text_contains=text_contains,
            shape_type=shape_type,
        )
        return service.find_shapes(
            path=request.path,
            slide_index=request.slide_index,
            shape_name_contains=request.shape_name_contains,
            text_contains=request.text_contains,
            shape_type=request.shape_type,
        ).model_dump()

    @mcp.tool(
        name="ppt_group_shapes",
        description="Group multiple shapes on the same slide into a single grouped shape.",
    )
    def ppt_group_shapes(
        path: str,
        slide_index: int,
        shape_indexes: list[int],
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Group several shapes together."""
        request = PowerPointShapeCollectionRequest(
            path=path,
            slide_index=slide_index,
            shape_indexes=shape_indexes,
            create_backup=create_backup,
        )
        return service.group_shapes(
            path=request.path,
            slide_index=request.slide_index,
            shape_indexes=request.shape_indexes,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_ungroup_shapes",
        description="Ungroup a grouped shape into its component shapes.",
    )
    def ppt_ungroup_shapes(
        path: str,
        slide_index: int,
        shape_index: int,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Ungroup one grouped shape."""
        request = PowerPointShapeRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            create_backup=create_backup,
        )
        return service.ungroup_shapes(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_align_shapes",
        description="Align multiple shapes left, center, right, top, middle, or bottom.",
    )
    def ppt_align_shapes(
        path: str,
        slide_index: int,
        shape_indexes: list[int],
        alignment: str,
        relative_to_slide: bool = False,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Apply a shared alignment to several shapes."""
        request = PowerPointShapesAlignRequest(
            path=path,
            slide_index=slide_index,
            shape_indexes=shape_indexes,
            alignment=alignment,
            relative_to_slide=relative_to_slide,
            create_backup=create_backup,
        )
        return service.align_shapes(
            path=request.path,
            slide_index=request.slide_index,
            shape_indexes=request.shape_indexes,
            alignment=request.alignment,
            relative_to_slide=request.relative_to_slide,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_distribute_shapes",
        description="Distribute multiple shapes horizontally or vertically.",
    )
    def ppt_distribute_shapes(
        path: str,
        slide_index: int,
        shape_indexes: list[int],
        direction: str,
        relative_to_slide: bool = False,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Evenly distribute several shapes."""
        request = PowerPointShapesDistributeRequest(
            path=path,
            slide_index=slide_index,
            shape_indexes=shape_indexes,
            direction=direction,
            relative_to_slide=relative_to_slide,
            create_backup=create_backup,
        )
        return service.distribute_shapes(
            path=request.path,
            slide_index=request.slide_index,
            shape_indexes=request.shape_indexes,
            direction=request.direction,
            relative_to_slide=request.relative_to_slide,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_resize_shapes",
        description="Match the width, height, or both dimensions of multiple shapes to a reference shape.",
    )
    def ppt_resize_shapes(
        path: str,
        slide_index: int,
        shape_indexes: list[int],
        mode: str,
        reference_shape_index: int | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Normalize size across multiple shapes."""
        request = PowerPointShapesResizeRequest(
            path=path,
            slide_index=slide_index,
            shape_indexes=shape_indexes,
            mode=mode,
            reference_shape_index=reference_shape_index,
            create_backup=create_backup,
        )
        return service.resize_shapes(
            path=request.path,
            slide_index=request.slide_index,
            shape_indexes=request.shape_indexes,
            mode=request.mode,
            reference_shape_index=request.reference_shape_index,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_rotate_shape",
        description="Set the rotation angle of a shape.",
    )
    def ppt_rotate_shape(
        path: str,
        slide_index: int,
        shape_index: int,
        rotation: float,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Update one shape rotation in degrees."""
        request = PowerPointShapeRotateRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            rotation=rotation,
            create_backup=create_backup,
        )
        return service.rotate_shape(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            rotation=request.rotation,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_flip_shape",
        description="Flip a shape horizontally or vertically.",
    )
    def ppt_flip_shape(
        path: str,
        slide_index: int,
        shape_index: int,
        direction: str,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Flip one shape along the requested axis."""
        request = PowerPointShapeFlipRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            direction=direction,
            create_backup=create_backup,
        )
        return service.flip_shape(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            direction=request.direction,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_shape_position",
        description="Move a shape to exact x and y coordinates.",
    )
    def ppt_set_shape_position(
        path: str,
        slide_index: int,
        shape_index: int,
        x: float,
        y: float,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Set the position of one shape."""
        request = PowerPointShapePositionRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            x=x,
            y=y,
            create_backup=create_backup,
        )
        return service.set_shape_position(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            x=request.x,
            y=request.y,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_shape_size",
        description="Set the width and or height of a shape.",
    )
    def ppt_set_shape_size(
        path: str,
        slide_index: int,
        shape_index: int,
        width: float | None = None,
        height: float | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Resize one shape to explicit dimensions."""
        request = PowerPointShapeSizeRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            width=width,
            height=height,
            create_backup=create_backup,
        )
        return service.set_shape_size(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            width=request.width,
            height=request.height,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_lock_aspect_ratio",
        description="Lock or unlock the aspect ratio of a shape.",
    )
    def ppt_lock_aspect_ratio(
        path: str,
        slide_index: int,
        shape_index: int,
        lock: bool,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Control aspect ratio locking for one shape."""
        request = PowerPointShapeAspectRatioRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            lock=lock,
            create_backup=create_backup,
        )
        return service.lock_aspect_ratio(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            lock=request.lock,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_shape_visibility",
        description="Show or hide a shape on a slide.",
    )
    def ppt_set_shape_visibility(
        path: str,
        slide_index: int,
        shape_index: int,
        visible: bool,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Update the visible state of one shape."""
        request = PowerPointShapeVisibilityRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            visible=visible,
            create_backup=create_backup,
        )
        return service.set_shape_visibility(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            visible=request.visible,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_bring_to_front",
        description="Bring a shape to the front of the z-order stack.",
    )
    def ppt_bring_to_front(
        path: str,
        slide_index: int,
        shape_index: int,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Move one shape to the very front."""
        request = PowerPointShapeRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            create_backup=create_backup,
        )
        return service.set_shape_z_order(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            command="bring_to_front",
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_send_to_back",
        description="Send a shape to the back of the z-order stack.",
    )
    def ppt_send_to_back(
        path: str,
        slide_index: int,
        shape_index: int,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Move one shape to the very back."""
        request = PowerPointShapeRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            create_backup=create_backup,
        )
        return service.set_shape_z_order(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            command="send_to_back",
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_bring_forward",
        description="Bring a shape one step forward in the z-order stack.",
    )
    def ppt_bring_forward(
        path: str,
        slide_index: int,
        shape_index: int,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Advance one shape by one z-order layer."""
        request = PowerPointShapeRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            create_backup=create_backup,
        )
        return service.set_shape_z_order(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            command="bring_forward",
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_send_backward",
        description="Send a shape one step backward in the z-order stack.",
    )
    def ppt_send_backward(
        path: str,
        slide_index: int,
        shape_index: int,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Move one shape back by one z-order layer."""
        request = PowerPointShapeRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            create_backup=create_backup,
        )
        return service.set_shape_z_order(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            command="send_backward",
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_merge_shapes",
        description="Merge multiple shapes using combine, union, intersect, subtract, or fragment.",
    )
    def ppt_merge_shapes(
        path: str,
        slide_index: int,
        shape_indexes: list[int],
        mode: str,
        primary_shape_index: int | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Merge several shapes into a single result shape."""
        request = PowerPointShapeMergeRequest(
            path=path,
            slide_index=slide_index,
            shape_indexes=shape_indexes,
            mode=mode,
            primary_shape_index=primary_shape_index,
            create_backup=create_backup,
        )
        return service.merge_shapes(
            path=request.path,
            slide_index=request.slide_index,
            shape_indexes=request.shape_indexes,
            mode=request.mode,
            primary_shape_index=request.primary_shape_index,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_crop_shape_to_content",
        description="Shrink a picture shape to its currently visible cropped content.",
    )
    def ppt_crop_shape_to_content(
        path: str,
        slide_index: int,
        shape_index: int,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Commit the current crop box into the visible picture frame size."""
        request = PowerPointShapeCropRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            create_backup=create_backup,
        )
        return service.crop_shape_to_content(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_get_slide_text",
        description="Extract text content from a specific slide in a PowerPoint presentation.",
    )
    def ppt_get_slide_text(path: str, slide_index: int) -> dict[str, object]:
        """Return all visible text fragments found in a slide."""
        request = PowerPointSlideRequest(path=path, slide_index=slide_index, create_backup=False)
        return service.get_slide_text(path=request.path, slide_index=request.slide_index).model_dump()

    @mcp.tool(
        name="ppt_get_slide_notes",
        description="Read the notes or speaker annotations attached to a specific slide.",
    )
    def ppt_get_slide_notes(path: str, slide_index: int) -> dict[str, object]:
        """Return the notes text configured for one slide."""
        request = PowerPointSlideRequest(path=path, slide_index=slide_index, create_backup=False)
        return service.get_slide_notes(path=request.path, slide_index=request.slide_index).model_dump()

    @mcp.tool(
        name="ppt_get_slide_transition",
        description="Inspect the transition effect, speed, and auto-advance settings for a specific slide.",
    )
    def ppt_get_slide_transition(path: str, slide_index: int) -> dict[str, object]:
        """Return the configured transition for one slide."""
        request = PowerPointSlideRequest(path=path, slide_index=slide_index, create_backup=False)
        return service.get_slide_transition(path=request.path, slide_index=request.slide_index).model_dump()

    @mcp.tool(
        name="ppt_get_slide_animations",
        description="Inspect the animation sequence configured on a specific slide.",
    )
    def ppt_get_slide_animations(path: str, slide_index: int) -> dict[str, object]:
        """Return the animations attached to a slide."""
        request = PowerPointSlideRequest(path=path, slide_index=slide_index, create_backup=False)
        return service.get_slide_animations(path=request.path, slide_index=request.slide_index).model_dump()

    @mcp.tool(
        name="ppt_apply_style_preset",
        description="Apply a reusable visual preset to a slide, including background, transition, and primary text styling.",
    )
    def ppt_apply_style_preset(
        path: str,
        slide_index: int,
        preset: str,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Apply a predefined visual style preset to one slide."""
        request = PowerPointPresetRequest(
            path=path,
            slide_index=slide_index,
            preset=preset,
            create_backup=create_backup,
        )
        return service.apply_style_preset(
            path=request.path,
            slide_index=request.slide_index,
            preset=request.preset,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_slide_transition",
        description="Apply a transition effect, speed, and advance timing to a specific slide.",
    )
    def ppt_set_slide_transition(
        path: str,
        slide_index: int,
        effect: str = "fade",
        speed: str = "medium",
        advance_on_click: bool = True,
        advance_after_seconds: float | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Set the transition used when a slide appears during the slideshow."""
        request = PowerPointTransitionRequest(
            path=path,
            slide_index=slide_index,
            effect=effect,
            speed=speed,
            advance_on_click=advance_on_click,
            advance_after_seconds=advance_after_seconds,
            create_backup=create_backup,
        )
        return service.set_slide_transition(
            path=request.path,
            slide_index=request.slide_index,
            effect=request.effect,
            speed=request.speed,
            advance_on_click=request.advance_on_click,
            advance_after_seconds=request.advance_after_seconds,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_shape_text_style",
        description="Update a shape's text content and font style, including font, size, weight, alignment, and color.",
    )
    def ppt_set_shape_text_style(
        path: str,
        slide_index: int,
        shape_index: int,
        text: str | None = None,
        font_name: str | None = None,
        font_size: float | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        color: str | None = None,
        alignment: str | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Apply text and font formatting to one shape."""
        request = PowerPointTextStyleRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            text=text,
            font_name=font_name,
            font_size=font_size,
            bold=bold,
            italic=italic,
            underline=underline,
            color=color,
            alignment=alignment,
            create_backup=create_backup,
        )
        return service.set_shape_text_style(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            text=request.text,
            font_name=request.font_name,
            font_size=request.font_size,
            bold=request.bold,
            italic=request.italic,
            underline=request.underline,
            color=request.color,
            alignment=request.alignment,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_text_gradient",
        description="Apply a two-color gradient fill to the text of a shape.",
    )
    def ppt_set_text_gradient(
        path: str,
        slide_index: int,
        shape_index: int,
        start_color: str,
        end_color: str,
        style: str = "horizontal",
        variant: int = 1,
        text: str | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Apply a gradient to shape text using PowerPoint's text fill engine."""
        request = PowerPointTextGradientRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            start_color=start_color,
            end_color=end_color,
            style=style,
            variant=variant,
            text=text,
            create_backup=create_backup,
        )
        return service.set_text_gradient(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            start_color=request.start_color,
            end_color=request.end_color,
            style=request.style,
            variant=request.variant,
            text=request.text,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_shape_fill",
        description="Set the fill color and transparency of a shape on a slide.",
    )
    def ppt_set_shape_fill(
        path: str,
        slide_index: int,
        shape_index: int,
        color: str,
        transparency: float | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Update the fill styling of one shape."""
        request = PowerPointShapeFillRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            color=color,
            transparency=transparency,
            create_backup=create_backup,
        )
        return service.set_shape_fill(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            color=request.color,
            transparency=request.transparency,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_shape_line",
        description="Set the border color, weight, visibility, and transparency of a shape.",
    )
    def ppt_set_shape_line(
        path: str,
        slide_index: int,
        shape_index: int,
        color: str | None = None,
        weight: float | None = None,
        transparency: float | None = None,
        visible: bool | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Update the line styling of one shape."""
        request = PowerPointShapeLineRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            color=color,
            weight=weight,
            transparency=transparency,
            visible=visible,
            create_backup=create_backup,
        )
        return service.set_shape_line(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            color=request.color,
            weight=request.weight,
            transparency=request.transparency,
            visible=request.visible,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_shape_shadow",
        description="Apply or update shadow settings for a shape.",
    )
    def ppt_set_shape_shadow(
        path: str,
        slide_index: int,
        shape_index: int,
        visible: bool | None = None,
        color: str | None = None,
        transparency: float | None = None,
        blur: float | None = None,
        offset_x: float | None = None,
        offset_y: float | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Update shadow styling for one shape."""
        request = PowerPointShapeShadowRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            visible=visible,
            color=color,
            transparency=transparency,
            blur=blur,
            offset_x=offset_x,
            offset_y=offset_y,
            create_backup=create_backup,
        )
        return service.set_shape_shadow(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            visible=request.visible,
            color=request.color,
            transparency=request.transparency,
            blur=request.blur,
            offset_x=request.offset_x,
            offset_y=request.offset_y,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_shape_glow",
        description="Apply or update glow settings for a shape.",
    )
    def ppt_set_shape_glow(
        path: str,
        slide_index: int,
        shape_index: int,
        color: str | None = None,
        radius: float | None = None,
        transparency: float | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Update glow styling for one shape."""
        request = PowerPointShapeGlowRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            color=color,
            radius=radius,
            transparency=transparency,
            create_backup=create_backup,
        )
        return service.set_shape_glow(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            color=request.color,
            radius=request.radius,
            transparency=request.transparency,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_shape_reflection",
        description="Apply or update reflection settings for a shape.",
    )
    def ppt_set_shape_reflection(
        path: str,
        slide_index: int,
        shape_index: int,
        preset_type: int | None = None,
        blur: float | None = None,
        size: float | None = None,
        offset: float | None = None,
        transparency: float | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Update reflection styling for one shape."""
        request = PowerPointShapeReflectionRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            preset_type=preset_type,
            blur=blur,
            size=size,
            offset=offset,
            transparency=transparency,
            create_backup=create_backup,
        )
        return service.set_shape_reflection(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            preset_type=request.preset_type,
            blur=request.blur,
            size=request.size,
            offset=request.offset,
            transparency=request.transparency,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_shape_soft_edges",
        description="Adjust the soft edge radius of a shape.",
    )
    def ppt_set_shape_soft_edges(
        path: str,
        slide_index: int,
        shape_index: int,
        radius: float,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Set soft edges for one shape."""
        request = PowerPointShapeSoftEdgesRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            radius=radius,
            create_backup=create_backup,
        )
        return service.set_shape_soft_edges(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            radius=request.radius,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_shape_3d",
        description="Apply or update basic 3D settings for a shape.",
    )
    def ppt_set_shape_3d(
        path: str,
        slide_index: int,
        shape_index: int,
        visible: bool | None = None,
        depth: float | None = None,
        rotation_x: float | None = None,
        rotation_y: float | None = None,
        rotation_z: float | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Update 3D styling for one shape."""
        request = PowerPointShape3DRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            visible=visible,
            depth=depth,
            rotation_x=rotation_x,
            rotation_y=rotation_y,
            rotation_z=rotation_z,
            create_backup=create_backup,
        )
        return service.set_shape_3d(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            visible=request.visible,
            depth=request.depth,
            rotation_x=request.rotation_x,
            rotation_y=request.rotation_y,
            rotation_z=request.rotation_z,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_slide_background",
        description="Set a slide background color or restore the master background.",
    )
    def ppt_set_slide_background(
        path: str,
        slide_index: int,
        color: str,
        follow_master: bool = False,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Update the background of one slide."""
        request = PowerPointBackgroundRequest(
            path=path,
            slide_index=slide_index,
            color=color,
            follow_master=follow_master,
            create_backup=create_backup,
        )
        return service.set_slide_background(
            path=request.path,
            slide_index=request.slide_index,
            color=request.color,
            follow_master=request.follow_master,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_slide_background_gradient",
        description="Apply a two-color gradient to the background of a slide.",
    )
    def ppt_set_slide_background_gradient(
        path: str,
        slide_index: int,
        start_color: str,
        end_color: str,
        style: str = "horizontal",
        variant: int = 1,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Set a slide background using a two-color gradient."""
        request = PowerPointBackgroundGradientRequest(
            path=path,
            slide_index=slide_index,
            start_color=start_color,
            end_color=end_color,
            style=style,
            variant=variant,
            create_backup=create_backup,
        )
        return service.set_slide_background_gradient(
            path=request.path,
            slide_index=request.slide_index,
            start_color=request.start_color,
            end_color=request.end_color,
            style=request.style,
            variant=request.variant,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_add_table",
        description="Create a new table shape on a slide, optionally filling initial cell values.",
    )
    def ppt_add_table(
        path: str,
        slide_index: int,
        rows: int,
        columns: int,
        x: float = 72.0,
        y: float = 72.0,
        width: float = 420.0,
        height: float = 220.0,
        values: list[list[object]] | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Insert a new table into one slide."""
        request = PowerPointAddTableRequest(
            path=path,
            slide_index=slide_index,
            rows=rows,
            columns=columns,
            x=x,
            y=y,
            width=width,
            height=height,
            values=values or [],
            create_backup=create_backup,
        )
        return service.add_table(
            path=request.path,
            slide_index=request.slide_index,
            rows=request.rows,
            columns=request.columns,
            x=request.x,
            y=request.y,
            width=request.width,
            height=request.height,
            values=request.values,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_get_slide_tables",
        description="Inspect table shapes on a slide and return their cell contents.",
    )
    def ppt_get_slide_tables(path: str, slide_index: int) -> dict[str, object]:
        """Return all tables detected on a slide."""
        request = PowerPointSlideRequest(path=path, slide_index=slide_index, create_backup=False)
        return service.get_slide_tables(path=request.path, slide_index=request.slide_index).model_dump()

    @mcp.tool(
        name="ppt_set_table_cell_text",
        description="Update the text of a specific PowerPoint table cell.",
    )
    def ppt_set_table_cell_text(
        path: str,
        slide_index: int,
        shape_index: int,
        row_index: int,
        column_index: int,
        text: str,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Write text into one cell of a PowerPoint table."""
        request = PowerPointTableCellRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            row_index=row_index,
            column_index=column_index,
            text=text,
            create_backup=create_backup,
        )
        return service.set_table_cell_text(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            row_index=request.row_index,
            column_index=request.column_index,
            text=request.text,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_add_row_to_table",
        description="Insert a row into an existing PowerPoint table.",
    )
    def ppt_add_row_to_table(
        path: str,
        slide_index: int,
        shape_index: int,
        row_index: int | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Add one row before the given row, or append when omitted."""
        request = PowerPointTableRowRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            row_index=row_index,
            create_backup=create_backup,
        )
        return service.add_row_to_table(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            row_index=request.row_index,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_add_column_to_table",
        description="Insert a column into an existing PowerPoint table.",
    )
    def ppt_add_column_to_table(
        path: str,
        slide_index: int,
        shape_index: int,
        column_index: int | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Add one column before the given column, or append when omitted."""
        request = PowerPointTableColumnRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            column_index=column_index,
            create_backup=create_backup,
        )
        return service.add_column_to_table(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            column_index=request.column_index,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_delete_row_from_table",
        description="Delete a row from an existing PowerPoint table.",
    )
    def ppt_delete_row_from_table(
        path: str,
        slide_index: int,
        shape_index: int,
        row_index: int,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Delete one table row by index."""
        request = PowerPointTableRowRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            row_index=row_index,
            create_backup=create_backup,
        )
        return service.delete_row_from_table(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            row_index=int(request.row_index),
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_delete_column_from_table",
        description="Delete a column from an existing PowerPoint table.",
    )
    def ppt_delete_column_from_table(
        path: str,
        slide_index: int,
        shape_index: int,
        column_index: int,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Delete one table column by index."""
        request = PowerPointTableColumnRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            column_index=column_index,
            create_backup=create_backup,
        )
        return service.delete_column_from_table(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            column_index=int(request.column_index),
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_merge_table_cells",
        description="Merge one PowerPoint table cell into another target cell.",
    )
    def ppt_merge_table_cells(
        path: str,
        slide_index: int,
        shape_index: int,
        row_index: int,
        column_index: int,
        merge_to_row_index: int,
        merge_to_column_index: int,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Merge one table cell with another."""
        request = PowerPointTableMergeCellsRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            row_index=row_index,
            column_index=column_index,
            merge_to_row_index=merge_to_row_index,
            merge_to_column_index=merge_to_column_index,
            create_backup=create_backup,
        )
        return service.merge_table_cells(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            row_index=request.row_index,
            column_index=request.column_index,
            merge_to_row_index=request.merge_to_row_index,
            merge_to_column_index=request.merge_to_column_index,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_split_table_cells",
        description="Split one PowerPoint table cell into multiple rows and columns.",
    )
    def ppt_split_table_cells(
        path: str,
        slide_index: int,
        shape_index: int,
        row_index: int,
        column_index: int,
        num_rows: int,
        num_columns: int,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Split one table cell into a grid."""
        request = PowerPointTableSplitCellsRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            row_index=row_index,
            column_index=column_index,
            num_rows=num_rows,
            num_columns=num_columns,
            create_backup=create_backup,
        )
        return service.split_table_cells(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            row_index=request.row_index,
            column_index=request.column_index,
            num_rows=request.num_rows,
            num_columns=request.num_columns,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_table_style",
        description="Apply a built-in PowerPoint table style and or toggle header, footer, and banding options.",
    )
    def ppt_set_table_style(
        path: str,
        slide_index: int,
        shape_index: int,
        style_id: str | None = None,
        save_formatting: bool = True,
        first_row: bool | None = None,
        first_col: bool | None = None,
        last_row: bool | None = None,
        last_col: bool | None = None,
        horiz_banding: bool | None = None,
        vert_banding: bool | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Apply a table style or update table banding/header options."""
        request = PowerPointTableStyleRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            style_id=style_id,
            save_formatting=save_formatting,
            first_row=first_row,
            first_col=first_col,
            last_row=last_row,
            last_col=last_col,
            horiz_banding=horiz_banding,
            vert_banding=vert_banding,
            create_backup=create_backup,
        )
        return service.set_table_style(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            style_id=request.style_id,
            save_formatting=request.save_formatting,
            first_row=request.first_row,
            first_col=request.first_col,
            last_row=request.last_row,
            last_col=request.last_col,
            horiz_banding=request.horiz_banding,
            vert_banding=request.vert_banding,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_table_cell_style",
        description="Update fill, border, and text styling for a specific PowerPoint table cell.",
    )
    def ppt_set_table_cell_style(
        path: str,
        slide_index: int,
        shape_index: int,
        row_index: int,
        column_index: int,
        text: str | None = None,
        font_name: str | None = None,
        font_size: float | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        color: str | None = None,
        alignment: str | None = None,
        fill_color: str | None = None,
        fill_transparency: float | None = None,
        line_color: str | None = None,
        line_weight: float | None = None,
        line_transparency: float | None = None,
        line_visible: bool | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Style a single table cell."""
        request = PowerPointTableCellStyleRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            row_index=row_index,
            column_index=column_index,
            text=text,
            font_name=font_name,
            font_size=font_size,
            bold=bold,
            italic=italic,
            underline=underline,
            color=color,
            alignment=alignment,
            fill_color=fill_color,
            fill_transparency=fill_transparency,
            line_color=line_color,
            line_weight=line_weight,
            line_transparency=line_transparency,
            line_visible=line_visible,
            create_backup=create_backup,
        )
        return service.set_table_cell_style(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            row_index=request.row_index,
            column_index=request.column_index,
            text=request.text,
            font_name=request.font_name,
            font_size=request.font_size,
            bold=request.bold,
            italic=request.italic,
            underline=request.underline,
            color=request.color,
            alignment=request.alignment,
            fill_color=request.fill_color,
            fill_transparency=request.fill_transparency,
            line_color=request.line_color,
            line_weight=request.line_weight,
            line_transparency=request.line_transparency,
            line_visible=request.line_visible,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_table_row_style",
        description="Update fill, border, and text styling for all cells in a specific table row.",
    )
    def ppt_set_table_row_style(
        path: str,
        slide_index: int,
        shape_index: int,
        row_index: int,
        text: str | None = None,
        font_name: str | None = None,
        font_size: float | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        color: str | None = None,
        alignment: str | None = None,
        fill_color: str | None = None,
        fill_transparency: float | None = None,
        line_color: str | None = None,
        line_weight: float | None = None,
        line_transparency: float | None = None,
        line_visible: bool | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Style a full table row."""
        request = PowerPointTableRowStyleRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            row_index=row_index,
            text=text,
            font_name=font_name,
            font_size=font_size,
            bold=bold,
            italic=italic,
            underline=underline,
            color=color,
            alignment=alignment,
            fill_color=fill_color,
            fill_transparency=fill_transparency,
            line_color=line_color,
            line_weight=line_weight,
            line_transparency=line_transparency,
            line_visible=line_visible,
            create_backup=create_backup,
        )
        return service.set_table_row_style(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            row_index=request.row_index,
            text=request.text,
            font_name=request.font_name,
            font_size=request.font_size,
            bold=request.bold,
            italic=request.italic,
            underline=request.underline,
            color=request.color,
            alignment=request.alignment,
            fill_color=request.fill_color,
            fill_transparency=request.fill_transparency,
            line_color=request.line_color,
            line_weight=request.line_weight,
            line_transparency=request.line_transparency,
            line_visible=request.line_visible,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_table_column_style",
        description="Update fill, border, and text styling for all cells in a specific table column.",
    )
    def ppt_set_table_column_style(
        path: str,
        slide_index: int,
        shape_index: int,
        column_index: int,
        text: str | None = None,
        font_name: str | None = None,
        font_size: float | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        color: str | None = None,
        alignment: str | None = None,
        fill_color: str | None = None,
        fill_transparency: float | None = None,
        line_color: str | None = None,
        line_weight: float | None = None,
        line_transparency: float | None = None,
        line_visible: bool | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Style a full table column."""
        request = PowerPointTableColumnStyleRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            column_index=column_index,
            text=text,
            font_name=font_name,
            font_size=font_size,
            bold=bold,
            italic=italic,
            underline=underline,
            color=color,
            alignment=alignment,
            fill_color=fill_color,
            fill_transparency=fill_transparency,
            line_color=line_color,
            line_weight=line_weight,
            line_transparency=line_transparency,
            line_visible=line_visible,
            create_backup=create_backup,
        )
        return service.set_table_column_style(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            column_index=request.column_index,
            text=request.text,
            font_name=request.font_name,
            font_size=request.font_size,
            bold=request.bold,
            italic=request.italic,
            underline=request.underline,
            color=request.color,
            alignment=request.alignment,
            fill_color=request.fill_color,
            fill_transparency=request.fill_transparency,
            line_color=request.line_color,
            line_weight=request.line_weight,
            line_transparency=request.line_transparency,
            line_visible=request.line_visible,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_distribute_table_rows",
        description="Distribute row heights evenly across a PowerPoint table.",
    )
    def ppt_distribute_table_rows(
        path: str,
        slide_index: int,
        shape_index: int,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Evenly distribute all row heights in a table."""
        request = PowerPointShapeRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            create_backup=create_backup,
        )
        return service.distribute_table_rows(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_distribute_table_columns",
        description="Distribute column widths evenly across a PowerPoint table.",
    )
    def ppt_distribute_table_columns(
        path: str,
        slide_index: int,
        shape_index: int,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Evenly distribute all column widths in a table."""
        request = PowerPointShapeRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            create_backup=create_backup,
        )
        return service.distribute_table_columns(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_autofit_table",
        description="Adjust a PowerPoint table size heuristically to better fit its text content.",
    )
    def ppt_autofit_table(
        path: str,
        slide_index: int,
        shape_index: int,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Apply a content-based autofit heuristic to a table."""
        request = PowerPointShapeRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            create_backup=create_backup,
        )
        return service.autofit_table(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_table_from_csv",
        description="Create a PowerPoint table from a CSV file on a target slide.",
    )
    def ppt_table_from_csv(
        path: str,
        slide_index: int,
        csv_path: str,
        x: float = 72,
        y: float = 72,
        width: float = 360,
        height: float = 220,
        delimiter: str = ",",
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Insert a new table populated from CSV data."""
        request = PowerPointTableFromCsvRequest(
            path=path,
            slide_index=slide_index,
            csv_path=csv_path,
            x=x,
            y=y,
            width=width,
            height=height,
            delimiter=delimiter,
            create_backup=create_backup,
        )
        return service.table_from_csv(
            path=request.path,
            slide_index=request.slide_index,
            csv_path=request.csv_path,
            x=request.x,
            y=request.y,
            width=request.width,
            height=request.height,
            delimiter=request.delimiter,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_sort_table",
        description="Sort PowerPoint table rows by one column, optionally preserving a header row.",
    )
    def ppt_sort_table(
        path: str,
        slide_index: int,
        shape_index: int,
        column_index: int,
        descending: bool = False,
        has_header: bool = False,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Sort an existing table by the selected column."""
        request = PowerPointTableSortRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            column_index=column_index,
            descending=descending,
            has_header=has_header,
            create_backup=create_backup,
        )
        return service.sort_table(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            column_index=request.column_index,
            descending=request.descending,
            has_header=request.has_header,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_table_from_excel_range",
        description="Create or refresh a PowerPoint table from an Excel worksheet range.",
    )
    def ppt_table_from_excel_range(
        path: str,
        slide_index: int,
        excel_path: str,
        sheet: str,
        cell_range: str,
        shape_index: int | None = None,
        x: float = 72,
        y: float = 72,
        width: float = 360,
        height: float = 220,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Insert a new table or refresh an existing one from Excel range data."""
        request = PowerPointTableFromExcelRangeRequest(
            path=path,
            slide_index=slide_index,
            excel_path=excel_path,
            sheet=sheet,
            cell_range=cell_range,
            shape_index=shape_index,
            x=x,
            y=y,
            width=width,
            height=height,
            create_backup=create_backup,
        )
        return service.table_from_excel_range(
            path=request.path,
            slide_index=request.slide_index,
            excel_path=request.excel_path,
            sheet=request.sheet,
            cell_range=request.cell_range,
            shape_index=request.shape_index,
            x=request.x,
            y=request.y,
            width=request.width,
            height=request.height,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_add_shape",
        description="Create an AutoShape on a slide, with optional text and basic styling.",
    )
    def ppt_add_shape(
        path: str,
        slide_index: int,
        shape_type: str = "rectangle",
        x: float = 72.0,
        y: float = 72.0,
        width: float = 180.0,
        height: float = 90.0,
        text: str | None = None,
        fill_color: str | None = None,
        line_color: str | None = None,
        line_weight: float | None = None,
        text_color: str | None = None,
        font_name: str | None = None,
        font_size: float | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Insert a new PowerPoint AutoShape."""
        request = PowerPointAddShapeRequest(
            path=path,
            slide_index=slide_index,
            shape_type=shape_type,
            x=x,
            y=y,
            width=width,
            height=height,
            text=text,
            fill_color=fill_color,
            line_color=line_color,
            line_weight=line_weight,
            text_color=text_color,
            font_name=font_name,
            font_size=font_size,
            create_backup=create_backup,
        )
        return service.add_shape(
            path=request.path,
            slide_index=request.slide_index,
            shape_type=request.shape_type,
            x=request.x,
            y=request.y,
            width=request.width,
            height=request.height,
            text=request.text,
            fill_color=request.fill_color,
            line_color=request.line_color,
            line_weight=request.line_weight,
            text_color=request.text_color,
            font_name=request.font_name,
            font_size=request.font_size,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_add_connector",
        description="Create a connector line between two points on a slide.",
    )
    def ppt_add_connector(
        path: str,
        slide_index: int,
        connector_type: str = "straight",
        begin_x: float = 72.0,
        begin_y: float = 72.0,
        end_x: float = 360.0,
        end_y: float = 72.0,
        line_color: str | None = None,
        line_weight: float | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Insert a new PowerPoint connector shape."""
        request = PowerPointAddConnectorRequest(
            path=path,
            slide_index=slide_index,
            connector_type=connector_type,
            begin_x=begin_x,
            begin_y=begin_y,
            end_x=end_x,
            end_y=end_y,
            line_color=line_color,
            line_weight=line_weight,
            create_backup=create_backup,
        )
        return service.add_connector(
            path=request.path,
            slide_index=request.slide_index,
            connector_type=request.connector_type,
            begin_x=request.begin_x,
            begin_y=request.begin_y,
            end_x=request.end_x,
            end_y=request.end_y,
            line_color=request.line_color,
            line_weight=request.line_weight,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_connect_shapes",
        description="Create a connector anchored to two existing shapes using connection sites.",
    )
    def ppt_connect_shapes(
        path: str,
        slide_index: int,
        start_shape_index: int,
        end_shape_index: int,
        begin_connection_site: int = 1,
        end_connection_site: int = 1,
        connector_type: str = "straight",
        line_color: str | None = None,
        line_weight: float | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Connect two shapes by anchor sites."""
        request = PowerPointConnectShapesRequest(
            path=path,
            slide_index=slide_index,
            start_shape_index=start_shape_index,
            end_shape_index=end_shape_index,
            begin_connection_site=begin_connection_site,
            end_connection_site=end_connection_site,
            connector_type=connector_type,
            line_color=line_color,
            line_weight=line_weight,
            create_backup=create_backup,
        )
        return service.connect_shapes(
            path=request.path,
            slide_index=request.slide_index,
            start_shape_index=request.start_shape_index,
            end_shape_index=request.end_shape_index,
            begin_connection_site=request.begin_connection_site,
            end_connection_site=request.end_connection_site,
            connector_type=request.connector_type,
            line_color=request.line_color,
            line_weight=request.line_weight,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_get_slide_charts",
        description="Inspect chart shapes on a slide and return titles, series, and categories.",
    )
    def ppt_get_slide_charts(path: str, slide_index: int) -> dict[str, object]:
        """Return chart summaries detected on a slide."""
        request = PowerPointSlideRequest(path=path, slide_index=slide_index, create_backup=False)
        return service.get_slide_charts(path=request.path, slide_index=request.slide_index).model_dump()

    @mcp.tool(
        name="ppt_set_chart_title",
        description="Update the title of a chart shape in PowerPoint.",
    )
    def ppt_set_chart_title(
        path: str,
        slide_index: int,
        shape_index: int,
        title: str,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Set the title text of a chart shape."""
        request = PowerPointChartTitleRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            title=title,
            create_backup=create_backup,
        )
        return service.set_chart_title(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            title=request.title,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_chart_data",
        description="Replace or update chart categories, series values, chart type, and title.",
    )
    def ppt_set_chart_data(
        path: str,
        slide_index: int,
        shape_index: int,
        chart_type: str | None = None,
        title: str | None = None,
        categories: list[str | int | float] | None = None,
        series: list[dict[str, object]] | None = None,
        replace_existing: bool = True,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Update the data and metadata of a chart shape."""
        request = PowerPointChartDataRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            chart_type=chart_type,
            title=title,
            categories=categories or [],
            series=series or [],
            replace_existing=replace_existing,
            create_backup=create_backup,
        )
        return service.set_chart_data(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            chart_type=request.chart_type,
            title=request.title,
            categories=request.categories,
            series=request.series,
            replace_existing=request.replace_existing,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_chart_series_style",
        description="Change a chart series fill, line color, and data label visibility.",
    )
    def ppt_set_chart_series_style(
        path: str,
        slide_index: int,
        shape_index: int,
        series_index: int,
        fill_color: str | None = None,
        line_color: str | None = None,
        show_data_labels: bool | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Apply style changes to a specific chart series."""
        request = PowerPointChartSeriesStyleRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            series_index=series_index,
            fill_color=fill_color,
            line_color=line_color,
            show_data_labels=show_data_labels,
            create_backup=create_backup,
        )
        return service.set_chart_series_style(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            series_index=request.series_index,
            fill_color=request.fill_color,
            line_color=request.line_color,
            show_data_labels=request.show_data_labels,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_chart_layout",
        description="Update chart legend visibility, legend position, and axis titles.",
    )
    def ppt_set_chart_layout(
        path: str,
        slide_index: int,
        shape_index: int,
        legend_visible: bool | None = None,
        legend_position: str | None = None,
        category_axis_title: str | None = None,
        value_axis_title: str | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Update chart legend and axis label layout."""
        request = PowerPointChartLayoutRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            legend_visible=legend_visible,
            legend_position=legend_position,
            category_axis_title=category_axis_title,
            value_axis_title=value_axis_title,
            create_backup=create_backup,
        )
        return service.set_chart_layout(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            legend_visible=request.legend_visible,
            legend_position=request.legend_position,
            category_axis_title=request.category_axis_title,
            value_axis_title=request.value_axis_title,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_add_chart",
        description="Create a chart shape on a slide, with optional categories and series data.",
    )
    def ppt_add_chart(
        path: str,
        slide_index: int,
        chart_type: str = "column_clustered",
        x: float = 72.0,
        y: float = 72.0,
        width: float = 420.0,
        height: float = 240.0,
        title: str | None = None,
        categories: list[str | int | float] | None = None,
        series: list[dict[str, object]] | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Insert a new chart into one slide."""
        request = PowerPointAddChartRequest(
            path=path,
            slide_index=slide_index,
            chart_type=chart_type,
            x=x,
            y=y,
            width=width,
            height=height,
            title=title,
            categories=categories or [],
            series=series or [],
            create_backup=create_backup,
        )
        return service.add_chart(
            path=request.path,
            slide_index=request.slide_index,
            chart_type=request.chart_type,
            x=request.x,
            y=request.y,
            width=request.width,
            height=request.height,
            title=request.title,
            categories=request.categories,
            series=request.series,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_refresh_chart",
        description="Refresh an embedded chart so cached data and rendering are recalculated.",
    )
    def ppt_refresh_chart(
        path: str,
        slide_index: int,
        shape_index: int,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Refresh one chart shape."""
        request = PowerPointShapeRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            create_backup=create_backup,
        )
        return service.refresh_chart(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_chart_axis_scale",
        description="Set minimum, maximum, and unit values for a chart axis.",
    )
    def ppt_set_chart_axis_scale(
        path: str,
        slide_index: int,
        shape_index: int,
        axis_kind: str = "value",
        minimum_scale: float | int | None = None,
        maximum_scale: float | int | None = None,
        major_unit: float | int | None = None,
        minor_unit: float | int | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Update the numeric scale settings of a chart axis."""
        request = PowerPointChartAxisScaleRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            axis_kind=axis_kind,
            minimum_scale=minimum_scale,
            maximum_scale=maximum_scale,
            major_unit=major_unit,
            minor_unit=minor_unit,
            create_backup=create_backup,
        )
        return service.set_chart_axis_scale(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            axis_kind=request.axis_kind,
            minimum_scale=request.minimum_scale,
            maximum_scale=request.maximum_scale,
            major_unit=request.major_unit,
            minor_unit=request.minor_unit,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_chart_series_order",
        description="Move a chart series to a different plot order position.",
    )
    def ppt_set_chart_series_order(
        path: str,
        slide_index: int,
        shape_index: int,
        series_index: int,
        position: int,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Reorder a chart series within its plot order."""
        request = PowerPointChartSeriesOrderRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            series_index=series_index,
            position=position,
            create_backup=create_backup,
        )
        return service.set_chart_series_order(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            series_index=request.series_index,
            position=request.position,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_add_chart_series",
        description="Add a new series to an existing chart.",
    )
    def ppt_add_chart_series(
        path: str,
        slide_index: int,
        shape_index: int,
        name: str,
        values: list[float | int],
        categories: list[str | int | float] | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Append a new data series to a chart."""
        request = PowerPointAddChartSeriesRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            name=name,
            values=values,
            categories=categories or [],
            create_backup=create_backup,
        )
        return service.add_chart_series(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            name=request.name,
            values=request.values,
            categories=request.categories,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_delete_chart_series",
        description="Delete one series from an existing chart.",
    )
    def ppt_delete_chart_series(
        path: str,
        slide_index: int,
        shape_index: int,
        series_index: int,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Remove a chart series by index."""
        request = PowerPointDeleteChartSeriesRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            series_index=series_index,
            create_backup=create_backup,
        )
        return service.delete_chart_series(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            series_index=request.series_index,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_chart_data_labels",
        description="Configure chart data labels for one series or all series in a chart.",
    )
    def ppt_set_chart_data_labels(
        path: str,
        slide_index: int,
        shape_index: int,
        series_index: int | None = None,
        visible: bool | None = None,
        show_value: bool | None = None,
        show_category_name: bool | None = None,
        show_series_name: bool | None = None,
        show_percentage: bool | None = None,
        separator: str | None = None,
        position: str | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Update chart data label visibility and content."""
        request = PowerPointChartDataLabelsRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            series_index=series_index,
            visible=visible,
            show_value=show_value,
            show_category_name=show_category_name,
            show_series_name=show_series_name,
            show_percentage=show_percentage,
            separator=separator,
            position=position,
            create_backup=create_backup,
        )
        return service.set_chart_data_labels(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            series_index=request.series_index,
            visible=request.visible,
            show_value=request.show_value,
            show_category_name=request.show_category_name,
            show_series_name=request.show_series_name,
            show_percentage=request.show_percentage,
            separator=request.separator,
            position=request.position,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_chart_gridlines",
        description="Enable or disable major and minor chart gridlines.",
    )
    def ppt_set_chart_gridlines(
        path: str,
        slide_index: int,
        shape_index: int,
        axis_kind: str = "value",
        major: bool | None = None,
        minor: bool | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Toggle chart gridlines on a selected axis."""
        request = PowerPointChartGridlinesRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            axis_kind=axis_kind,
            major=major,
            minor=minor,
            create_backup=create_backup,
        )
        return service.set_chart_gridlines(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            axis_kind=request.axis_kind,
            major=request.major,
            minor=request.minor,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_chart_colors",
        description="Apply fill and line colors to one chart series or to the whole chart palette.",
    )
    def ppt_set_chart_colors(
        path: str,
        slide_index: int,
        shape_index: int,
        series_index: int | None = None,
        fill_color: str | None = None,
        line_color: str | None = None,
        series_fill_colors: list[str] | None = None,
        series_line_colors: list[str] | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Apply a single color or a per-series palette to chart series."""
        request = PowerPointChartColorsRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            series_index=series_index,
            fill_color=fill_color,
            line_color=line_color,
            series_fill_colors=series_fill_colors or [],
            series_line_colors=series_line_colors or [],
            create_backup=create_backup,
        )
        return service.set_chart_colors(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            series_index=request.series_index,
            fill_color=request.fill_color,
            line_color=request.line_color,
            series_fill_colors=request.series_fill_colors,
            series_line_colors=request.series_line_colors,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_change_chart_type",
        description="Convert an existing chart to another chart type.",
    )
    def ppt_change_chart_type(
        path: str,
        slide_index: int,
        shape_index: int,
        chart_type: str,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Change the chart type of one chart shape."""
        request = PowerPointChartTypeChangeRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            chart_type=chart_type,
            create_backup=create_backup,
        )
        return service.change_chart_type(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            chart_type=request.chart_type,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_export_chart_data",
        description="Extract chart categories and series, optionally exporting them to JSON or CSV.",
    )
    def ppt_export_chart_data(
        path: str,
        slide_index: int,
        shape_index: int,
        out_path: str | None = None,
        export_format: str = "json",
    ) -> dict[str, object]:
        """Read chart data from a PowerPoint chart shape."""
        request = PowerPointChartExportDataRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            out_path=out_path,
            export_format=export_format,
        )
        return service.export_chart_data(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            out_path=request.out_path,
            export_format=request.export_format,
        ).model_dump()

    @mcp.tool(
        name="ppt_link_chart_to_excel",
        description="Link a chart to an external Excel workbook range.",
    )
    def ppt_link_chart_to_excel(
        path: str,
        slide_index: int,
        shape_index: int,
        excel_path: str,
        sheet: str,
        cell_range: str,
        plot_by: str | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Bind an existing chart to an external Excel range."""
        request = PowerPointChartLinkRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            excel_path=excel_path,
            sheet=sheet,
            cell_range=cell_range,
            plot_by=plot_by,
            create_backup=create_backup,
        )
        return service.link_chart_to_excel(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            excel_path=request.excel_path,
            sheet=request.sheet,
            cell_range=request.cell_range,
            plot_by=request.plot_by,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_break_chart_link",
        description="Break the external Excel link for a chart and keep embedded data.",
    )
    def ppt_break_chart_link(
        path: str,
        slide_index: int,
        shape_index: int,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Remove the external workbook link from a chart."""
        request = PowerPointShapeRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            create_backup=create_backup,
        )
        return service.break_chart_link(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_get_slide_smartart",
        description="Inspect SmartArt shapes on a slide and return their node texts.",
    )
    def ppt_get_slide_smartart(path: str, slide_index: int) -> dict[str, object]:
        """Return SmartArt summaries detected on a slide."""
        request = PowerPointSlideRequest(path=path, slide_index=slide_index, create_backup=False)
        return service.get_slide_smartart(path=request.path, slide_index=request.slide_index).model_dump()

    @mcp.tool(
        name="ppt_set_smartart_node_text",
        description="Update the text of a specific SmartArt node.",
    )
    def ppt_set_smartart_node_text(
        path: str,
        slide_index: int,
        shape_index: int,
        node_index: int,
        text: str,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Set the text for one SmartArt node."""
        request = PowerPointSmartArtNodeRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            node_index=node_index,
            text=text,
            create_backup=create_backup,
        )
        return service.set_smartart_node_text(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            node_index=request.node_index,
            text=request.text,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_add_smartart",
        description="Create a SmartArt shape on a slide using a supported layout alias or layout id.",
    )
    def ppt_add_smartart(
        path: str,
        slide_index: int,
        layout: str = "basic_list",
        x: float = 72.0,
        y: float = 72.0,
        width: float = 420.0,
        height: float = 240.0,
        node_texts: list[str] | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Insert a new SmartArt object into one slide."""
        request = PowerPointAddSmartArtRequest(
            path=path,
            slide_index=slide_index,
            layout=layout,
            x=x,
            y=y,
            width=width,
            height=height,
            node_texts=node_texts or [],
            create_backup=create_backup,
        )
        return service.add_smartart(
            path=request.path,
            slide_index=request.slide_index,
            layout=request.layout,
            x=request.x,
            y=request.y,
            width=request.width,
            height=request.height,
            node_texts=request.node_texts,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_add_shape_animation",
        description="Add an animation to a shape or one of its elements, including text and table cells, rows, or columns.",
    )
    def ppt_add_shape_animation(
        path: str,
        slide_index: int,
        shape_index: int,
        effect: str = "fade",
        trigger: str = "on_click",
        duration_seconds: float | None = None,
        delay_seconds: float | None = None,
        target_kind: str = "shape",
        animation_level: str | None = None,
        row_index: int | None = None,
        column_index: int | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Attach an animation effect to one shape or a nested table/text target."""
        request = PowerPointAnimationRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            effect=effect,
            trigger=trigger,
            duration_seconds=duration_seconds,
            delay_seconds=delay_seconds,
            target_kind=target_kind,
            animation_level=animation_level,
            row_index=row_index,
            column_index=column_index,
            create_backup=create_backup,
        )
        return service.add_shape_animation(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            effect=request.effect,
            trigger=request.trigger,
            duration_seconds=request.duration_seconds,
            delay_seconds=request.delay_seconds,
            target_kind=request.target_kind,
            animation_level=request.animation_level,
            row_index=request.row_index,
            column_index=request.column_index,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_add_element_animation",
        description="Add an animation to a specific PowerPoint element such as a shape, text box, or table cell.",
    )
    def ppt_add_element_animation(
        path: str,
        slide_index: int,
        shape_index: int,
        effect: str = "fade",
        trigger: str = "on_click",
        duration_seconds: float | None = None,
        delay_seconds: float | None = None,
        target_kind: str = "shape",
        animation_level: str | None = None,
        row_index: int | None = None,
        column_index: int | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Attach an animation effect to an individual slide element."""
        request = PowerPointAnimationRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            effect=effect,
            trigger=trigger,
            duration_seconds=duration_seconds,
            delay_seconds=delay_seconds,
            target_kind=target_kind,
            animation_level=animation_level,
            row_index=row_index,
            column_index=column_index,
            create_backup=create_backup,
        )
        return service.add_shape_animation(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            effect=request.effect,
            trigger=request.trigger,
            duration_seconds=request.duration_seconds,
            delay_seconds=request.delay_seconds,
            target_kind=request.target_kind,
            animation_level=request.animation_level,
            row_index=request.row_index,
            column_index=request.column_index,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_clear_slide_animations",
        description="Remove all animation effects from a specific slide.",
    )
    def ppt_clear_slide_animations(
        path: str,
        slide_index: int,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Delete all animations currently attached to a slide."""
        request = PowerPointSlideRequest(path=path, slide_index=slide_index, create_backup=create_backup)
        return service.clear_slide_animations(
            path=request.path,
            slide_index=request.slide_index,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_replace_text",
        description="Replace text within a specific slide and save the presentation, optionally creating a backup first.",
    )
    def ppt_replace_text(
        path: str,
        slide_index: int,
        find: str,
        replace: str,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Replace text occurrences inside one slide."""
        request = PowerPointTextReplaceRequest(
            path=path,
            slide_index=slide_index,
            find_text=find,
            replace_text=replace,
            create_backup=create_backup,
        )
        return service.replace_text(
            path=request.path,
            slide_index=request.slide_index,
            find_text=request.find_text,
            replace_text=request.replace_text,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_slide_notes",
        description="Replace or append notes text for a specific PowerPoint slide.",
    )
    def ppt_set_slide_notes(
        path: str,
        slide_index: int,
        text: str,
        append: bool = False,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Write notes text into one slide's notes page."""
        request = PowerPointNotesRequest(
            path=path,
            slide_index=slide_index,
            text=text,
            append=append,
            create_backup=create_backup,
        )
        return service.set_slide_notes(
            path=request.path,
            slide_index=request.slide_index,
            text=request.text,
            append=request.append,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_add_slide",
        description="Add a new slide using a layout name or layout id, with optional title and body text.",
    )
    def ppt_add_slide(
        path: str,
        layout: str = "title_and_text",
        position: int | None = None,
        title: str | None = None,
        body_text: str | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Insert a new slide into a presentation."""
        request = PowerPointAddSlideRequest(
            path=path,
            layout=layout,
            position=position,
            title=title,
            body_text=body_text,
            create_backup=create_backup,
        )
        return service.add_slide(
            path=request.path,
            layout=request.layout,
            position=request.position,
            title=request.title,
            body_text=request.body_text,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_insert_image",
        description="Insert an image into a slide at specific coordinates, optionally controlling its size.",
    )
    def ppt_insert_image(
        path: str,
        slide_index: int,
        image_path: str,
        x: float = 72.0,
        y: float = 72.0,
        width: float | None = None,
        height: float | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Place an image on a PowerPoint slide."""
        request = PowerPointImageRequest(
            path=path,
            slide_index=slide_index,
            image_path=image_path,
            x=x,
            y=y,
            width=width,
            height=height,
            create_backup=create_backup,
        )
        return service.insert_image(
            path=request.path,
            slide_index=request.slide_index,
            image_path=request.image_path,
            x=request.x,
            y=request.y,
            width=request.width,
            height=request.height,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_insert_svg",
        description="Insert an SVG into a slide at specific coordinates, optionally controlling its size.",
    )
    def ppt_insert_svg(
        path: str,
        slide_index: int,
        image_path: str,
        x: float = 72,
        y: float = 72,
        width: float | None = None,
        height: float | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Insert one SVG asset on a slide."""
        request = PowerPointImageRequest(
            path=path,
            slide_index=slide_index,
            image_path=image_path,
            x=x,
            y=y,
            width=width,
            height=height,
            create_backup=create_backup,
        )
        return service.insert_svg(
            path=request.path,
            slide_index=request.slide_index,
            image_path=request.image_path,
            x=request.x,
            y=request.y,
            width=request.width,
            height=request.height,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_add_video",
        description="Insert a video file as a media object on a slide.",
    )
    def ppt_add_video(
        path: str,
        slide_index: int,
        media_path: str,
        x: float = 72,
        y: float = 72,
        width: float | None = None,
        height: float | None = None,
        link_to_file: bool = False,
        save_with_document: bool = True,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Insert one video asset on a slide."""
        request = PowerPointMediaRequest(
            path=path,
            slide_index=slide_index,
            media_path=media_path,
            x=x,
            y=y,
            width=width,
            height=height,
            link_to_file=link_to_file,
            save_with_document=save_with_document,
            create_backup=create_backup,
        )
        return service.add_video(
            path=request.path,
            slide_index=request.slide_index,
            media_path=request.media_path,
            x=request.x,
            y=request.y,
            width=request.width,
            height=request.height,
            link_to_file=request.link_to_file,
            save_with_document=request.save_with_document,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_video_playback",
        description="Configure video playback options such as autoplay, loop, pause behavior, and volume.",
    )
    def ppt_set_video_playback(
        path: str,
        slide_index: int,
        shape_index: int,
        autoplay: bool | None = None,
        loop_until_stopped: bool | None = None,
        pause_animation: bool | None = None,
        hide_while_not_playing: bool | None = None,
        stop_after_slides: int | None = None,
        volume: float | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Update playback settings for one video shape."""
        request = PowerPointMediaPlaybackRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            autoplay=autoplay,
            loop_until_stopped=loop_until_stopped,
            pause_animation=pause_animation,
            hide_while_not_playing=hide_while_not_playing,
            stop_after_slides=stop_after_slides,
            volume=volume,
            create_backup=create_backup,
        )
        return service.set_video_playback(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            autoplay=request.autoplay,
            loop_until_stopped=request.loop_until_stopped,
            pause_animation=request.pause_animation,
            hide_while_not_playing=request.hide_while_not_playing,
            stop_after_slides=request.stop_after_slides,
            volume=request.volume,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_trim_video",
        description="Adjust the trim start and or end points of a video media object.",
    )
    def ppt_trim_video(
        path: str,
        slide_index: int,
        shape_index: int,
        start_point: int | None = None,
        end_point: int | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Trim one video shape."""
        request = PowerPointMediaTrimRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            start_point=start_point,
            end_point=end_point,
            create_backup=create_backup,
        )
        return service.trim_video(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            start_point=request.start_point,
            end_point=request.end_point,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_add_audio",
        description="Insert an audio file as a media object on a slide.",
    )
    def ppt_add_audio(
        path: str,
        slide_index: int,
        media_path: str,
        x: float = 72,
        y: float = 72,
        width: float | None = None,
        height: float | None = None,
        link_to_file: bool = False,
        save_with_document: bool = True,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Insert one audio asset on a slide."""
        request = PowerPointMediaRequest(
            path=path,
            slide_index=slide_index,
            media_path=media_path,
            x=x,
            y=y,
            width=width,
            height=height,
            link_to_file=link_to_file,
            save_with_document=save_with_document,
            create_backup=create_backup,
        )
        return service.add_audio(
            path=request.path,
            slide_index=request.slide_index,
            media_path=request.media_path,
            x=request.x,
            y=request.y,
            width=request.width,
            height=request.height,
            link_to_file=request.link_to_file,
            save_with_document=request.save_with_document,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_audio_playback",
        description="Configure audio playback options such as autoplay, loop, cross-slide playback, and volume.",
    )
    def ppt_set_audio_playback(
        path: str,
        slide_index: int,
        shape_index: int,
        autoplay: bool | None = None,
        loop_until_stopped: bool | None = None,
        pause_animation: bool | None = None,
        hide_while_not_playing: bool | None = None,
        stop_after_slides: int | None = None,
        volume: float | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Update playback settings for one audio shape."""
        request = PowerPointMediaPlaybackRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            autoplay=autoplay,
            loop_until_stopped=loop_until_stopped,
            pause_animation=pause_animation,
            hide_while_not_playing=hide_while_not_playing,
            stop_after_slides=stop_after_slides,
            volume=volume,
            create_backup=create_backup,
        )
        return service.set_audio_playback(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            autoplay=request.autoplay,
            loop_until_stopped=request.loop_until_stopped,
            pause_animation=request.pause_animation,
            hide_while_not_playing=request.hide_while_not_playing,
            stop_after_slides=request.stop_after_slides,
            volume=request.volume,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_extract_media_inventory",
        description="Inspect audio and video media used across the presentation.",
    )
    def ppt_extract_media_inventory(path: str) -> dict[str, object]:
        """Inventory media shapes present in the presentation."""
        request = DocumentPathRequest(path=path, create_backup=False)
        return service.extract_media_inventory(path=request.path).model_dump()

    @mcp.tool(
        name="ppt_replace_image",
        description="Replace an existing image shape while preserving its position and size.",
    )
    def ppt_replace_image(
        path: str,
        slide_index: int,
        shape_index: int,
        image_path: str,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Swap the image content of a shape for a new file."""
        request = PowerPointReplaceImageRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            image_path=image_path,
            create_backup=create_backup,
        )
        return service.replace_image(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            image_path=request.image_path,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_crop_image",
        description="Crop a picture shape by setting one or more crop margins in points.",
    )
    def ppt_crop_image(
        path: str,
        slide_index: int,
        shape_index: int,
        crop_left: float | None = None,
        crop_right: float | None = None,
        crop_top: float | None = None,
        crop_bottom: float | None = None,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Adjust crop margins on one image shape."""
        request = PowerPointCropImageRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            crop_left=crop_left,
            crop_right=crop_right,
            crop_top=crop_top,
            crop_bottom=crop_bottom,
            create_backup=create_backup,
        )
        return service.crop_image(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            crop_left=request.crop_left,
            crop_right=request.crop_right,
            crop_top=request.crop_top,
            crop_bottom=request.crop_bottom,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_reset_image",
        description="Reset image crop and basic picture adjustments back to default values.",
    )
    def ppt_reset_image(
        path: str,
        slide_index: int,
        shape_index: int,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Reset a picture shape crop and basic formatting."""
        request = PowerPointShapeRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            create_backup=create_backup,
        )
        return service.reset_image(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_image_transparency",
        description="Adjust picture transparency using the best COM-compatible fallback available.",
    )
    def ppt_set_image_transparency(
        path: str,
        slide_index: int,
        shape_index: int,
        value: float,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Update image transparency for one picture shape."""
        request = PowerPointImageFormatRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            value=value,
            create_backup=create_backup,
        )
        return service.set_image_transparency(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            value=request.value,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_image_brightness",
        description="Set the brightness of a picture shape between 0.0 and 1.0.",
    )
    def ppt_set_image_brightness(
        path: str,
        slide_index: int,
        shape_index: int,
        value: float,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Update picture brightness for one image shape."""
        request = PowerPointImageFormatRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            value=value,
            create_backup=create_backup,
        )
        return service.set_image_brightness(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            value=request.value,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_set_image_contrast",
        description="Set the contrast of a picture shape between 0.0 and 1.0.",
    )
    def ppt_set_image_contrast(
        path: str,
        slide_index: int,
        shape_index: int,
        value: float,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Update picture contrast for one image shape."""
        request = PowerPointImageFormatRequest(
            path=path,
            slide_index=slide_index,
            shape_index=shape_index,
            value=value,
            create_backup=create_backup,
        )
        return service.set_image_contrast(
            path=request.path,
            slide_index=request.slide_index,
            shape_index=request.shape_index,
            value=request.value,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_apply_theme",
        description="Apply a .thmx theme file to the presentation and save the result.",
    )
    def ppt_apply_theme(path: str, theme_path: str, create_backup: bool = True) -> dict[str, object]:
        """Apply a PowerPoint theme file to a presentation."""
        request = PowerPointThemeRequest(path=path, theme_path=theme_path, create_backup=create_backup)
        return service.apply_theme(
            path=request.path,
            theme_path=request.theme_path,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_apply_builtin_theme",
        description="Apply an installed PowerPoint built-in theme by name.",
    )
    def ppt_apply_builtin_theme(path: str, theme_name: str, create_backup: bool = True) -> dict[str, object]:
        """Apply a PowerPoint built-in theme resolved from the local Office installation."""
        request = PowerPointBuiltinThemeRequest(path=path, theme_name=theme_name, create_backup=create_backup)
        return service.apply_builtin_theme(
            path=request.path,
            theme_name=request.theme_name,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_apply_design_ideas",
        description="Trigger PowerPoint Design Ideas and optionally apply a deterministic style-preset fallback.",
    )
    def ppt_apply_design_ideas(
        path: str,
        slide_index: int | None = None,
        fallback_preset: str | None = "executive",
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Run the Design Ideas workflow with an automated fallback for non-interactive use."""
        request = PowerPointDesignIdeasRequest(
            path=path,
            slide_index=slide_index,
            fallback_preset=fallback_preset,
            create_backup=create_backup,
        )
        return service.apply_design_ideas(
            path=request.path,
            slide_index=request.slide_index,
            fallback_preset=request.fallback_preset,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_apply_theme_variant",
        description="Apply a theme variant for one slide master, with a style-preset fallback when COM does not expose variants.",
    )
    def ppt_apply_theme_variant(
        path: str,
        master_index: int,
        variant: str,
        create_backup: bool = True,
    ) -> dict[str, object]:
        """Apply a theme variant or a compatible visual fallback."""
        request = PowerPointThemeVariantRequest(
            path=path,
            master_index=master_index,
            variant=variant,
            create_backup=create_backup,
        )
        return service.apply_theme_variant(
            path=request.path,
            master_index=request.master_index,
            variant=request.variant,
            create_backup=request.create_backup,
        ).model_dump()

    @mcp.tool(
        name="ppt_extract_theme",
        description="Summarize the active theme, masters, layouts, placeholders, fonts, and colors.",
    )
    def ppt_extract_theme(path: str, master_index: int | None = None) -> dict[str, object]:
        """Export a theme summary for the whole presentation or one slide master."""
        request = PowerPointOptionalMasterRequest(path=path, master_index=master_index, create_backup=False)
        return service.extract_theme(path=request.path, master_index=request.master_index).model_dump()

    @mcp.tool(
        name="ppt_export_pdf",
        description="Export a PowerPoint presentation to PDF without editing the original file.",
    )
    def ppt_export_pdf(path: str, out_path: str) -> dict[str, object]:
        """Export a PowerPoint presentation to a PDF file."""
        request = ExportPdfRequest(path=path, out_path=out_path)
        return service.export_pdf(path=request.path, out_path=request.out_path).model_dump()

    @mcp.tool(
        name="ppt_export_slide_images",
        description="Export every slide in a presentation as image files into an output directory.",
    )
    def ppt_export_slide_images(
        path: str,
        out_dir: str,
        image_format: str = "png",
        width: int | None = None,
        height: int | None = None,
    ) -> dict[str, object]:
        """Render all slides as image files."""
        request = PowerPointExportSlideImagesRequest(
            path=path,
            out_dir=out_dir,
            image_format=image_format,
            width=width,
            height=height,
        )
        return service.export_slide_images(
            path=request.path,
            out_dir=request.out_dir,
            image_format=request.image_format,
            width=request.width,
            height=request.height,
        ).model_dump()

    @mcp.tool(
        name="ppt_save_as",
        description="Save a presentation as a new PowerPoint file without overwriting the original path.",
    )
    def ppt_save_as(path: str, out_path: str) -> dict[str, object]:
        """Save a copy of a presentation to a new path."""
        request = SaveAsRequest(path=path, out_path=out_path)
        return service.save_as(path=request.path, out_path=request.out_path).model_dump()
