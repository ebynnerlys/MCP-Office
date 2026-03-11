from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from office_ai_mcp.models.requests import (
    ExportPdfRequest,
    PowerPointAddSlideRequest,
    PowerPointAddChartRequest,
    PowerPointAddConnectorRequest,
    PowerPointAddShapeRequest,
    PowerPointAddSmartArtRequest,
    PowerPointAddTableRequest,
    PowerPointAnimationRequest,
    PowerPointBackgroundRequest,
    PowerPointChartLayoutRequest,
    PowerPointChartDataRequest,
    PowerPointChartSeriesStyleRequest,
    PowerPointChartTitleRequest,
    PowerPointConnectShapesRequest,
    PowerPointExportSlideImagesRequest,
    PowerPointImageRequest,
    PowerPointNotesRequest,
    PowerPointPresetRequest,
    PowerPointShapeFillRequest,
    PowerPointShapeLineRequest,
    PowerPointSlideRequest,
    PowerPointSmartArtNodeRequest,
    PowerPointTableCellRequest,
    PowerPointTextReplaceRequest,
    PowerPointTextStyleRequest,
    PowerPointThemeRequest,
    PowerPointTransitionRequest,
    SaveAsRequest,
)
from office_ai_mcp.services.powerpoint_service import PowerPointService


def register_powerpoint_tools(mcp: FastMCP, service: PowerPointService) -> None:
    @mcp.tool(
        name="ppt_list_slides",
        description="List slides in a PowerPoint presentation with titles and shape counts.",
    )
    def ppt_list_slides(path: str) -> dict[str, object]:
        """Summarize the slides present in a PowerPoint file."""
        return service.list_slides(path).model_dump()

    @mcp.tool(
        name="ppt_get_slide_shapes",
        description="Inspect all shapes on a slide, including type, position, size, and text previews.",
    )
    def ppt_get_slide_shapes(path: str, slide_index: int) -> dict[str, object]:
        """Return structural information about the shapes on one slide."""
        request = PowerPointSlideRequest(path=path, slide_index=slide_index, create_backup=False)
        return service.get_slide_shapes(path=request.path, slide_index=request.slide_index).model_dump()

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
