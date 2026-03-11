from __future__ import annotations

from contextlib import contextmanager, suppress
from pathlib import Path
from typing import Iterator

from tenacity import retry, stop_after_attempt, wait_fixed

from office_ai_mcp.config import Settings
from office_ai_mcp.models.responses import (
    AnimationSummary,
    ChartSeriesSummary,
    ChartSummary,
    OperationResult,
    PresentationSummary,
    ShapeSummary,
    SlideAnimationsResult,
    SlideChartsResult,
    SlideNotesResult,
    SlideShapesResult,
    SlideSmartArtResult,
    SlideSummary,
    SlideTablesResult,
    SlideTextResult,
    SlideTransitionResult,
    SmartArtNodeSummary,
    SmartArtSummary,
    TableSummary,
)
from office_ai_mcp.services.base import OfficeService
from office_ai_mcp.utils.com_cleanup import office_application
from office_ai_mcp.utils.paths import ensure_directory, validate_file_path

SLIDE_LAYOUTS: dict[str, int] = {
    "title": 1,
    "title_and_text": 2,
    "two_column_text": 3,
    "table": 4,
    "chart": 8,
    "title_only": 11,
    "blank": 12,
    "section_header": 33,
    "two_content": 34,
    "content_with_caption": 35,
    "picture_with_caption": 36,
}

IMAGE_SUFFIXES = (".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff")
THEME_SUFFIXES = (".thmx",)
EXPORT_IMAGE_FORMATS = {
    "png": ("PNG", ".png"),
    "jpg": ("JPG", ".jpg"),
    "jpeg": ("JPG", ".jpg"),
    "gif": ("GIF", ".gif"),
    "bmp": ("BMP", ".bmp"),
    "tif": ("TIF", ".tif"),
    "tiff": ("TIF", ".tif"),
}
CHART_TYPES = {
    "area": 1,
    "line": 4,
    "pie": 5,
    "doughnut": -4120,
    "scatter": -4169,
    "radar": -4151,
    "column_clustered": 51,
    "column_stacked": 52,
    "bar_clustered": 57,
}
SHAPE_TYPES = {
    "rectangle": 1,
    "diamond": 4,
    "rounded_rectangle": 5,
    "oval": 9,
    "hexagon": 10,
    "parallelogram": 2,
    "chevron": 52,
    "right_arrow": 33,
}
CONNECTOR_TYPES = {
    "straight": 1,
    "elbow": 2,
    "curved": 3,
}
CHART_LEGEND_POSITIONS = {
    "bottom": -4107,
    "corner": 2,
    "left": -4131,
    "right": -4152,
    "top": -4160,
}

TRANSITION_EFFECTS = {
    "none": 0,
    "cut": 257,
    "fade": 3849,
    "push_left": 3853,
    "push_right": 3854,
    "push_up": 3855,
    "push_down": 3852,
    "wipe_left": 2817,
    "wipe_right": 2819,
    "wipe_up": 2818,
    "wipe_down": 2820,
    "random": 513,
}
TRANSITION_SPEEDS = {
    "slow": 1,
    "medium": 2,
    "fast": 3,
}
TEXT_ALIGNMENTS = {
    "left": 1,
    "center": 2,
    "right": 3,
    "justify": 4,
    "distribute": 5,
    "thai_distribute": 6,
    "justify_low": 7,
}
ANIMATION_EFFECTS = {
    "appear": 1,
    "fly": 2,
    "fade": 10,
    "split": 16,
    "swivel": 19,
    "wipe": 22,
    "zoom": 23,
    "bounce": 26,
    "float": 30,
    "rise_up": 34,
    "descend": 42,
}
ANIMATION_TRIGGERS = {
    "on_click": 1,
    "with_previous": 2,
    "after_previous": 3,
}
ANIMATION_LEVELS = {
    "none": 0,
    "all_text_levels": 1,
    "first_text_level": 2,
    "second_text_level": 3,
    "third_text_level": 4,
    "fourth_text_level": 5,
    "fifth_text_level": 6,
}
ANIMATION_TARGET_ALIASES = {
    "shape": "shape",
    "element": "shape",
    "text": "text",
    "table": "table",
    "cell": "table_cell",
    "table_cell": "table_cell",
    "row": "table_row",
    "table_row": "table_row",
    "column": "table_column",
    "table_column": "table_column",
    "cells": "table_cells",
    "table_cells": "table_cells",
}
STYLE_PRESETS = {
    "executive": {
        "background": "#F8FAFC",
        "transition_effect": "fade",
        "transition_speed": "medium",
        "title_fill": "#E2E8F0",
        "title_line": "#0F172A",
        "title_text_color": "#0F172A",
        "title_font_name": "Aptos Display",
        "title_font_size": 28,
        "title_bold": True,
        "body_text_color": "#1D4ED8",
        "body_font_name": "Aptos",
        "body_font_size": 20,
        "body_alignment": "left",
    },
    "spotlight": {
        "background": "#101418",
        "transition_effect": "fade",
        "transition_speed": "fast",
        "title_fill": "#E4572E",
        "title_line": "#E4572E",
        "title_text_color": "#FFFFFF",
        "title_font_name": "Aptos Display",
        "title_font_size": 30,
        "title_bold": True,
        "body_text_color": "#F8FAFC",
        "body_font_name": "Aptos",
        "body_font_size": 21,
        "body_alignment": "center",
    },
    "warm_briefing": {
        "background": "#FFF7ED",
        "transition_effect": "wipe_left",
        "transition_speed": "medium",
        "title_fill": "#FED7AA",
        "title_line": "#C2410C",
        "title_text_color": "#7C2D12",
        "title_font_name": "Aptos Display",
        "title_font_size": 28,
        "title_bold": True,
        "body_text_color": "#9A3412",
        "body_font_name": "Aptos",
        "body_font_size": 20,
        "body_alignment": "left",
    },
}
SMARTART_LAYOUTS = {
    "basic_list": "urn:microsoft.com/office/officeart/2005/8/layout/default",
    "picture_title_list": "urn:microsoft.com/office/officeart/2005/8/layout/pList1",
    "vertical_bullet_list": "urn:microsoft.com/office/officeart/2005/8/layout/vList2",
    "vertical_box_list": "urn:microsoft.com/office/officeart/2005/8/layout/list1",
    "horizontal_bullet_list": "urn:microsoft.com/office/officeart/2005/8/layout/hList1",
    "picture_accent_list": "urn:microsoft.com/office/officeart/2008/layout/PictureAccentList",
    "segmented_process": "urn:microsoft.com/office/officeart/2005/8/layout/process4",
    "basic_process": "urn:microsoft.com/office/officeart/2005/8/layout/process1",
    "hierarchy_list": "urn:microsoft.com/office/officeart/2005/8/layout/hierarchy3",
    "table_hierarchy": "urn:microsoft.com/office/officeart/2005/8/layout/hierarchy4",
    "step_up_process": "urn:microsoft.com/office/officeart/2009/3/layout/StepUpProcess",
    "step_down_process": "urn:microsoft.com/office/officeart/2005/8/layout/StepDownProcess",
}


def resolve_slide_layout(layout: str) -> int:
    normalized = layout.strip().lower().replace("-", "_").replace(" ", "_")
    if normalized.isdigit():
        return int(normalized)
    if normalized not in SLIDE_LAYOUTS:
        supported = ", ".join(sorted(SLIDE_LAYOUTS))
        raise ValueError(f"Unsupported slide layout: {layout}. Supported values: {supported}")
    return SLIDE_LAYOUTS[normalized]


def resolve_export_image_format(image_format: str) -> tuple[str, str]:
    normalized = image_format.strip().lower()
    if normalized not in EXPORT_IMAGE_FORMATS:
        supported = ", ".join(sorted(EXPORT_IMAGE_FORMATS))
        raise ValueError(f"Unsupported image format: {image_format}. Supported values: {supported}")
    return EXPORT_IMAGE_FORMATS[normalized]


def resolve_chart_type(chart_type: str) -> int:
    normalized = chart_type.strip().lower().replace("-", "_").replace(" ", "_")
    if normalized.lstrip("-").isdigit():
        return int(normalized)
    if normalized not in CHART_TYPES:
        supported = ", ".join(sorted(CHART_TYPES))
        raise ValueError(f"Unsupported chart type: {chart_type}. Supported values: {supported}")
    return CHART_TYPES[normalized]


def resolve_shape_type(shape_type: str) -> int:
    normalized = shape_type.strip().lower().replace("-", "_").replace(" ", "_")
    if normalized.lstrip("-").isdigit():
        return int(normalized)
    if normalized not in SHAPE_TYPES:
        supported = ", ".join(sorted(SHAPE_TYPES))
        raise ValueError(f"Unsupported shape type: {shape_type}. Supported values: {supported}")
    return SHAPE_TYPES[normalized]


def resolve_connector_type(connector_type: str) -> int:
    normalized = connector_type.strip().lower().replace("-", "_").replace(" ", "_")
    if normalized.lstrip("-").isdigit():
        return int(normalized)
    if normalized not in CONNECTOR_TYPES:
        supported = ", ".join(sorted(CONNECTOR_TYPES))
        raise ValueError(f"Unsupported connector type: {connector_type}. Supported values: {supported}")
    return CONNECTOR_TYPES[normalized]


def resolve_chart_legend_position(position: str) -> int:
    normalized = position.strip().lower().replace("-", "_").replace(" ", "_")
    if normalized.lstrip("-").isdigit():
        return int(normalized)
    if normalized not in CHART_LEGEND_POSITIONS:
        supported = ", ".join(sorted(CHART_LEGEND_POSITIONS))
        raise ValueError(f"Unsupported chart legend position: {position}. Supported values: {supported}")
    return CHART_LEGEND_POSITIONS[normalized]


def resolve_named_constant_alias(value: str, aliases: dict[str, str | int]) -> str | int:
    normalized = value.strip().lower().replace("-", "_").replace(" ", "_")
    if normalized.lstrip("-").isdigit():
        return int(normalized)
    return aliases.get(normalized, value.strip())


def parse_office_color(color: str) -> int:
    cleaned = color.strip()
    if cleaned.startswith("#"):
        hex_value = cleaned[1:]
        if len(hex_value) != 6:
            raise ValueError("Hex colors must use the #RRGGBB format")
        red = int(hex_value[0:2], 16)
        green = int(hex_value[2:4], 16)
        blue = int(hex_value[4:6], 16)
        return red + (green << 8) + (blue << 16)

    if "," in cleaned:
        parts = [part.strip() for part in cleaned.split(",")]
        if len(parts) != 3:
            raise ValueError("RGB colors must use the r,g,b format")
        red, green, blue = (int(part) for part in parts)
        for value in (red, green, blue):
            if not 0 <= value <= 255:
                raise ValueError("RGB components must be between 0 and 255")
        return red + (green << 8) + (blue << 16)

    if cleaned.isdigit():
        return int(cleaned)

    raise ValueError(f"Unsupported color value: {color}")


def office_color_to_hex(rgb_value: int | None) -> str | None:
    if rgb_value is None:
        return None
    red = rgb_value & 0xFF
    green = (rgb_value >> 8) & 0xFF
    blue = (rgb_value >> 16) & 0xFF
    return f"#{red:02X}{green:02X}{blue:02X}"


def alias_for_value(value: int | None, aliases: dict[str, int]) -> str | None:
    if value is None:
        return None
    for alias, alias_value in aliases.items():
        if alias_value == value:
            return alias
    return None


def normalize_office_value(value: object) -> object:
    if isinstance(value, tuple):
        return [normalize_office_value(item) for item in value]
    if isinstance(value, list):
        return [normalize_office_value(item) for item in value]
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    try:
        return [normalize_office_value(item) for item in value]  # type: ignore[arg-type]
    except Exception:
        return str(value)


def resolve_style_preset(preset: str) -> dict[str, object]:
    normalized = preset.strip().lower().replace("-", "_").replace(" ", "_")
    if normalized not in STYLE_PRESETS:
        supported = ", ".join(sorted(STYLE_PRESETS))
        raise ValueError(f"Unsupported style preset: {preset}. Supported presets: {supported}")
    return STYLE_PRESETS[normalized]


def resolve_animation_level(level: str) -> int:
    normalized = level.strip().lower().replace("-", "_").replace(" ", "_")
    if normalized.lstrip("-").isdigit():
        return int(normalized)
    if normalized not in ANIMATION_LEVELS:
        supported = ", ".join(sorted(ANIMATION_LEVELS))
        raise ValueError(f"Unsupported animation level: {level}. Supported values: {supported}")
    return ANIMATION_LEVELS[normalized]


def resolve_animation_target(target_kind: str) -> str:
    normalized = target_kind.strip().lower().replace("-", "_").replace(" ", "_")
    if normalized not in ANIMATION_TARGET_ALIASES:
        supported = ", ".join(sorted(ANIMATION_TARGET_ALIASES))
        raise ValueError(f"Unsupported animation target: {target_kind}. Supported values: {supported}")
    return ANIMATION_TARGET_ALIASES[normalized]


def resolve_smartart_layout_identifier(layout: str) -> int | str:
    normalized = layout.strip().lower().replace("-", "_").replace(" ", "_")
    if normalized.isdigit():
        return int(normalized)
    if layout.strip().startswith("urn:"):
        return layout.strip()
    if normalized not in SMARTART_LAYOUTS:
        supported = ", ".join(sorted(SMARTART_LAYOUTS))
        raise ValueError(f"Unsupported SmartArt layout: {layout}. Supported values: {supported}")
    return SMARTART_LAYOUTS[normalized]


class PowerPointService(OfficeService):
    allowed_suffixes = (".ppt", ".pptx", ".pptm")

    def __init__(self, settings: Settings) -> None:
        super().__init__(settings)

    def _resolve_office_constant(self, constants: object, value: str, aliases: dict[str, str | int]) -> int:
        resolved = resolve_named_constant_alias(value, aliases)
        if isinstance(resolved, int):
            return resolved

        try:
            return int(getattr(constants, resolved))
        except AttributeError as exc:
            supported = ", ".join(sorted(aliases))
            raise ValueError(f"Unsupported constant value: {value}. Supported aliases: {supported}") from exc

    @contextmanager
    def _open_presentation_session(
        self,
        source: Path,
        *,
        read_only: bool,
        visible: bool | None = None,
        with_window: bool | None = None,
    ) -> Iterator[tuple[object, object]]:
        effective_visible = self.settings.office_visible if visible is None else visible
        effective_with_window = effective_visible if with_window is None else with_window

        with office_application("PowerPoint.Application", visible=effective_visible) as powerpoint:
            presentation = None
            try:
                presentation = powerpoint.Presentations.Open(
                    str(source),
                    ReadOnly=read_only,
                    Untitled=False,
                    WithWindow=effective_with_window,
                )
                yield powerpoint, presentation
            finally:
                if presentation is not None:
                    with suppress(Exception):
                        presentation.Close()

    @contextmanager
    def _open_presentation(self, source: Path, *, read_only: bool) -> Iterator[object]:
        with self._open_presentation_session(source, read_only=read_only, visible=self.settings.office_visible, with_window=False) as (_, presentation):
            yield presentation

    def _get_slide_title(self, slide: object) -> str | None:
        with suppress(Exception):
            if slide.Shapes.HasTitle:
                title = str(slide.Shapes.Title.TextFrame.TextRange.Text).strip()
                return title or None
        return None

    def _iter_shapes(self, slide: object) -> Iterator[tuple[int, object]]:
        for index in range(1, int(slide.Shapes.Count) + 1):
            yield index, slide.Shapes(index)

    def _shape_has_text(self, shape: object) -> bool:
        with suppress(Exception):
            return bool(shape.HasTextFrame and shape.TextFrame.HasText)
        return False

    def _shape_has_table(self, shape: object) -> bool:
        with suppress(Exception):
            return bool(shape.HasTable)
        return False

    def _shape_has_chart(self, shape: object) -> bool:
        with suppress(Exception):
            return bool(shape.HasChart)
        return False

    def _shape_has_smartart(self, shape: object) -> bool:
        with suppress(Exception):
            return bool(shape.HasSmartArt)
        return False

    def _find_title_shape(self, slide: object) -> object | None:
        with suppress(Exception):
            if slide.Shapes.HasTitle:
                return slide.Shapes.Title
        return None

    def _find_primary_body_shape(self, slide: object) -> object | None:
        title_shape = self._find_title_shape(slide)
        title_name = None
        if title_shape is not None:
            with suppress(Exception):
                title_name = str(title_shape.Name)

        for _, shape in self._iter_shapes(slide):
            with suppress(Exception):
                if title_name and str(shape.Name) == title_name:
                    continue
                if self._shape_has_text(shape):
                    return shape
        return None

    def _get_slide_texts(self, slide: object) -> list[str]:
        texts: list[str] = []
        for _, shape in self._iter_shapes(slide):
            with suppress(Exception):
                if shape.HasTextFrame and shape.TextFrame.HasText:
                    text = str(shape.TextFrame.TextRange.Text).strip()
                    if text:
                        texts.append(text)
        return texts

    def _shape_summary(self, shape: object, index: int) -> ShapeSummary:
        placeholder_type = None
        with suppress(Exception):
            placeholder_type = int(shape.PlaceholderFormat.Type)

        text_preview = None
        has_text = False
        with suppress(Exception):
            has_text = bool(shape.HasTextFrame and shape.TextFrame.HasText)
            if has_text:
                preview = str(shape.TextFrame.TextRange.Text).strip()
                text_preview = preview[:120] if preview else None

        fill_color = None
        fill_transparency = None
        with suppress(Exception):
            fill_color = office_color_to_hex(int(shape.Fill.ForeColor.RGB))
            fill_transparency = float(shape.Fill.Transparency)

        line_color = None
        line_transparency = None
        line_weight = None
        with suppress(Exception):
            line_color = office_color_to_hex(int(shape.Line.ForeColor.RGB))
            line_transparency = float(shape.Line.Transparency)
            line_weight = float(shape.Line.Weight)

        text_color = None
        font_name = None
        font_size = None
        font_bold = None
        font_italic = None
        with suppress(Exception):
            if has_text:
                font = shape.TextFrame.TextRange.Font
                text_color = office_color_to_hex(int(font.Color.RGB))
                font_name = str(font.Name) or None
                font_size = float(font.Size) if font.Size else None
                font_bold = bool(font.Bold)
                font_italic = bool(font.Italic)

        return ShapeSummary(
            shape_index=index,
            name=str(shape.Name),
            shape_type=int(shape.Type),
            placeholder_type=placeholder_type,
            has_text=has_text,
            text_preview=text_preview,
            left=float(shape.Left),
            top=float(shape.Top),
            width=float(shape.Width),
            height=float(shape.Height),
            fill_color=fill_color,
            fill_transparency=fill_transparency,
            line_color=line_color,
            line_transparency=line_transparency,
            line_weight=line_weight,
            text_color=text_color,
            font_name=font_name,
            font_size=font_size,
            font_bold=font_bold,
            font_italic=font_italic,
        )

    def _table_summary(self, shape: object, index: int) -> TableSummary:
        table = shape.Table
        rows = int(table.Rows.Count)
        columns = int(table.Columns.Count)
        cells: list[list[str]] = []
        for row_index in range(1, rows + 1):
            row_values: list[str] = []
            for column_index in range(1, columns + 1):
                cell_text = ""
                with suppress(Exception):
                    cell_text = str(table.Cell(row_index, column_index).Shape.TextFrame.TextRange.Text).strip()
                row_values.append(cell_text)
            cells.append(row_values)

        return TableSummary(
            shape_index=index,
            shape_name=str(shape.Name),
            rows=rows,
            columns=columns,
            cells=cells,
        )

    def _chart_summary(self, shape: object, index: int) -> ChartSummary:
        chart = shape.Chart
        chart_title = None
        with suppress(Exception):
            if chart.HasTitle:
                chart_title = str(chart.ChartTitle.Text).strip() or None

        legend_visible = None
        legend_position_id = None
        legend_position_name = None
        with suppress(Exception):
            legend_visible = bool(chart.HasLegend)
            if legend_visible:
                legend_position_id = int(chart.Legend.Position)
                legend_position_name = alias_for_value(legend_position_id, CHART_LEGEND_POSITIONS)

        category_axis_title = None
        value_axis_title = None
        with suppress(Exception):
            axis = chart.Axes(1)
            if axis.HasTitle:
                category_axis_title = str(axis.AxisTitle.Text).strip() or None
        with suppress(Exception):
            axis = chart.Axes(2)
            if axis.HasTitle:
                value_axis_title = str(axis.AxisTitle.Text).strip() or None

        series_summaries: list[ChartSeriesSummary] = []
        with suppress(Exception):
            for series_index in range(1, int(chart.SeriesCollection().Count) + 1):
                series = chart.SeriesCollection(series_index)
                name = None
                with suppress(Exception):
                    name = str(series.Name).strip() or None

                values = normalize_office_value(series.Values)
                categories = normalize_office_value(series.XValues)
                series_summaries.append(
                    ChartSeriesSummary(
                        series_index=series_index,
                        name=name,
                        values=values if isinstance(values, list) else ([values] if values is not None else []),
                        categories=categories
                        if isinstance(categories, list)
                        else ([categories] if categories is not None else []),
                    )
                )

        chart_type = None
        with suppress(Exception):
            chart_type = int(chart.ChartType)

        return ChartSummary(
            shape_index=index,
            shape_name=str(shape.Name),
            chart_type=chart_type,
            chart_title=chart_title,
            legend_visible=legend_visible,
            legend_position_id=legend_position_id,
            legend_position_name=legend_position_name,
            category_axis_title=category_axis_title,
            value_axis_title=value_axis_title,
            series=series_summaries,
        )

    def _smartart_summary(self, shape: object, index: int) -> SmartArtSummary:
        smartart = shape.SmartArt
        nodes: list[SmartArtNodeSummary] = []
        for node_index in range(1, int(smartart.AllNodes.Count) + 1):
            node = smartart.AllNodes(node_index)
            text = ""
            level = None
            with suppress(Exception):
                text = str(node.TextFrame2.TextRange.Text).strip()
            with suppress(Exception):
                level = int(node.Level)
            nodes.append(SmartArtNodeSummary(node_index=node_index, text=text, level=level))

        return SmartArtSummary(
            shape_index=index,
            shape_name=str(shape.Name),
            node_count=len(nodes),
            nodes=nodes,
        )

    def _get_shape(self, slide: object, shape_index: int) -> object:
        return slide.Shapes(shape_index)

    def _get_table_cell_shape(self, table: object, row_index: int, column_index: int) -> object:
        rows = int(table.Rows.Count)
        columns = int(table.Columns.Count)
        if not 1 <= row_index <= rows:
            raise ValueError(f"The table does not contain row_index={row_index}")
        if not 1 <= column_index <= columns:
            raise ValueError(f"The table does not contain column_index={column_index}")
        return table.Cell(row_index, column_index).Shape

    def _resolve_animation_targets(
        self,
        shape: object,
        target_kind: str,
        row_index: int | None,
        column_index: int | None,
    ) -> list[object]:
        normalized_target = resolve_animation_target(target_kind)
        if normalized_target == "shape":
            return [shape]
        if normalized_target == "text":
            self._get_text_range(shape)
            return [shape]
        if normalized_target == "table":
            self._require_table_shape(shape)
            return [shape]

        table = self._require_table_shape(shape)
        row_count = int(table.Rows.Count)
        column_count = int(table.Columns.Count)

        if normalized_target == "table_cell":
            return [self._get_table_cell_shape(table, int(row_index), int(column_index))]

        if normalized_target == "table_row":
            return [self._get_table_cell_shape(table, int(row_index), index) for index in range(1, column_count + 1)]

        if normalized_target == "table_column":
            return [self._get_table_cell_shape(table, index, int(column_index)) for index in range(1, row_count + 1)]

        return [
            self._get_table_cell_shape(table, row, column)
            for row in range(1, row_count + 1)
            for column in range(1, column_count + 1)
        ]

    def _get_text_range(self, shape: object) -> object:
        if not (shape.HasTextFrame and shape.TextFrame is not None):
            raise ValueError("The selected shape does not support text")
        return shape.TextFrame.TextRange

    def _require_table_shape(self, shape: object) -> object:
        if not self._shape_has_table(shape):
            raise ValueError("The selected shape does not contain a table")
        return shape.Table

    def _require_chart_shape(self, shape: object) -> object:
        if not self._shape_has_chart(shape):
            raise ValueError("The selected shape does not contain a chart")
        return shape.Chart

    def _require_smartart_shape(self, shape: object) -> object:
        if not self._shape_has_smartart(shape):
            raise ValueError("The selected shape does not contain SmartArt")
        return shape.SmartArt

    def _get_chart_series(self, chart: object, series_index: int) -> object:
        series_collection = chart.SeriesCollection()
        if series_index > int(series_collection.Count):
            raise ValueError(f"The chart does not contain series_index={series_index}")
        return series_collection(series_index)

    def _normalize_chart_series_item(self, series_item: object) -> tuple[str, tuple[object, ...]]:
        if isinstance(series_item, dict):
            return str(series_item["name"]), tuple(series_item["values"])
        return str(series_item.name), tuple(series_item.values)

    def _apply_chart_series_data(
        self,
        chart: object,
        categories: list[object],
        series: list[object],
        replace_existing: bool,
    ) -> None:
        if not series:
            return

        series_collection = chart.SeriesCollection()
        if replace_existing:
            for index in range(int(series_collection.Count), 0, -1):
                with suppress(Exception):
                    series_collection(index).Delete()

        categories_tuple = tuple(categories) if categories else None
        existing_count = int(series_collection.Count)
        for index, series_item in enumerate(series, start=1):
            series_name, series_values = self._normalize_chart_series_item(series_item)
            if replace_existing or index > existing_count:
                current_series = series_collection.NewSeries()
            else:
                current_series = series_collection(index)

            current_series.Name = series_name
            if categories_tuple is not None:
                current_series.XValues = categories_tuple
            current_series.Values = series_values

        if not replace_existing and existing_count > len(series) and categories_tuple is not None:
            for index in range(len(series) + 1, existing_count + 1):
                with suppress(Exception):
                    series_collection(index).XValues = categories_tuple

    def _apply_shape_text_format(
        self,
        shape: object,
        text: str | None,
        text_color: str | None,
        font_name: str | None,
        font_size: float | None,
    ) -> None:
        if text is None and text_color is None and font_name is None and font_size is None:
            return

        text_range = self._get_text_range(shape)
        if text is not None:
            text_range.Text = text
        font = text_range.Font
        if text_color is not None:
            font.Color.RGB = parse_office_color(text_color)
        if font_name is not None:
            font.Name = font_name
        if font_size is not None:
            font.Size = font_size

    def _resolve_smartart_layout(self, powerpoint: object, layout: str) -> object:
        identifier = resolve_smartart_layout_identifier(layout)
        collection = powerpoint.SmartArtLayouts
        if isinstance(identifier, int):
            return collection(identifier)

        target = identifier.strip().lower()
        for index in range(1, int(collection.Count) + 1):
            candidate = collection(index)
            with suppress(Exception):
                if str(candidate.Id).strip().lower() == target:
                    return candidate
            with suppress(Exception):
                normalized_name = str(candidate.Name).strip().lower().replace("-", "_").replace(" ", "_")
                if normalized_name == target:
                    return candidate

        supported = ", ".join(sorted(SMARTART_LAYOUTS))
        raise ValueError(f"Unsupported SmartArt layout: {layout}. Supported values: {supported}")

    def _set_title_text(self, slide: object, title: str) -> None:
        with suppress(Exception):
            if slide.Shapes.HasTitle:
                slide.Shapes.Title.TextFrame.TextRange.Text = title
                return
        slide.Shapes.AddTextbox(1, 72, 36, 600, 40).TextFrame.TextRange.Text = title

    def _set_body_text(self, slide: object, body_text: str) -> None:
        title_name = None
        with suppress(Exception):
            if slide.Shapes.HasTitle:
                title_name = str(slide.Shapes.Title.Name)

        for index in range(1, int(slide.Shapes.Count) + 1):
            shape = slide.Shapes(index)
            with suppress(Exception):
                if title_name and str(shape.Name) == title_name:
                    continue
                if shape.HasTextFrame:
                    shape.TextFrame.TextRange.Text = body_text
                    return

        slide.Shapes.AddTextbox(1, 72, 140, 600, 240).TextFrame.TextRange.Text = body_text

    def _get_notes_text_shape(self, slide: object, create_if_missing: bool = False) -> object | None:
        notes_page = slide.NotesPage

        with suppress(Exception):
            placeholder = notes_page.Shapes.Placeholders(2)
            if placeholder.HasTextFrame:
                return placeholder

        fallback_shape = None
        for index in range(1, int(notes_page.Shapes.Count) + 1):
            shape = notes_page.Shapes(index)
            with suppress(Exception):
                if not shape.HasTextFrame:
                    continue

                placeholder_type = None
                with suppress(Exception):
                    placeholder_type = int(shape.PlaceholderFormat.Type)

                if placeholder_type == 2:
                    return shape
                if fallback_shape is None:
                    fallback_shape = shape

        if fallback_shape is not None:
            return fallback_shape

        if not create_if_missing:
            return None

        textbox = notes_page.Shapes.AddTextbox(1, 36, 72, 648, 360)
        with suppress(Exception):
            textbox.Name = f"SlideNotes_{int(slide.SlideIndex)}"
        return textbox

    def _get_slide_notes_text(self, slide: object) -> str:
        shape = self._get_notes_text_shape(slide, create_if_missing=False)
        if shape is None:
            return ""

        with suppress(Exception):
            return str(shape.TextFrame.TextRange.Text)
        return ""

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def list_slides(self, path: str) -> PresentationSummary:
        source = self.resolve_document_path(path)

        with self._open_presentation(source, read_only=True) as presentation:
            slides: list[SlideSummary] = []

            for index in range(1, int(presentation.Slides.Count) + 1):
                slide = presentation.Slides(index)
                slides.append(
                    SlideSummary(
                        slide_index=index,
                        title=self._get_slide_title(slide),
                        shape_count=int(slide.Shapes.Count),
                    )
                )

            return PresentationSummary(
                file_path=str(source),
                slide_count=int(presentation.Slides.Count),
                slides=slides,
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def get_slide_shapes(self, path: str, slide_index: int) -> SlideShapesResult:
        source = self.resolve_document_path(path)

        with self._open_presentation(source, read_only=True) as presentation:
            slide = presentation.Slides(slide_index)
            shapes = [self._shape_summary(shape, index) for index, shape in self._iter_shapes(slide)]
            return SlideShapesResult(file_path=str(source), slide_index=slide_index, shapes=shapes)

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def get_slide_text(self, path: str, slide_index: int) -> SlideTextResult:
        source = self.resolve_document_path(path)

        with self._open_presentation(source, read_only=True) as presentation:
            slide = presentation.Slides(slide_index)
            return SlideTextResult(
                file_path=str(source),
                slide_index=slide_index,
                texts=self._get_slide_texts(slide),
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def get_slide_notes(self, path: str, slide_index: int) -> SlideNotesResult:
        source = self.resolve_document_path(path)

        with self._open_presentation(source, read_only=True) as presentation:
            slide = presentation.Slides(slide_index)
            return SlideNotesResult(
                file_path=str(source),
                slide_index=slide_index,
                notes_text=self._get_slide_notes_text(slide),
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def get_slide_transition(self, path: str, slide_index: int) -> SlideTransitionResult:
        source = self.resolve_document_path(path)

        with self._open_presentation(source, read_only=True) as presentation:
            slide = presentation.Slides(slide_index)
            transition = slide.SlideShowTransition
            effect_id = int(transition.EntryEffect)
            speed_id = int(transition.Speed)
            advance_on_click = bool(transition.AdvanceOnClick)
            advance_after_seconds = None
            if bool(transition.AdvanceOnTime):
                advance_after_seconds = float(transition.AdvanceTime)

            return SlideTransitionResult(
                file_path=str(source),
                slide_index=slide_index,
                effect_id=effect_id,
                effect_name=alias_for_value(effect_id, TRANSITION_EFFECTS),
                speed_id=speed_id,
                speed_name=alias_for_value(speed_id, TRANSITION_SPEEDS),
                advance_on_click=advance_on_click,
                advance_after_seconds=advance_after_seconds,
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def get_slide_animations(self, path: str, slide_index: int) -> SlideAnimationsResult:
        source = self.resolve_document_path(path)

        with self._open_presentation(source, read_only=True) as presentation:
            slide = presentation.Slides(slide_index)
            sequence = slide.TimeLine.MainSequence
            animations: list[AnimationSummary] = []
            for animation_index in range(1, int(sequence.Count) + 1):
                effect = sequence(animation_index)
                shape_name = None
                effect_id = None
                trigger_id = None
                with suppress(Exception):
                    shape_name = str(effect.Shape.Name)
                with suppress(Exception):
                    effect_id = int(effect.EffectType)
                with suppress(Exception):
                    trigger_id = int(effect.Timing.TriggerType)

                duration_seconds = None
                delay_seconds = None
                with suppress(Exception):
                    duration_seconds = float(effect.Timing.Duration)
                with suppress(Exception):
                    delay_seconds = float(effect.Timing.TriggerDelayTime)

                animations.append(
                    AnimationSummary(
                        animation_index=animation_index,
                        shape_name=shape_name,
                        effect_id=effect_id,
                        effect_name=alias_for_value(effect_id, ANIMATION_EFFECTS),
                        trigger_id=trigger_id,
                        trigger_name=alias_for_value(trigger_id, ANIMATION_TRIGGERS),
                        duration_seconds=duration_seconds,
                        delay_seconds=delay_seconds,
                    )
                )

            return SlideAnimationsResult(
                file_path=str(source),
                slide_index=slide_index,
                animation_count=len(animations),
                animations=animations,
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def apply_style_preset(self, path: str, slide_index: int, preset: str, create_backup: bool) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)
        style = resolve_style_preset(preset)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            slide.FollowMasterBackground = 0
            slide.Background.Fill.Visible = -1
            slide.Background.Fill.Solid()
            slide.Background.Fill.ForeColor.RGB = parse_office_color(str(style["background"]))

            transition = slide.SlideShowTransition
            transition.EntryEffect = int(TRANSITION_EFFECTS[str(style["transition_effect"])])
            transition.Speed = int(TRANSITION_SPEEDS[str(style["transition_speed"])])

            title_shape = self._find_title_shape(slide)
            body_shape = self._find_primary_body_shape(slide)
            title_shape_name = None
            body_shape_name = None

            if title_shape is not None:
                title_shape_name = str(title_shape.Name)
                title_shape.Fill.Visible = -1
                title_shape.Fill.Solid()
                title_shape.Fill.ForeColor.RGB = parse_office_color(str(style["title_fill"]))
                title_shape.Line.Visible = -1
                title_shape.Line.ForeColor.RGB = parse_office_color(str(style["title_line"]))

                title_text = self._get_text_range(title_shape)
                title_font = title_text.Font
                title_font.Color.RGB = parse_office_color(str(style["title_text_color"]))
                title_font.Name = str(style["title_font_name"])
                title_font.Size = float(style["title_font_size"])
                title_font.Bold = -1 if bool(style["title_bold"]) else 0

            if body_shape is not None:
                body_shape_name = str(body_shape.Name)
                body_text = self._get_text_range(body_shape)
                body_font = body_text.Font
                body_font.Color.RGB = parse_office_color(str(style["body_text_color"]))
                body_font.Name = str(style["body_font_name"])
                body_font.Size = float(style["body_font_size"])
                body_text.ParagraphFormat.Alignment = int(TEXT_ALIGNMENTS[str(style["body_alignment"])])

            presentation.Save()
            return OperationResult(
                message="PowerPoint style preset applied",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "preset": preset,
                    "title_shape": title_shape_name,
                    "body_shape": body_shape_name,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def get_slide_tables(self, path: str, slide_index: int) -> SlideTablesResult:
        source = self.resolve_document_path(path)

        with self._open_presentation(source, read_only=True) as presentation:
            slide = presentation.Slides(slide_index)
            tables = [
                self._table_summary(shape, index)
                for index, shape in self._iter_shapes(slide)
                if self._shape_has_table(shape)
            ]
            return SlideTablesResult(file_path=str(source), slide_index=slide_index, tables=tables)

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_table_cell_text(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        row_index: int,
        column_index: int,
        text: str,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            table = self._require_table_shape(self._get_shape(slide, shape_index))
            table.Cell(row_index, column_index).Shape.TextFrame.TextRange.Text = text
            presentation.Save()
            return OperationResult(
                message="PowerPoint table cell updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "row_index": row_index,
                    "column_index": column_index,
                    "text": text,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def add_table(
        self,
        path: str,
        slide_index: int,
        rows: int,
        columns: int,
        x: float,
        y: float,
        width: float,
        height: float,
        values: list[list[object]],
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = slide.Shapes.AddTable(rows, columns, x, y, width, height)
            table = shape.Table
            for row_index, row_values in enumerate(values, start=1):
                if row_index > rows:
                    break
                for column_index, cell_value in enumerate(row_values, start=1):
                    if column_index > columns:
                        break
                    table.Cell(row_index, column_index).Shape.TextFrame.TextRange.Text = (
                        "" if cell_value is None else str(cell_value)
                    )

            presentation.Save()
            return OperationResult(
                message="PowerPoint table added",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_name": str(shape.Name),
                    "rows": rows,
                    "columns": columns,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def get_slide_charts(self, path: str, slide_index: int) -> SlideChartsResult:
        source = self.resolve_document_path(path)

        with self._open_presentation(source, read_only=True) as presentation:
            slide = presentation.Slides(slide_index)
            charts = [
                self._chart_summary(shape, index)
                for index, shape in self._iter_shapes(slide)
                if self._shape_has_chart(shape)
            ]
            return SlideChartsResult(file_path=str(source), slide_index=slide_index, charts=charts)

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_chart_title(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        title: str,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            chart = self._require_chart_shape(self._get_shape(slide, shape_index))
            chart.HasTitle = True
            chart.ChartTitle.Text = title
            presentation.Save()
            return OperationResult(
                message="PowerPoint chart title updated",
                file_path=str(source),
                backup_path=backup_path,
                details={"slide_index": slide_index, "shape_index": shape_index, "title": title},
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_chart_data(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        chart_type: str | None,
        title: str | None,
        categories: list[object],
        series: list[object],
        replace_existing: bool,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation_session(source, read_only=False, visible=True, with_window=True) as (_, presentation):
            slide = presentation.Slides(slide_index)
            chart = self._require_chart_shape(self._get_shape(slide, shape_index))

            if chart_type is not None:
                chart.ChartType = resolve_chart_type(chart_type)
            if title is not None:
                chart.HasTitle = True
                chart.ChartTitle.Text = title
            self._apply_chart_series_data(chart, categories, series, replace_existing)

            presentation.Save()
            return OperationResult(
                message="PowerPoint chart data updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "chart_type": int(chart.ChartType),
                    "title": str(chart.ChartTitle.Text).strip() if bool(chart.HasTitle) else None,
                    "series_count": int(chart.SeriesCollection().Count),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_chart_series_style(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        series_index: int,
        fill_color: str | None,
        line_color: str | None,
        show_data_labels: bool | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            chart = self._require_chart_shape(self._get_shape(slide, shape_index))
            series = self._get_chart_series(chart, series_index)

            if fill_color is not None:
                series.Format.Fill.Visible = -1
                series.Format.Fill.Solid()
                series.Format.Fill.ForeColor.RGB = parse_office_color(fill_color)
            if line_color is not None:
                series.Format.Line.Visible = -1
                series.Format.Line.ForeColor.RGB = parse_office_color(line_color)
            if show_data_labels is not None:
                if show_data_labels:
                    series.ApplyDataLabels()
                else:
                    series.HasDataLabels = False

            presentation.Save()
            return OperationResult(
                message="PowerPoint chart series style updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "series_index": series_index,
                    "fill_color": fill_color,
                    "line_color": line_color,
                    "show_data_labels": show_data_labels,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_chart_layout(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        legend_visible: bool | None,
        legend_position: str | None,
        category_axis_title: str | None,
        value_axis_title: str | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation_session(source, read_only=False, visible=True, with_window=True) as (_, presentation):
            slide = presentation.Slides(slide_index)
            chart = self._require_chart_shape(self._get_shape(slide, shape_index))

            if legend_visible is not None:
                chart.HasLegend = bool(legend_visible)
            if legend_position is not None:
                chart.HasLegend = True
                chart.Legend.Position = resolve_chart_legend_position(legend_position)

            if category_axis_title is not None:
                category_axis = chart.Axes(1)
                category_axis.HasTitle = bool(category_axis_title)
                if category_axis_title:
                    category_axis.AxisTitle.Text = category_axis_title

            if value_axis_title is not None:
                value_axis = chart.Axes(2)
                value_axis.HasTitle = bool(value_axis_title)
                if value_axis_title:
                    value_axis.AxisTitle.Text = value_axis_title

            presentation.Save()
            return OperationResult(
                message="PowerPoint chart layout updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "legend_visible": bool(chart.HasLegend),
                    "legend_position": alias_for_value(
                        int(chart.Legend.Position) if bool(chart.HasLegend) else None,
                        CHART_LEGEND_POSITIONS,
                    ),
                    "category_axis_title": category_axis_title,
                    "value_axis_title": value_axis_title,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def add_chart(
        self,
        path: str,
        slide_index: int,
        chart_type: str,
        x: float,
        y: float,
        width: float,
        height: float,
        title: str | None,
        categories: list[object],
        series: list[object],
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation_session(source, read_only=False, visible=True, with_window=True) as (_, presentation):
            slide = presentation.Slides(slide_index)
            shape = slide.Shapes.AddChart(resolve_chart_type(chart_type), x, y, width, height)
            chart = shape.Chart

            if title:
                chart.HasTitle = True
                chart.ChartTitle.Text = title

            if series:
                series_collection = chart.SeriesCollection()
                for index in range(int(series_collection.Count), 0, -1):
                    with suppress(Exception):
                        series_collection(index).Delete()

                categories_tuple = tuple(categories) if categories else None
                for series_item in series:
                    if isinstance(series_item, dict):
                        series_name = str(series_item["name"])
                        series_values = tuple(series_item["values"])
                    else:
                        series_name = str(series_item.name)
                        series_values = tuple(series_item.values)

                    created_series = series_collection.NewSeries()
                    created_series.Name = series_name
                    if categories_tuple is not None:
                        created_series.XValues = categories_tuple
                    created_series.Values = series_values

            presentation.Save()
            return OperationResult(
                message="PowerPoint chart added",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_name": str(shape.Name),
                    "chart_type": resolve_chart_type(chart_type),
                    "title": title,
                    "series_count": int(chart.SeriesCollection().Count),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def get_slide_smartart(self, path: str, slide_index: int) -> SlideSmartArtResult:
        source = self.resolve_document_path(path)

        with self._open_presentation(source, read_only=True) as presentation:
            slide = presentation.Slides(slide_index)
            smartart_items = [
                self._smartart_summary(shape, index)
                for index, shape in self._iter_shapes(slide)
                if self._shape_has_smartart(shape)
            ]
            return SlideSmartArtResult(
                file_path=str(source),
                slide_index=slide_index,
                smartart_items=smartart_items,
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_smartart_node_text(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        node_index: int,
        text: str,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            smartart = self._require_smartart_shape(self._get_shape(slide, shape_index))
            node = smartart.AllNodes(node_index)
            node.TextFrame2.TextRange.Text = text
            presentation.Save()
            return OperationResult(
                message="PowerPoint SmartArt node updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "node_index": node_index,
                    "text": text,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def add_smartart(
        self,
        path: str,
        slide_index: int,
        layout: str,
        x: float,
        y: float,
        width: float,
        height: float,
        node_texts: list[str],
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation_session(source, read_only=False, visible=True, with_window=True) as (powerpoint, presentation):
            slide = presentation.Slides(slide_index)
            layout_object = self._resolve_smartart_layout(powerpoint, layout)
            shape = slide.Shapes.AddSmartArt(layout_object, x, y, width, height)
            smartart = shape.SmartArt

            for node_index, node_text in enumerate(node_texts, start=1):
                while int(smartart.AllNodes.Count) < node_index:
                    smartart.AllNodes.Add()
                smartart.AllNodes(node_index).TextFrame2.TextRange.Text = node_text

            presentation.Save()
            return OperationResult(
                message="PowerPoint SmartArt added",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_name": str(shape.Name),
                    "layout": resolve_smartart_layout_identifier(layout),
                    "node_count": int(smartart.AllNodes.Count),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def add_shape(
        self,
        path: str,
        slide_index: int,
        shape_type: str,
        x: float,
        y: float,
        width: float,
        height: float,
        text: str | None,
        fill_color: str | None,
        line_color: str | None,
        line_weight: float | None,
        text_color: str | None,
        font_name: str | None,
        font_size: float | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = slide.Shapes.AddShape(resolve_shape_type(shape_type), x, y, width, height)

            if fill_color is not None:
                shape.Fill.Visible = -1
                shape.Fill.Solid()
                shape.Fill.ForeColor.RGB = parse_office_color(fill_color)
            if line_color is not None:
                shape.Line.Visible = -1
                shape.Line.ForeColor.RGB = parse_office_color(line_color)
            if line_weight is not None:
                shape.Line.Weight = line_weight
            self._apply_shape_text_format(shape, text, text_color, font_name, font_size)

            presentation.Save()
            return OperationResult(
                message="PowerPoint shape added",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_name": str(shape.Name),
                    "shape_type": resolve_shape_type(shape_type),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def add_connector(
        self,
        path: str,
        slide_index: int,
        connector_type: str,
        begin_x: float,
        begin_y: float,
        end_x: float,
        end_y: float,
        line_color: str | None,
        line_weight: float | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = slide.Shapes.AddConnector(
                resolve_connector_type(connector_type),
                begin_x,
                begin_y,
                end_x,
                end_y,
            )
            if line_color is not None:
                shape.Line.Visible = -1
                shape.Line.ForeColor.RGB = parse_office_color(line_color)
            if line_weight is not None:
                shape.Line.Weight = line_weight

            presentation.Save()
            return OperationResult(
                message="PowerPoint connector added",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_name": str(shape.Name),
                    "connector_type": resolve_connector_type(connector_type),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def connect_shapes(
        self,
        path: str,
        slide_index: int,
        start_shape_index: int,
        end_shape_index: int,
        begin_connection_site: int,
        end_connection_site: int,
        connector_type: str,
        line_color: str | None,
        line_weight: float | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            start_shape = self._get_shape(slide, start_shape_index)
            end_shape = self._get_shape(slide, end_shape_index)
            connector = slide.Shapes.AddConnector(resolve_connector_type(connector_type), 0, 0, 0, 0)
            connector.ConnectorFormat.BeginConnect(start_shape, begin_connection_site)
            connector.ConnectorFormat.EndConnect(end_shape, end_connection_site)
            with suppress(Exception):
                connector.RerouteConnections()
            if line_color is not None:
                connector.Line.Visible = -1
                connector.Line.ForeColor.RGB = parse_office_color(line_color)
            if line_weight is not None:
                connector.Line.Weight = line_weight

            presentation.Save()
            return OperationResult(
                message="PowerPoint shapes connected",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_name": str(connector.Name),
                    "start_shape_index": start_shape_index,
                    "end_shape_index": end_shape_index,
                    "begin_connection_site": begin_connection_site,
                    "end_connection_site": end_connection_site,
                    "connector_type": resolve_connector_type(connector_type),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def replace_text(
        self,
        path: str,
        slide_index: int,
        find_text: str,
        replace_text: str,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            replacement_count = 0

            for index in range(1, int(slide.Shapes.Count) + 1):
                shape = slide.Shapes(index)
                with suppress(Exception):
                    if shape.HasTextFrame and shape.TextFrame.HasText:
                        current_text = str(shape.TextFrame.TextRange.Text)
                        occurrences = current_text.count(find_text)
                        if occurrences:
                            shape.TextFrame.TextRange.Text = current_text.replace(find_text, replace_text)
                            replacement_count += occurrences

            presentation.Save()
            return OperationResult(
                message="PowerPoint slide text replacement completed",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "find_text": find_text,
                    "replace_text": replace_text,
                    "replacement_count": replacement_count,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_slide_notes(
        self,
        path: str,
        slide_index: int,
        text: str,
        append: bool,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_notes_text_shape(slide, create_if_missing=True)
            if shape is None:
                raise ValueError("Could not access the notes area for the selected slide")

            text_range = shape.TextFrame.TextRange
            previous_text = ""
            with suppress(Exception):
                previous_text = str(text_range.Text)

            next_text = text
            if append and previous_text:
                separator = "\r\n" if not previous_text.endswith(("\r\n", "\n", "\r")) and text else ""
                next_text = f"{previous_text}{separator}{text}"

            text_range.Text = next_text

            presentation.Save()
            return OperationResult(
                message="PowerPoint slide notes updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "append": append,
                    "previous_length": len(previous_text),
                    "notes_length": len(next_text),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def add_slide(
        self,
        path: str,
        layout: str,
        position: int | None,
        title: str | None,
        body_text: str | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            insert_position = position or int(presentation.Slides.Count) + 1
            slide = presentation.Slides.Add(insert_position, resolve_slide_layout(layout))

            if title:
                self._set_title_text(slide, title)
            if body_text:
                self._set_body_text(slide, body_text)

            presentation.Save()
            return OperationResult(
                message="PowerPoint slide added",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": int(slide.SlideIndex),
                    "slide_id": int(slide.SlideID),
                    "layout": layout,
                    "shape_count": int(slide.Shapes.Count),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def insert_image(
        self,
        path: str,
        slide_index: int,
        image_path: str,
        x: float,
        y: float,
        width: float | None,
        height: float | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        image = validate_file_path(
            image_path,
            allowed_roots=self.settings.allowed_roots,
            allowed_suffixes=IMAGE_SUFFIXES,
            must_exist=True,
        )
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = slide.Shapes.AddPicture(
                str(image),
                False,
                True,
                x,
                y,
                width if width is not None else -1,
                height if height is not None else -1,
            )
            presentation.Save()
            return OperationResult(
                message="Image inserted into PowerPoint slide",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "image_path": str(image),
                    "shape_name": str(shape.Name),
                    "left": float(shape.Left),
                    "top": float(shape.Top),
                    "width": float(shape.Width),
                    "height": float(shape.Height),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def apply_theme(self, path: str, theme_path: str, create_backup: bool) -> OperationResult:
        source = self.resolve_document_path(path)
        theme = validate_file_path(
            theme_path,
            allowed_roots=self.settings.allowed_roots,
            allowed_suffixes=THEME_SUFFIXES,
            must_exist=True,
        )
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            presentation.ApplyTheme(str(theme))
            presentation.Save()
            return OperationResult(
                message="PowerPoint theme applied",
                file_path=str(source),
                backup_path=backup_path,
                details={"theme_path": str(theme)},
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_slide_transition(
        self,
        path: str,
        slide_index: int,
        effect: str,
        speed: str,
        advance_on_click: bool,
        advance_after_seconds: float | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        from win32com.client import constants

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            transition = slide.SlideShowTransition
            transition.EntryEffect = self._resolve_office_constant(constants, effect, TRANSITION_EFFECTS)
            transition.Speed = self._resolve_office_constant(constants, speed, TRANSITION_SPEEDS)
            transition.AdvanceOnClick = -1 if advance_on_click else 0

            if advance_after_seconds is None:
                transition.AdvanceOnTime = 0
                transition.AdvanceTime = 0
            else:
                transition.AdvanceOnTime = -1
                transition.AdvanceTime = float(advance_after_seconds)

            presentation.Save()
            return OperationResult(
                message="PowerPoint slide transition updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "effect": effect,
                    "speed": speed,
                    "advance_on_click": advance_on_click,
                    "advance_after_seconds": advance_after_seconds,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_shape_text_style(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        text: str | None,
        font_name: str | None,
        font_size: float | None,
        bold: bool | None,
        italic: bool | None,
        underline: bool | None,
        color: str | None,
        alignment: str | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        from win32com.client import constants

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            text_range = self._get_text_range(shape)

            if text is not None:
                text_range.Text = text
            font = text_range.Font
            if font_name is not None:
                font.Name = font_name
            if font_size is not None:
                font.Size = font_size
            if bold is not None:
                font.Bold = -1 if bold else 0
            if italic is not None:
                font.Italic = -1 if italic else 0
            if underline is not None:
                font.Underline = -1 if underline else 0
            if color is not None:
                font.Color.RGB = parse_office_color(color)
            if alignment is not None:
                text_range.ParagraphFormat.Alignment = self._resolve_office_constant(
                    constants, alignment, TEXT_ALIGNMENTS
                )

            presentation.Save()
            return OperationResult(
                message="PowerPoint text style updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "text_updated": text is not None,
                    "font_name": font_name,
                    "font_size": font_size,
                    "bold": bold,
                    "italic": italic,
                    "underline": underline,
                    "color": color,
                    "alignment": alignment,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_shape_fill(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        color: str,
        transparency: float | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            shape.Fill.Visible = -1
            shape.Fill.Solid()
            shape.Fill.ForeColor.RGB = parse_office_color(color)
            if transparency is not None:
                shape.Fill.Transparency = transparency

            presentation.Save()
            return OperationResult(
                message="PowerPoint shape fill updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "color": color,
                    "transparency": transparency,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_shape_line(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        color: str | None,
        weight: float | None,
        transparency: float | None,
        visible: bool | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            if visible is not None:
                shape.Line.Visible = -1 if visible else 0
            if color is not None:
                shape.Line.Visible = -1
                shape.Line.ForeColor.RGB = parse_office_color(color)
            if weight is not None:
                shape.Line.Weight = weight
            if transparency is not None:
                shape.Line.Transparency = transparency

            presentation.Save()
            return OperationResult(
                message="PowerPoint shape line updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "color": color,
                    "weight": weight,
                    "transparency": transparency,
                    "visible": visible,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_slide_background(
        self,
        path: str,
        slide_index: int,
        color: str,
        follow_master: bool,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            slide.FollowMasterBackground = -1 if follow_master else 0
            if not follow_master:
                slide.Background.Fill.Visible = -1
                slide.Background.Fill.Solid()
                slide.Background.Fill.ForeColor.RGB = parse_office_color(color)

            presentation.Save()
            return OperationResult(
                message="PowerPoint slide background updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "color": color,
                    "follow_master": follow_master,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def add_shape_animation(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        effect: str,
        trigger: str,
        duration_seconds: float | None,
        delay_seconds: float | None,
        target_kind: str,
        animation_level: str | None,
        row_index: int | None,
        column_index: int | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        from win32com.client import constants

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            sequence = slide.TimeLine.MainSequence
            normalized_target = resolve_animation_target(target_kind)
            resolved_level = resolve_animation_level(
                animation_level if animation_level is not None else ("all_text_levels" if normalized_target == "text" else "none")
            )
            targets = self._resolve_animation_targets(shape, normalized_target, row_index, column_index)
            effect_id = self._resolve_office_constant(constants, effect, ANIMATION_EFFECTS)
            trigger_id = self._resolve_office_constant(constants, trigger, ANIMATION_TRIGGERS)

            for target_shape in targets:
                effect_object = sequence.AddEffect(target_shape, effect_id, resolved_level, trigger_id)

                if duration_seconds is not None:
                    effect_object.Timing.Duration = float(duration_seconds)
                if delay_seconds is not None:
                    effect_object.Timing.TriggerDelayTime = float(delay_seconds)

            presentation.Save()
            return OperationResult(
                message="PowerPoint shape animation added",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "target_kind": normalized_target,
                    "animation_level": resolved_level,
                    "row_index": row_index,
                    "column_index": column_index,
                    "effect": effect,
                    "trigger": trigger,
                    "duration_seconds": duration_seconds,
                    "delay_seconds": delay_seconds,
                    "targets_applied": len(targets),
                    "animation_count": int(sequence.Count),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def clear_slide_animations(self, path: str, slide_index: int, create_backup: bool) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            sequence = slide.TimeLine.MainSequence
            deleted = int(sequence.Count)
            for index in range(int(sequence.Count), 0, -1):
                sequence(index).Delete()

            presentation.Save()
            return OperationResult(
                message="PowerPoint slide animations cleared",
                file_path=str(source),
                backup_path=backup_path,
                details={"slide_index": slide_index, "deleted_count": deleted},
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def export_pdf(self, path: str, out_path: str) -> OperationResult:
        source = self.resolve_document_path(path)
        target = self.resolve_output_path(out_path, allowed_suffixes=(".pdf",))

        with self._open_presentation(source, read_only=True) as presentation:
            presentation.SaveAs(str(target), 32)
            return OperationResult(
                message="PowerPoint presentation exported to PDF",
                file_path=str(source),
                details={"out_path": str(target)},
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def export_slide_images(
        self,
        path: str,
        out_dir: str,
        image_format: str,
        width: int | None,
        height: int | None,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        target_dir = ensure_directory(out_dir)
        filter_name, extension = resolve_export_image_format(image_format)
        exported_files: list[str] = []

        with self._open_presentation(source, read_only=True) as presentation:
            for index in range(1, int(presentation.Slides.Count) + 1):
                slide = presentation.Slides(index)
                output_file = target_dir / f"slide-{index:03d}{extension}"
                if width is not None and height is not None:
                    slide.Export(str(output_file), filter_name, width, height)
                else:
                    slide.Export(str(output_file), filter_name)
                exported_files.append(str(output_file))

        return OperationResult(
            message="PowerPoint slides exported as images",
            file_path=str(source),
            details={
                "out_dir": str(target_dir),
                "image_format": image_format,
                "exported_count": len(exported_files),
                "exported_files": exported_files,
            },
        )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def save_as(self, path: str, out_path: str) -> OperationResult:
        source = self.resolve_document_path(path)
        target = self.resolve_output_path(out_path, allowed_suffixes=self.allowed_suffixes)

        with self._open_presentation(source, read_only=False) as presentation:
            try:
                presentation.SaveCopyAs(str(target))
            except Exception:
                presentation.SaveAs(str(target))

        return OperationResult(
            message="PowerPoint presentation saved as a new file",
            file_path=str(source),
            details={"out_path": str(target)},
        )
