from __future__ import annotations

import csv
import json
import re
from contextlib import contextmanager, suppress
from pathlib import Path
from typing import Iterator

from tenacity import retry, stop_after_attempt, wait_fixed

from office_ai_mcp.config import Settings
from office_ai_mcp.models.responses import (
    AnimationSummary,
    ChartDataExportResult,
    ChartSeriesSummary,
    ChartSummary,
    DocumentPropertiesResult,
    ExtendedSlideSummaryResult,
    FileLinksResult,
    LayoutSummary,
    MasterDetailsResult,
    MasterSummary,
    MasterThemeSummary,
    MediaSummary,
    OperationResult,
    PlaceholderSummary,
    PresentationLayoutsResult,
    PresentationMediaInventoryResult,
    PresentationMastersResult,
    PresentationNotesResult,
    PresentationSummary,
    PresentationSpellcheckResult,
    PresentationThemeResult,
    PresentationTextSearchResult,
    ShapeSummary,
    ShapeSearchResult,
    SlideAnimationsResult,
    SlideChartsResult,
    SlideLayoutResult,
    SlideMetadataResult,
    SlideMetadataSummary,
    SlideNotesResult,
    SlideSpellcheckResult,
    SlidePlaceholdersResult,
    SlideShapesResult,
    SlideSmartArtResult,
    SlideSummary,
    SlideTablesResult,
    SlideTextResult,
    SlideTransitionResult,
    ShapeTextRunsResult,
    SmartArtNodeSummary,
    SmartArtSummary,
    TableSummary,
    SpellingIssueSummary,
    TextRunSummary,
    TextMatchSummary,
    ThemeVariantSummary,
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
SVG_SUFFIXES = (".svg",)
VIDEO_SUFFIXES = (".mp4", ".mov", ".avi", ".wmv", ".m4v", ".mpeg", ".mpg")
AUDIO_SUFFIXES = (".mp3", ".wav", ".wma", ".m4a", ".aac", ".mid", ".midi")
CSV_SUFFIXES = (".csv",)
EXCEL_SUFFIXES = (".xls", ".xlsx", ".xlsm")
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
CHART_AXIS_TYPES = {
    "category": 1,
    "x": 1,
    "value": 2,
    "y": 2,
    "series": 3,
    "z": 3,
}
CHART_DATA_LABEL_POSITIONS = {
    "center": -4108,
    "inside_end": 2,
    "inside_base": 4,
    "outside_end": 3,
    "best_fit": 5,
    "above": 0,
    "below": 1,
    "left": -4131,
    "right": -4152,
}
CHART_PLOT_BY = {
    "rows": 1,
    "row": 1,
    "xlrows": 1,
    "columns": 2,
    "column": 2,
    "xlcolumns": 2,
}
SHAPE_ALIGNMENT_COMMANDS = {
    "left": 0,
    "center": 1,
    "middle": 4,
    "right": 2,
    "top": 3,
    "bottom": 5,
}
SHAPE_DISTRIBUTION_COMMANDS = {
    "horizontal": 0,
    "vertical": 1,
}
SHAPE_FLIP_COMMANDS = {
    "horizontal": 0,
    "vertical": 1,
}
SHAPE_Z_ORDER_COMMANDS = {
    "bring_to_front": 0,
    "send_to_back": 1,
    "bring_forward": 2,
    "send_backward": 3,
}
SHAPE_RESIZE_MODES = {"width", "height", "both"}
SHAPE_MERGE_COMMANDS = {
    "combine": 1,
    "union": 2,
    "intersect": 3,
    "subtract": 4,
    "fragment": 5,
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
GRADIENT_STYLES = {
    "horizontal": "msoGradientHorizontal",
    "vertical": "msoGradientVertical",
    "diagonal_up": "msoGradientDiagonalUp",
    "diagonal_down": "msoGradientDiagonalDown",
    "from_corner": "msoGradientFromCorner",
    "from_title": "msoGradientFromTitle",
    "from_center": "msoGradientFromCenter",
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
TEXT_DIRECTIONS = {
    "horizontal": "msoTextOrientationHorizontal",
    "vertical": "msoTextOrientationVertical",
    "upward": "msoTextOrientationUpward",
    "downward": "msoTextOrientationDownward",
    "vertical_far_east": "msoTextOrientationVerticalFarEast",
    "horizontal_rotated_far_east": "msoTextOrientationHorizontalRotatedFarEast",
}
TEXT_AUTOFIT_MODES = {
    "none": "ppAutoSizeNone",
    "shape_to_fit_text": "ppAutoSizeShapeToFitText",
}
PROOFING_LANGUAGE_ALIASES = {
    "english_us": "msoLanguageIDEnglishUS",
    "english_uk": "msoLanguageIDEnglishUK",
    "spanish": "msoLanguageIDSpanish",
    "mexican_spanish": "msoLanguageIDMexicanSpanish",
    "french": "msoLanguageIDFrench",
    "german": "msoLanguageIDGerman",
    "italian": "msoLanguageIDItalian",
    "portuguese": "msoLanguageIDPortuguese",
    "brazilian_portuguese": "msoLanguageIDBrazilianPortuguese",
    "no_proofing": "msoLanguageIDNoProofing",
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
SMARTART_NODE_POSITIONS = {
    "after": "msoSmartArtNodeAfter",
    "before": "msoSmartArtNodeBefore",
    "above": "msoSmartArtNodeAbove",
    "below": "msoSmartArtNodeBelow",
}
SMARTART_NODE_TYPES = {
    "default": "msoSmartArtNodeTypeDefault",
    "assistant": "msoSmartArtNodeTypeAssistant",
}
SMARTART_REORDER_DIRECTIONS = {"up", "down"}
SMARTART_CONVERT_COMMANDS = (
    "SmartArtConvertToShapes",
    "SmartArtDesignConvertToShapes",
    "ConvertToShapes",
)
TITLE_PLACEHOLDER_TYPES = {1, 3}
THEME_COLOR_CONSTANTS = {
    "background_1": "msoThemeBackground1",
    "text_1": "msoThemeText1",
    "background_2": "msoThemeBackground2",
    "text_2": "msoThemeText2",
    "accent_1": "msoThemeAccent1",
    "accent_2": "msoThemeAccent2",
    "accent_3": "msoThemeAccent3",
    "accent_4": "msoThemeAccent4",
    "accent_5": "msoThemeAccent5",
    "accent_6": "msoThemeAccent6",
    "hyperlink": "msoThemeHyperlink",
    "followed_hyperlink": "msoThemeFollowedHyperlink",
}
THEME_FONT_LANGUAGE_CONSTANTS = {
    "latin": "msoThemeLatin",
    "east_asian": "msoThemeEastAsian",
    "complex_script": "msoThemeComplexScript",
}
DESIGN_IDEAS_COMMANDS = (
    "DesignIdeas",
    "DesignerDesignIdeas",
    "Designer",
)
BUILTIN_DOCUMENT_PROPERTY_NAMES = {
    "author": "Author",
    "title": "Title",
    "subject": "Subject",
    "keywords": "Keywords",
    "comments": "Comments",
    "category": "Category",
    "company": "Company",
    "manager": "Manager",
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


def resolve_shape_alignment(alignment: str) -> int:
    normalized = normalize_powerpoint_token(alignment)
    if normalized.lstrip("-").isdigit():
        return int(normalized)
    if normalized not in SHAPE_ALIGNMENT_COMMANDS:
        supported = ", ".join(sorted(SHAPE_ALIGNMENT_COMMANDS))
        raise ValueError(f"Unsupported shape alignment: {alignment}. Supported values: {supported}")
    return SHAPE_ALIGNMENT_COMMANDS[normalized]


def resolve_shape_distribution(direction: str) -> int:
    normalized = normalize_powerpoint_token(direction)
    if normalized.lstrip("-").isdigit():
        return int(normalized)
    if normalized not in SHAPE_DISTRIBUTION_COMMANDS:
        supported = ", ".join(sorted(SHAPE_DISTRIBUTION_COMMANDS))
        raise ValueError(f"Unsupported shape distribution: {direction}. Supported values: {supported}")
    return SHAPE_DISTRIBUTION_COMMANDS[normalized]


def resolve_shape_flip(direction: str) -> int:
    normalized = normalize_powerpoint_token(direction)
    if normalized.lstrip("-").isdigit():
        return int(normalized)
    if normalized not in SHAPE_FLIP_COMMANDS:
        supported = ", ".join(sorted(SHAPE_FLIP_COMMANDS))
        raise ValueError(f"Unsupported shape flip direction: {direction}. Supported values: {supported}")
    return SHAPE_FLIP_COMMANDS[normalized]


def resolve_shape_z_order(command: str) -> int:
    normalized = normalize_powerpoint_token(command)
    if normalized.lstrip("-").isdigit():
        return int(normalized)
    if normalized not in SHAPE_Z_ORDER_COMMANDS:
        supported = ", ".join(sorted(SHAPE_Z_ORDER_COMMANDS))
        raise ValueError(f"Unsupported shape z-order command: {command}. Supported values: {supported}")
    return SHAPE_Z_ORDER_COMMANDS[normalized]


def resolve_shape_resize_mode(mode: str) -> str:
    normalized = normalize_powerpoint_token(mode)
    if normalized not in SHAPE_RESIZE_MODES:
        supported = ", ".join(sorted(SHAPE_RESIZE_MODES))
        raise ValueError(f"Unsupported shape resize mode: {mode}. Supported values: {supported}")
    return normalized


def resolve_shape_merge_mode(mode: str) -> int:
    normalized = normalize_powerpoint_token(mode)
    if normalized.lstrip("-").isdigit():
        return int(normalized)
    if normalized not in SHAPE_MERGE_COMMANDS:
        supported = ", ".join(sorted(SHAPE_MERGE_COMMANDS))
        raise ValueError(f"Unsupported shape merge mode: {mode}. Supported values: {supported}")
    return SHAPE_MERGE_COMMANDS[normalized]


def resolve_text_direction(direction: str) -> str | int:
    normalized = normalize_powerpoint_token(direction)
    if normalized.lstrip("-").isdigit():
        return int(normalized)
    if normalized not in TEXT_DIRECTIONS:
        supported = ", ".join(sorted(TEXT_DIRECTIONS))
        raise ValueError(f"Unsupported text direction: {direction}. Supported values: {supported}")
    return TEXT_DIRECTIONS[normalized]


def resolve_gradient_style(style: str) -> str | int:
    normalized = normalize_powerpoint_token(style)
    if normalized.lstrip("-").isdigit():
        return int(normalized)
    if normalized not in GRADIENT_STYLES:
        supported = ", ".join(sorted(GRADIENT_STYLES))
        raise ValueError(f"Unsupported gradient style: {style}. Supported values: {supported}")
    return GRADIENT_STYLES[normalized]


def resolve_text_autofit_mode(mode: str) -> str | int:
    normalized = normalize_powerpoint_token(mode)
    if normalized.lstrip("-").isdigit():
        return int(normalized)
    if normalized not in TEXT_AUTOFIT_MODES:
        supported = ", ".join(sorted(TEXT_AUTOFIT_MODES))
        raise ValueError(f"Unsupported text autofit mode: {mode}. Supported values: {supported}")
    return TEXT_AUTOFIT_MODES[normalized]


def resolve_proofing_language(language: str) -> str | int:
    normalized = normalize_powerpoint_token(language)
    if normalized.lstrip("-").isdigit():
        return int(normalized)
    return PROOFING_LANGUAGE_ALIASES.get(normalized, language.strip())


def resolve_chart_legend_position(position: str) -> int:
    normalized = position.strip().lower().replace("-", "_").replace(" ", "_")
    if normalized.lstrip("-").isdigit():
        return int(normalized)
    if normalized not in CHART_LEGEND_POSITIONS:
        supported = ", ".join(sorted(CHART_LEGEND_POSITIONS))
        raise ValueError(f"Unsupported chart legend position: {position}. Supported values: {supported}")
    return CHART_LEGEND_POSITIONS[normalized]


def resolve_chart_axis_kind(axis_kind: str) -> int:
    normalized = normalize_powerpoint_token(axis_kind)
    if normalized.lstrip("-").isdigit():
        return int(normalized)
    if normalized not in CHART_AXIS_TYPES:
        supported = ", ".join(sorted(CHART_AXIS_TYPES))
        raise ValueError(f"Unsupported chart axis kind: {axis_kind}. Supported values: {supported}")
    return CHART_AXIS_TYPES[normalized]


def resolve_chart_data_label_position(position: str) -> int:
    normalized = normalize_powerpoint_token(position)
    if normalized.lstrip("-").isdigit():
        return int(normalized)
    if normalized not in CHART_DATA_LABEL_POSITIONS:
        supported = ", ".join(sorted(CHART_DATA_LABEL_POSITIONS))
        raise ValueError(f"Unsupported chart data label position: {position}. Supported values: {supported}")
    return CHART_DATA_LABEL_POSITIONS[normalized]


def resolve_chart_plot_by(plot_by: str) -> int:
    normalized = normalize_powerpoint_token(plot_by)
    if normalized.lstrip("-").isdigit():
        return int(normalized)
    if normalized not in CHART_PLOT_BY:
        supported = ", ".join(sorted(CHART_PLOT_BY))
        raise ValueError(f"Unsupported chart plot_by value: {plot_by}. Supported values: {supported}")
    return CHART_PLOT_BY[normalized]


def resolve_named_constant_alias(value: str, aliases: dict[str, str | int]) -> str | int:
    normalized = value.strip().lower().replace("-", "_").replace(" ", "_")
    if normalized.lstrip("-").isdigit():
        return int(normalized)
    return aliases.get(normalized, value.strip())


def normalize_powerpoint_token(value: str) -> str:
    return value.strip().lower().replace("-", "_").replace(" ", "_")


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


def resolve_smartart_node_position(position: str) -> str | int:
    normalized = normalize_powerpoint_token(position)
    if normalized.lstrip("-").isdigit():
        return int(normalized)
    if normalized not in SMARTART_NODE_POSITIONS:
        supported = ", ".join(sorted(SMARTART_NODE_POSITIONS))
        raise ValueError(f"Unsupported SmartArt node position: {position}. Supported values: {supported}")
    return SMARTART_NODE_POSITIONS[normalized]


def resolve_smartart_node_type(node_type: str) -> str | int:
    normalized = normalize_powerpoint_token(node_type)
    if normalized.lstrip("-").isdigit():
        return int(normalized)
    if normalized not in SMARTART_NODE_TYPES:
        supported = ", ".join(sorted(SMARTART_NODE_TYPES))
        raise ValueError(f"Unsupported SmartArt node type: {node_type}. Supported values: {supported}")
    return SMARTART_NODE_TYPES[normalized]


def resolve_smartart_reorder_direction(direction: str) -> str:
    normalized = normalize_powerpoint_token(direction)
    if normalized not in SMARTART_REORDER_DIRECTIONS:
        supported = ", ".join(sorted(SMARTART_REORDER_DIRECTIONS))
        raise ValueError(f"Unsupported SmartArt reorder direction: {direction}. Supported values: {supported}")
    return normalized


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

    def _extract_spelling_issues(self, word: object, text: str, language_id: int | None) -> list[str]:
        document = None
        try:
            document = word.Documents.Add()
            document.Content.Text = text
            if language_id is not None:
                with suppress(Exception):
                    document.Content.LanguageID = language_id

            issues: list[str] = []
            with suppress(Exception):
                for index in range(1, int(document.SpellingErrors.Count) + 1):
                    issue_text = str(document.SpellingErrors(index).Text).strip()
                    if issue_text:
                        issues.append(issue_text)
            return issues
        finally:
            if document is not None:
                with suppress(Exception):
                    document.Close(SaveChanges=False)

    def _iter_slide_spell_fragments(
        self,
        slide: object,
        *,
        include_notes: bool,
    ) -> list[dict[str, object]]:
        fragments: list[dict[str, object]] = []

        for shape_index, shape in self._iter_shapes(slide):
            with suppress(Exception):
                if not (shape.HasTextFrame and shape.TextFrame.HasText):
                    continue
                text_range = shape.TextFrame.TextRange
                text = str(text_range.Text).strip()
                if not text:
                    continue
                language_id = None
                with suppress(Exception):
                    language_id = int(text_range.LanguageID)
                fragments.append(
                    {
                        "location": "shape_text",
                        "shape_index": shape_index,
                        "shape_name": str(shape.Name),
                        "text": text,
                        "language_id": language_id,
                    }
                )

        if include_notes:
            notes_shape = self._get_notes_text_shape(slide, create_if_missing=False)
            if notes_shape is not None:
                with suppress(Exception):
                    text_range = notes_shape.TextFrame.TextRange
                    text = str(text_range.Text).strip()
                    if text:
                        language_id = None
                        with suppress(Exception):
                            language_id = int(text_range.LanguageID)
                        fragments.append(
                            {
                                "location": "notes",
                                "shape_index": None,
                                "shape_name": None,
                                "text": text,
                                "language_id": language_id,
                            }
                        )

        return fragments

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

    def _shape_is_connector(self, shape: object) -> bool:
        with suppress(Exception):
            return bool(shape.Connector)
        return False

    def _shape_supports_text(self, shape: object) -> bool:
        with suppress(Exception):
            return bool(shape.HasTextFrame and shape.TextFrame is not None)
        return False

    def _get_placeholder_type(self, shape: object) -> int | None:
        placeholder_type = None
        with suppress(Exception):
            placeholder_type = int(shape.PlaceholderFormat.Type)
        return placeholder_type

    def _shape_is_title_placeholder(self, shape: object) -> bool:
        return self._get_placeholder_type(shape) in TITLE_PLACEHOLDER_TYPES

    def _extract_placeholder_summary(self, shape: object, index: int) -> PlaceholderSummary | None:
        placeholder_type = self._get_placeholder_type(shape)
        if placeholder_type is None:
            return None

        shape_summary = self._shape_summary(shape, index)
        return PlaceholderSummary(
            shape_index=index,
            name=shape_summary.name,
            placeholder_type=placeholder_type,
            has_text=shape_summary.has_text,
            text_preview=shape_summary.text_preview,
            left=shape_summary.left,
            top=shape_summary.top,
            width=shape_summary.width,
            height=shape_summary.height,
        )

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
                if self._shape_supports_text(shape):
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

        series_summaries = self._collect_chart_series_summaries(chart)

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

    def _get_picture_format(self, shape: object) -> object:
        with suppress(Exception):
            return shape.PictureFormat
        raise ValueError("The selected shape does not support picture formatting")

    def _get_media_format(self, shape: object) -> object:
        with suppress(Exception):
            return shape.MediaFormat
        raise ValueError("The selected shape does not support media formatting")

    def _get_play_settings(self, shape: object) -> object:
        with suppress(Exception):
            return shape.AnimationSettings.PlaySettings
        raise ValueError("The selected shape does not expose media playback settings")

    def _infer_media_kind(self, shape: object) -> str | None:
        media_type = None
        with suppress(Exception):
            media_type = int(shape.MediaType)
        if media_type == 1:
            return "sound"
        if media_type == 2:
            return "video"

        source_path = None
        with suppress(Exception):
            source_path = str(shape.LinkFormat.SourceFullName).strip() or None
        if source_path is None:
            with suppress(Exception):
                source_path = str(shape.AlternativeText).strip() or None
        if source_path:
            suffix = Path(source_path).suffix.lower()
            if suffix in AUDIO_SUFFIXES:
                return "sound"
            if suffix in VIDEO_SUFFIXES:
                return "video"
            if suffix in SVG_SUFFIXES:
                return "svg"
        return None

    def _media_summary(self, shape: object, slide_index: int, shape_index: int) -> MediaSummary | None:
        source_path = None
        linked = None
        with suppress(Exception):
            source_path = str(shape.LinkFormat.SourceFullName).strip() or None
            linked = bool(source_path)

        media_kind = self._infer_media_kind(shape)
        media_format = None
        with suppress(Exception):
            media_format = shape.MediaFormat

        volume = None
        trim_start = None
        trim_end = None
        if media_format is not None:
            with suppress(Exception):
                volume = float(media_format.Volume)
            with suppress(Exception):
                trim_start = int(media_format.StartPoint)
            with suppress(Exception):
                trim_end = int(media_format.EndPoint)

        if media_kind is None and source_path is None and media_format is None:
            return None

        return MediaSummary(
            slide_index=slide_index,
            shape_index=shape_index,
            shape_name=str(shape.Name),
            media_kind=media_kind,
            source_path=source_path,
            linked=linked,
            left=float(shape.Left),
            top=float(shape.Top),
            width=float(shape.Width),
            height=float(shape.Height),
            volume=volume,
            trim_start=trim_start,
            trim_end=trim_end,
        )

    def _resolve_shape_selector(
        self,
        slide: object,
        *,
        shape_index: int | None = None,
        shape_name: str | None = None,
    ) -> tuple[int, object]:
        if shape_index is not None:
            return shape_index, self._get_shape(slide, shape_index)

        assert shape_name is not None
        target_name = shape_name.strip().lower()
        for current_index, shape in self._iter_shapes(slide):
            with suppress(Exception):
                if str(shape.Name).strip().lower() == target_name:
                    return current_index, shape
        raise ValueError(f"The slide does not contain a shape named '{shape_name}'")

    def _get_shape_range(self, slide: object, shape_indexes: list[int]) -> object:
        return slide.Shapes.Range(list(shape_indexes))

    def _collect_shape_indexes(self, shape_range: object) -> list[int]:
        indexes: list[int] = []
        with suppress(Exception):
            for index in range(1, int(shape_range.Count) + 1):
                current_shape = shape_range(index)
                with suppress(Exception):
                    indexes.append(int(current_shape.Id))
        return indexes

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

    def _get_text_frame(self, shape: object) -> object:
        if not (shape.HasTextFrame and shape.TextFrame is not None):
            raise ValueError("The selected shape does not support text")
        return shape.TextFrame

    def _get_text_characters(self, text_range: object, start: int, length: int) -> object:
        try:
            total_characters = int(text_range.Characters().Count)
        except Exception as exc:
            raise ValueError("The selected text range does not expose character positions") from exc

        if start > total_characters:
            raise ValueError(f"Text start position {start} exceeds the shape text length {total_characters}")
        if start + length - 1 > total_characters:
            raise ValueError(
                f"Requested text range ({start}, {length}) exceeds the shape text length {total_characters}"
            )
        return text_range.Characters(start, length)

    def _get_paragraph_range(self, text_range: object, paragraph_index: int | None) -> object:
        if paragraph_index is None:
            return text_range
        try:
            paragraph_count = int(text_range.Paragraphs().Count)
        except Exception as exc:
            raise ValueError("The selected text range does not expose paragraphs") from exc

        if paragraph_index > paragraph_count:
            raise ValueError(f"The text does not contain paragraph_index={paragraph_index}")
        return text_range.Paragraphs(paragraph_index, 1)

    def _summarize_text_run(self, run: object, run_index: int, start: int) -> TextRunSummary:
        text = ""
        with suppress(Exception):
            text = str(run.Text)

        font = run.Font
        color = None
        with suppress(Exception):
            color = office_color_to_hex(int(font.Color.RGB))

        language_id = None
        with suppress(Exception):
            language_id = int(run.LanguageID)

        return TextRunSummary(
            run_index=run_index,
            start=start,
            length=len(text),
            text=text,
            font_name=str(font.Name).strip() or None if getattr(font, "Name", None) is not None else None,
            font_size=float(font.Size) if getattr(font, "Size", None) is not None else None,
            bold=bool(font.Bold) if getattr(font, "Bold", None) is not None else None,
            italic=bool(font.Italic) if getattr(font, "Italic", None) is not None else None,
            underline=bool(font.Underline) if getattr(font, "Underline", None) is not None else None,
            color=color,
            language_id=language_id,
        )

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

    def _get_smartart_node(self, smartart: object, node_index: int) -> object:
        node_count = int(smartart.AllNodes.Count)
        if node_index > node_count:
            raise ValueError(f"The SmartArt does not contain node_index={node_index}")
        return smartart.AllNodes(node_index)

    def _describe_smartart_collection_item(self, item: object) -> dict[str, object | None]:
        name = None
        item_id = None
        category = None
        description = None
        with suppress(Exception):
            name = str(item.Name).strip() or None
        with suppress(Exception):
            item_id = str(item.Id).strip() or None
        with suppress(Exception):
            category = str(item.Category).strip() or None
        with suppress(Exception):
            description = str(item.Description).strip() or None
        return {
            "name": name,
            "id": item_id,
            "category": category,
            "description": description,
        }

    def _get_chart_series(self, chart: object, series_index: int) -> object:
        series_collection = chart.SeriesCollection()
        if series_index > int(series_collection.Count):
            raise ValueError(f"The chart does not contain series_index={series_index}")
        return series_collection(series_index)

    def _normalize_chart_items(self, value: object) -> list[object]:
        normalized = normalize_office_value(value)
        if isinstance(normalized, list):
            return normalized
        if normalized is None:
            return []
        return [normalized]

    def _get_chart_axis(self, chart: object, axis_kind: str) -> object:
        axis_type = resolve_chart_axis_kind(axis_kind)
        try:
            return chart.Axes(axis_type)
        except Exception as exc:
            raise ValueError(f"The chart does not expose the requested axis: {axis_kind}") from exc

    def _collect_chart_series_summaries(self, chart: object) -> list[ChartSeriesSummary]:
        series_summaries: list[ChartSeriesSummary] = []
        with suppress(Exception):
            for series_index in range(1, int(chart.SeriesCollection().Count) + 1):
                series = chart.SeriesCollection(series_index)
                name = None
                with suppress(Exception):
                    name = str(series.Name).strip() or None

                series_summaries.append(
                    ChartSeriesSummary(
                        series_index=series_index,
                        name=name,
                        values=self._normalize_chart_items(series.Values),
                        categories=self._normalize_chart_items(series.XValues),
                    )
                )
        return series_summaries

    def _apply_chart_series_colors(self, series: object, fill_color: str | None, line_color: str | None) -> None:
        if fill_color is not None:
            series.Format.Fill.Visible = -1
            series.Format.Fill.Solid()
            series.Format.Fill.ForeColor.RGB = parse_office_color(fill_color)
        if line_color is not None:
            series.Format.Line.Visible = -1
            series.Format.Line.ForeColor.RGB = parse_office_color(line_color)

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

    def _apply_text_range_style(
        self,
        text_range: object,
        *,
        text: str | None,
        font_name: str | None,
        font_size: float | None,
        bold: bool | None,
        italic: bool | None,
        underline: bool | None,
        color: str | None,
        alignment: str | None,
    ) -> None:
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
            from win32com.client import constants

            text_range.ParagraphFormat.Alignment = self._resolve_office_constant(constants, alignment, TEXT_ALIGNMENTS)

    def _apply_shape_fill_and_line(
        self,
        shape: object,
        *,
        fill_color: str | None,
        fill_transparency: float | None,
        line_color: str | None,
        line_weight: float | None,
        line_transparency: float | None,
        line_visible: bool | None,
    ) -> None:
        if fill_color is not None:
            shape.Fill.Visible = -1
            shape.Fill.Solid()
            shape.Fill.ForeColor.RGB = parse_office_color(fill_color)
        if fill_transparency is not None:
            shape.Fill.Transparency = fill_transparency

        if line_visible is not None:
            shape.Line.Visible = -1 if line_visible else 0
        if line_color is not None:
            shape.Line.Visible = -1
            shape.Line.ForeColor.RGB = parse_office_color(line_color)
        if line_weight is not None:
            shape.Line.Weight = line_weight
        if line_transparency is not None:
            shape.Line.Transparency = line_transparency

    def _apply_two_color_gradient(
        self,
        fill: object,
        *,
        start_color: str,
        end_color: str,
        style: str,
        variant: int,
    ) -> None:
        from win32com.client import constants

        fill.Visible = -1
        with suppress(Exception):
            fill.ForeColor.RGB = parse_office_color(start_color)
        with suppress(Exception):
            fill.BackColor.RGB = parse_office_color(end_color)
        fill.TwoColorGradient(
            self._resolve_office_constant(constants, style, GRADIENT_STYLES),
            variant,
        )

    def _get_text_gradient_fill(self, shape: object) -> object:
        with suppress(Exception):
            return shape.TextFrame2.TextRange.Font.Fill
        raise ValueError("The selected shape does not support gradient text formatting")

    def _estimate_table_row_height(self, row_values: list[str], font_size: float = 14.0) -> float:
        max_lines = 1
        max_chars = 0
        for value in row_values:
            text = value or ""
            lines = max(1, text.count("\n") + text.count("\r") + 1)
            max_lines = max(max_lines, lines)
            max_chars = max(max_chars, len(text))
        wrapped_lines = max(1, (max_chars // 28) + 1)
        estimated_lines = max(max_lines, wrapped_lines)
        return max(18.0, estimated_lines * (font_size * 1.35))

    def _normalize_excel_range_values(self, values: object) -> list[list[str]]:
        if values is None:
            return []
        if not isinstance(values, tuple):
            return [["" if values is None else str(values)]]

        normalized: list[list[str]] = []
        for row in values:
            if isinstance(row, tuple):
                normalized.append(["" if cell is None else str(cell) for cell in row])
            else:
                normalized.append(["" if row is None else str(row)])
        return normalized

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

    def _resolve_smartart_collection_item(self, collection: object, value: str, label: str) -> object:
        normalized_value = normalize_powerpoint_token(value)
        if normalized_value.isdigit():
            return collection(int(normalized_value))

        for index in range(1, int(collection.Count) + 1):
            candidate = collection(index)
            with suppress(Exception):
                if normalize_powerpoint_token(str(candidate.Id)) == normalized_value:
                    return candidate
            with suppress(Exception):
                if normalize_powerpoint_token(str(candidate.Name)) == normalized_value:
                    return candidate
            with suppress(Exception):
                if normalize_powerpoint_token(str(candidate.Description)) == normalized_value:
                    return candidate

        raise ValueError(
            f"Unsupported SmartArt {label}: {value}. Use a collection index, Id, Name, or Description exposed by Office"
        )

    def _resolve_smartart_quickstyle(self, powerpoint: object, style: str) -> object:
        return self._resolve_smartart_collection_item(powerpoint.SmartArtQuickStyles, style, "style")

    def _resolve_smartart_color_theme(self, powerpoint: object, color_theme: str) -> object:
        return self._resolve_smartart_collection_item(powerpoint.SmartArtColors, color_theme, "color theme")

    def _collect_shape_range_names(self, shape_range: object) -> list[str]:
        shape_names: list[str] = []
        for index in range(1, int(shape_range.Count) + 1):
            with suppress(Exception):
                shape_names.append(str(shape_range(index).Name))
        return shape_names

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

    def _get_slide_section_info(self, presentation: object, slide_index: int) -> tuple[int | None, str | None]:
        with suppress(Exception):
            section_properties = presentation.SectionProperties
            for section_index in range(1, int(section_properties.Count) + 1):
                first_slide = int(section_properties.FirstSlide(section_index))
                slide_count = int(section_properties.SlidesCount(section_index))
                if first_slide <= slide_index < first_slide + max(slide_count, 0):
                    section_name = str(section_properties.Name(section_index)).strip() or None
                    return section_index, section_name
        return None, None

    def _extract_layout_summary(
        self,
        slide: object,
        *,
        design_index: int | None = None,
        master_name: str | None = None,
    ) -> LayoutSummary:
        layout_id = None
        layout_index = None
        layout_name = None
        with suppress(Exception):
            layout_id = int(slide.Layout)
        with suppress(Exception):
            layout = slide.CustomLayout
            layout_index = int(layout.Index)
            layout_name = str(layout.Name).strip() or None

        return LayoutSummary(
            layout_id=layout_id,
            layout_index=layout_index,
            layout_name=layout_name,
            design_index=design_index,
            master_name=master_name,
        )

    def _build_slide_metadata(self, presentation: object, slide: object) -> SlideMetadataSummary:
        slide_index = int(slide.SlideIndex)
        slide_id = None
        slide_name = None
        hidden = False
        with suppress(Exception):
            slide_id = int(slide.SlideID)
        with suppress(Exception):
            slide_name = str(slide.Name).strip() or None
        with suppress(Exception):
            hidden = bool(slide.SlideShowTransition.Hidden)

        section_index, section_name = self._get_slide_section_info(presentation, slide_index)
        layout = self._extract_layout_summary(slide)
        return SlideMetadataSummary(
            slide_index=slide_index,
            slide_id=slide_id,
            name=slide_name,
            title=self._get_slide_title(slide),
            shape_count=int(slide.Shapes.Count),
            hidden=hidden,
            layout_id=layout.layout_id,
            layout_index=layout.layout_index,
            layout_name=layout.layout_name,
            section_index=section_index,
            section_name=section_name,
        )

    def _iter_presentation_layouts(self, presentation: object) -> Iterator[LayoutSummary]:
        with suppress(Exception):
            designs = presentation.Designs
            for design_index in range(1, int(designs.Count) + 1):
                design = designs(design_index)
                master_name = None
                with suppress(Exception):
                    master_name = str(design.SlideMaster.Name).strip() or None

                custom_layouts = design.SlideMaster.CustomLayouts
                for layout_index in range(1, int(custom_layouts.Count) + 1):
                    layout = custom_layouts(layout_index)
                    layout_name = None
                    with suppress(Exception):
                        layout_name = str(layout.Name).strip() or None
                    yield LayoutSummary(
                        layout_index=int(layout.Index),
                        layout_name=layout_name,
                        design_index=design_index,
                        master_name=master_name,
                    )

    def _resolve_slide_layout_target(self, presentation: object, layout: str) -> tuple[str, object | int]:
        normalized = normalize_powerpoint_token(layout)
        for design_index in range(1, int(presentation.Designs.Count) + 1):
            design = presentation.Designs(design_index)
            custom_layouts = design.SlideMaster.CustomLayouts
            for layout_index in range(1, int(custom_layouts.Count) + 1):
                custom_layout = custom_layouts(layout_index)
                candidate_name = None
                candidate_index = None
                with suppress(Exception):
                    candidate_name = str(custom_layout.Name).strip() or None
                with suppress(Exception):
                    candidate_index = int(custom_layout.Index)

                if candidate_index is not None and normalized.isdigit() and candidate_index == int(normalized):
                    return "custom", custom_layout
                if candidate_name and normalize_powerpoint_token(candidate_name) == normalized:
                    return "custom", custom_layout

        return "built_in", resolve_slide_layout(layout)

    def _get_design(self, presentation: object, master_index: int) -> object:
        designs = presentation.Designs
        design_count = int(designs.Count)
        if not 1 <= master_index <= design_count:
            raise ValueError(f"The presentation does not contain master_index={master_index}")
        return designs(master_index)

    def _get_design_theme_name(self, design: object) -> str | None:
        with suppress(Exception):
            return str(design.Name).strip() or None
        with suppress(Exception):
            return str(design.SlideMaster.Theme.Name).strip() or None
        return None

    def _extract_master_summary(self, design_index: int, design: object) -> MasterSummary:
        master_name = None
        layout_count = 0
        with suppress(Exception):
            master_name = str(design.SlideMaster.Name).strip() or None
        with suppress(Exception):
            layout_count = int(design.SlideMaster.CustomLayouts.Count)
        return MasterSummary(
            master_index=design_index,
            master_name=master_name,
            theme_name=self._get_design_theme_name(design),
            layout_count=layout_count,
        )

    def _iter_master_layouts(self, design_index: int, design: object) -> Iterator[LayoutSummary]:
        master_name = None
        with suppress(Exception):
            master_name = str(design.SlideMaster.Name).strip() or None

        custom_layouts = design.SlideMaster.CustomLayouts
        for layout_index in range(1, int(custom_layouts.Count) + 1):
            layout = custom_layouts(layout_index)
            layout_name = None
            with suppress(Exception):
                layout_name = str(layout.Name).strip() or None
            yield LayoutSummary(
                layout_index=int(layout.Index),
                layout_name=layout_name,
                design_index=design_index,
                master_name=master_name,
            )

    def _get_design_variants(self, design: object) -> object | None:
        with suppress(Exception):
            return design.ThemeVariants
        with suppress(Exception):
            return design.Variants
        with suppress(Exception):
            return design.SlideMaster.Theme.ThemeVariants
        return None

    def _iter_theme_variants(self, design: object) -> Iterator[ThemeVariantSummary]:
        variants = self._get_design_variants(design)
        if variants is None:
            return

        for variant_index in range(1, int(variants.Count) + 1):
            variant = variants(variant_index)
            name = None
            color_scheme_name = None
            font_scheme_name = None
            with suppress(Exception):
                name = str(variant.Name).strip() or None
            with suppress(Exception):
                color_scheme_name = str(variant.ColorScheme.Name).strip() or None
            with suppress(Exception):
                font_scheme_name = str(variant.FontScheme.Name).strip() or None
            yield ThemeVariantSummary(
                variant_index=variant_index,
                name=name,
                color_scheme_name=color_scheme_name,
                font_scheme_name=font_scheme_name,
            )

    def _extract_master_background_color(self, master: object) -> str | None:
        with suppress(Exception):
            return office_color_to_hex(int(master.Background.Fill.ForeColor.RGB))
        return None

    def _extract_theme_fonts(self, design: object) -> dict[str, object]:
        fonts: dict[str, object] = {
            "title_font_name": None,
            "title_font_size": None,
            "body_font_name": None,
            "body_font_size": None,
        }

        with suppress(Exception):
            from win32com.client import constants

            font_scheme = design.SlideMaster.Theme.ThemeFontScheme
            for label, constant_name in THEME_FONT_LANGUAGE_CONSTANTS.items():
                constant_value = getattr(constants, constant_name)
                with suppress(Exception):
                    fonts[f"major_{label}"] = str(font_scheme.MajorFont(constant_value).Name).strip() or None
                with suppress(Exception):
                    fonts[f"minor_{label}"] = str(font_scheme.MinorFont(constant_value).Name).strip() or None

        containers = [design.SlideMaster]
        custom_layouts = design.SlideMaster.CustomLayouts
        for layout_index in range(1, int(custom_layouts.Count) + 1):
            containers.append(custom_layouts(layout_index))

        for container in containers:
            for _, shape in self._iter_shapes(container):
                if not self._shape_supports_text(shape):
                    continue
                with suppress(Exception):
                    font = self._get_text_range(shape).Font
                    font_name = str(font.Name).strip() or None
                    font_size = float(font.Size) if getattr(font, "Size", None) else None
                    if self._shape_is_title_placeholder(shape):
                        if fonts["title_font_name"] is None and font_name is not None:
                            fonts["title_font_name"] = font_name
                        if fonts["title_font_size"] is None and font_size is not None:
                            fonts["title_font_size"] = font_size
                    else:
                        if fonts["body_font_name"] is None and font_name is not None:
                            fonts["body_font_name"] = font_name
                        if fonts["body_font_size"] is None and font_size is not None:
                            fonts["body_font_size"] = font_size
                if fonts["title_font_name"] is not None and fonts["body_font_name"] is not None:
                    return fonts

        return fonts

    def _extract_theme_colors(self, design: object, background_color: str | None) -> dict[str, object]:
        colors: dict[str, object] = {
            "background_color": background_color,
            "title_text_color": None,
            "body_text_color": None,
            "accent_color": None,
        }

        with suppress(Exception):
            from win32com.client import constants

            scheme = design.SlideMaster.Theme.ThemeColorScheme
            for key, constant_name in THEME_COLOR_CONSTANTS.items():
                with suppress(Exception):
                    colors[key] = office_color_to_hex(int(scheme(getattr(constants, constant_name)).RGB))

        containers = [design.SlideMaster]
        custom_layouts = design.SlideMaster.CustomLayouts
        for layout_index in range(1, int(custom_layouts.Count) + 1):
            containers.append(custom_layouts(layout_index))

        for container in containers:
            for _, shape in self._iter_shapes(container):
                if self._shape_supports_text(shape):
                    with suppress(Exception):
                        font = self._get_text_range(shape).Font
                        font_color = office_color_to_hex(int(font.Color.RGB))
                        if self._shape_is_title_placeholder(shape):
                            if colors["title_text_color"] is None and font_color is not None:
                                colors["title_text_color"] = font_color
                        elif colors["body_text_color"] is None and font_color is not None:
                            colors["body_text_color"] = font_color
                if colors["accent_color"] is None:
                    with suppress(Exception):
                        colors["accent_color"] = office_color_to_hex(int(shape.Fill.ForeColor.RGB))

        return colors

    def _build_master_theme_summary(self, design_index: int, design: object) -> MasterThemeSummary:
        master = design.SlideMaster
        background_color = self._extract_master_background_color(master)
        placeholders = [
            summary
            for index, shape in self._iter_shapes(master)
            if (summary := self._extract_placeholder_summary(shape, index)) is not None
        ]
        return MasterThemeSummary(
            master_index=design_index,
            master_name=self._extract_master_summary(design_index, design).master_name,
            theme_name=self._get_design_theme_name(design),
            background_color=background_color,
            fonts=self._extract_theme_fonts(design),
            colors=self._extract_theme_colors(design, background_color),
            layouts=list(self._iter_master_layouts(design_index, design)),
            placeholders=placeholders,
            variants=list(self._iter_theme_variants(design)),
        )

    def _candidate_builtin_theme_dirs(self) -> list[Path]:
        candidates: list[Path] = []
        for env_name in ("ProgramFiles", "ProgramFiles(x86)"):
            root = Path(__import__("os").environ.get(env_name, "")).expanduser()
            if not str(root):
                continue
            candidates.extend(
                [
                    root / "Microsoft Office" / "root" / "Document Themes 16",
                    root / "Microsoft Office" / "Document Themes 16",
                    root / "Microsoft Office" / "root" / "Document Themes",
                ]
            )

        for env_name in ("APPDATA", "LOCALAPPDATA"):
            root = Path(__import__("os").environ.get(env_name, "")).expanduser()
            if not str(root):
                continue
            candidates.append(root / "Microsoft" / "Templates" / "Document Themes")

        unique_candidates: list[Path] = []
        seen: set[str] = set()
        for candidate in candidates:
            resolved = str(candidate)
            if resolved in seen:
                continue
            seen.add(resolved)
            unique_candidates.append(candidate)
        return unique_candidates

    def _iter_builtin_theme_paths(self) -> Iterator[Path]:
        seen: set[Path] = set()
        for directory in self._candidate_builtin_theme_dirs():
            if not directory.exists():
                continue
            with suppress(Exception):
                for theme_path in directory.rglob("*.thmx"):
                    resolved = theme_path.resolve()
                    if resolved in seen:
                        continue
                    seen.add(resolved)
                    yield resolved

    def _resolve_builtin_theme_path(self, theme_name: str) -> Path:
        requested = normalize_powerpoint_token(Path(theme_name).stem)
        exact_matches: list[Path] = []
        partial_matches: list[Path] = []

        for theme_path in self._iter_builtin_theme_paths():
            stem = normalize_powerpoint_token(theme_path.stem)
            if stem == requested:
                exact_matches.append(theme_path)
            elif requested in stem:
                partial_matches.append(theme_path)

        matches = exact_matches or partial_matches
        if not matches:
            raise ValueError(f"Built-in PowerPoint theme not found: {theme_name}")
        if len(matches) > 1:
            sample = ", ".join(sorted(path.stem for path in matches[:6]))
            raise ValueError(f"Theme name '{theme_name}' is ambiguous. Matches: {sample}")
        return matches[0]

    def _resolve_placeholder_shape(
        self,
        slide: object,
        *,
        shape_index: int | None,
        shape_name: str | None,
        placeholder_type: int | None,
        placeholder_occurrence: int,
    ) -> tuple[int, object]:
        if shape_index is not None:
            shape = self._get_shape(slide, shape_index)
            if self._get_placeholder_type(shape) is None:
                raise ValueError(f"shape_index={shape_index} is not a placeholder")
            return shape_index, shape

        if shape_name is not None:
            target = normalize_powerpoint_token(shape_name)
            for index, shape in self._iter_shapes(slide):
                with suppress(Exception):
                    if normalize_powerpoint_token(str(shape.Name)) == target and self._get_placeholder_type(shape) is not None:
                        return index, shape
            raise ValueError(f"No placeholder found with shape_name='{shape_name}'")

        matches: list[tuple[int, object]] = []
        for index, shape in self._iter_shapes(slide):
            if self._get_placeholder_type(shape) == placeholder_type:
                matches.append((index, shape))

        if len(matches) < placeholder_occurrence:
            raise ValueError(
                f"No placeholder found for placeholder_type={placeholder_type} occurrence={placeholder_occurrence}"
            )
        return matches[placeholder_occurrence - 1]

    def _set_theme_scheme_color(self, design: object, constant_name: str, color: str) -> None:
        with suppress(Exception):
            from win32com.client import constants

            scheme = design.SlideMaster.Theme.ThemeColorScheme
            scheme(getattr(constants, constant_name)).RGB = parse_office_color(color)

    def _iter_master_text_shapes(self, design: object) -> Iterator[object]:
        containers = [design.SlideMaster]
        custom_layouts = design.SlideMaster.CustomLayouts
        for layout_index in range(1, int(custom_layouts.Count) + 1):
            containers.append(custom_layouts(layout_index))

        for container in containers:
            for _, shape in self._iter_shapes(container):
                if self._shape_supports_text(shape):
                    yield shape

    def _apply_style_preset_to_slide(self, slide: object, preset: str) -> tuple[str | None, str | None]:
        style = resolve_style_preset(preset)
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

        return title_shape_name, body_shape_name

    def _replace_text_in_shapes(self, slide: object, find_text: str, replace_text: str) -> tuple[int, list[int]]:
        replacement_count = 0
        updated_shapes: list[int] = []
        for index in range(1, int(slide.Shapes.Count) + 1):
            shape = slide.Shapes(index)
            with suppress(Exception):
                if shape.HasTextFrame and shape.TextFrame.HasText:
                    current_text = str(shape.TextFrame.TextRange.Text)
                    occurrences = current_text.count(find_text)
                    if occurrences:
                        shape.TextFrame.TextRange.Text = current_text.replace(find_text, replace_text)
                        replacement_count += occurrences
                        updated_shapes.append(index)
        return replacement_count, updated_shapes

    def _collect_text_matches(
        self,
        *,
        slide_index: int,
        query: str,
        text: str,
        location: str,
        shape_index: int | None = None,
        shape_name: str | None = None,
    ) -> TextMatchSummary | None:
        occurrences = text.count(query)
        if occurrences <= 0:
            return None

        preview = " ".join(text.split())[:160] or None
        return TextMatchSummary(
            slide_index=slide_index,
            location=location,
            occurrences=occurrences,
            shape_index=shape_index,
            shape_name=shape_name,
            text_preview=preview,
        )

    def _get_document_property(self, properties: object, property_name: str) -> object | None:
        with suppress(Exception):
            return normalize_office_value(properties(property_name).Value)
        with suppress(Exception):
            return normalize_office_value(properties.Item(property_name).Value)
        return None

    def _set_document_property(self, properties: object, property_name: str, value: object) -> None:
        with suppress(Exception):
            properties(property_name).Value = value
            return
        properties.Item(property_name).Value = value

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def get_document_properties(self, path: str) -> DocumentPropertiesResult:
        source = self.resolve_document_path(path)

        with self._open_presentation(source, read_only=True) as presentation:
            built_in = {
                key: self._get_document_property(presentation.BuiltInDocumentProperties, property_name)
                for key, property_name in BUILTIN_DOCUMENT_PROPERTY_NAMES.items()
            }
            custom: dict[str, object] = {}
            with suppress(Exception):
                properties = presentation.CustomDocumentProperties
                for index in range(1, int(properties.Count) + 1):
                    item = properties(index)
                    custom[str(item.Name)] = normalize_office_value(item.Value)

            return DocumentPropertiesResult(file_path=str(source), built_in=built_in, custom=custom)

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_document_properties(
        self,
        path: str,
        *,
        author: str | None,
        title: str | None,
        subject: str | None,
        keywords: str | None,
        comments: str | None,
        category: str | None,
        company: str | None,
        manager: str | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)
        updates = {
            "author": author,
            "title": title,
            "subject": subject,
            "keywords": keywords,
            "comments": comments,
            "category": category,
            "company": company,
            "manager": manager,
        }

        with self._open_presentation(source, read_only=False) as presentation:
            properties = presentation.BuiltInDocumentProperties
            applied_updates: dict[str, object] = {}
            for key, value in updates.items():
                if value is None:
                    continue
                property_name = BUILTIN_DOCUMENT_PROPERTY_NAMES[key]
                self._set_document_property(properties, property_name, value)
                applied_updates[key] = value

            presentation.Save()
            return OperationResult(
                message="PowerPoint document properties updated",
                file_path=str(source),
                backup_path=backup_path,
                details={"updated_properties": applied_updates},
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def get_file_links(self, path: str) -> FileLinksResult:
        source = self.resolve_document_path(path)

        with self._open_presentation(source, read_only=True) as presentation:
            raw_links = None
            with suppress(Exception):
                raw_links = presentation.LinkSources()

            normalized = normalize_office_value(raw_links)
            if normalized is None:
                links: list[object] = []
            elif isinstance(normalized, list):
                links = normalized
            else:
                links = [normalized]

            return FileLinksResult(file_path=str(source), links=links)

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def save(self, path: str, create_backup: bool) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            presentation.Save()

        return OperationResult(
            message="PowerPoint presentation saved",
            file_path=str(source),
            backup_path=backup_path,
        )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def save_copy(self, path: str, out_path: str) -> OperationResult:
        source = self.resolve_document_path(path)
        target = self.resolve_output_path(out_path, allowed_suffixes=self.allowed_suffixes)

        with self._open_presentation(source, read_only=False) as presentation:
            try:
                presentation.SaveCopyAs(str(target))
            except Exception:
                presentation.SaveAs(str(target))

        return OperationResult(
            message="PowerPoint presentation copy saved",
            file_path=str(source),
            details={"out_path": str(target)},
        )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def create_presentation(
        self,
        path: str,
        layout: str,
        title: str | None,
        body_text: str | None,
    ) -> OperationResult:
        target = self.resolve_output_path(path, allowed_suffixes=self.allowed_suffixes)
        if target.exists():
            raise ValueError(f"Target PowerPoint file already exists: {target}")

        with office_application("PowerPoint.Application", visible=self.settings.office_visible) as powerpoint:
            presentation = None
            try:
                presentation = powerpoint.Presentations.Add(WithWindow=self.settings.office_visible)
                slide = None

                if int(presentation.Slides.Count) < 1:
                    slide = presentation.Slides.Add(1, resolve_slide_layout("blank"))
                else:
                    slide = presentation.Slides(1)

                layout_kind, target_layout = self._resolve_slide_layout_target(presentation, layout)
                if layout_kind == "custom":
                    slide.CustomLayout = target_layout
                else:
                    slide.Layout = int(target_layout)

                if title:
                    self._set_title_text(slide, title)
                if body_text:
                    self._set_body_text(slide, body_text)

                presentation.SaveAs(str(target))

                slide_id = None
                with suppress(Exception):
                    slide_id = int(slide.SlideID)

                return OperationResult(
                    message="PowerPoint presentation created",
                    file_path=str(target),
                    details={
                        "slide_count": int(presentation.Slides.Count),
                        "layout": layout,
                        "layout_kind": layout_kind,
                        "slide_index": int(slide.SlideIndex),
                        "slide_id": slide_id,
                    },
                )
            finally:
                if presentation is not None:
                    with suppress(Exception):
                        presentation.Close()

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def duplicate_slide(self, path: str, slide_index: int, create_backup: bool) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            duplicated_range = slide.Duplicate()
            duplicated_slide = None
            with suppress(Exception):
                duplicated_slide = duplicated_range(1)
            if duplicated_slide is None:
                duplicated_slide = presentation.Slides(slide_index + 1)

            presentation.Save()
            return OperationResult(
                message="PowerPoint slide duplicated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "source_slide_index": slide_index,
                    "slide_index": int(duplicated_slide.SlideIndex),
                    "slide_id": int(duplicated_slide.SlideID),
                    "title": self._get_slide_title(duplicated_slide),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def delete_slide(self, path: str, slide_index: int, create_backup: bool) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            slide_id = None
            slide_name = None
            with suppress(Exception):
                slide_id = int(slide.SlideID)
            with suppress(Exception):
                slide_name = str(slide.Name).strip() or None
            title = self._get_slide_title(slide)
            slide.Delete()
            presentation.Save()

            return OperationResult(
                message="PowerPoint slide deleted",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "slide_id": slide_id,
                    "name": slide_name,
                    "title": title,
                    "remaining_slide_count": int(presentation.Slides.Count),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def move_slide(self, path: str, slide_index: int, position: int, create_backup: bool) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide_count = int(presentation.Slides.Count)
            if not 1 <= position <= slide_count:
                raise ValueError(f"position must be between 1 and {slide_count}")

            slide = presentation.Slides(slide_index)
            slide.MoveTo(position)
            presentation.Save()

            return OperationResult(
                message="PowerPoint slide moved",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "previous_slide_index": slide_index,
                    "slide_index": int(slide.SlideIndex),
                    "position": position,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def hide_slide(self, path: str, slide_index: int, create_backup: bool) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            slide.SlideShowTransition.Hidden = -1
            presentation.Save()

            return OperationResult(
                message="PowerPoint slide hidden",
                file_path=str(source),
                backup_path=backup_path,
                details={"slide_index": slide_index, "hidden": True},
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def unhide_slide(self, path: str, slide_index: int, create_backup: bool) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            slide.SlideShowTransition.Hidden = 0
            presentation.Save()

            return OperationResult(
                message="PowerPoint slide unhidden",
                file_path=str(source),
                backup_path=backup_path,
                details={"slide_index": slide_index, "hidden": False},
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_slide_name(self, path: str, slide_index: int, name: str, create_backup: bool) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            slide.Name = name
            presentation.Save()

            return OperationResult(
                message="PowerPoint slide name updated",
                file_path=str(source),
                backup_path=backup_path,
                details={"slide_index": slide_index, "name": name},
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def get_slide_metadata(self, path: str, slide_index: int) -> SlideMetadataResult:
        source = self.resolve_document_path(path)

        with self._open_presentation(source, read_only=True) as presentation:
            slide = presentation.Slides(slide_index)
            return SlideMetadataResult(file_path=str(source), slide=self._build_slide_metadata(presentation, slide))

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def get_slide_summary_extended(self, path: str, slide_index: int) -> ExtendedSlideSummaryResult:
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

            shapes = [self._shape_summary(shape, index) for index, shape in self._iter_shapes(slide)]
            tables = [self._table_summary(shape, index) for index, shape in self._iter_shapes(slide) if self._shape_has_table(shape)]
            charts = [self._chart_summary(shape, index) for index, shape in self._iter_shapes(slide) if self._shape_has_chart(shape)]
            smartart_items = [
                self._smartart_summary(shape, index) for index, shape in self._iter_shapes(slide) if self._shape_has_smartart(shape)
            ]

            return ExtendedSlideSummaryResult(
                file_path=str(source),
                slide_index=slide_index,
                metadata=self._build_slide_metadata(presentation, slide),
                texts=self._get_slide_texts(slide),
                notes_text=self._get_slide_notes_text(slide),
                shapes=shapes,
                tables=tables,
                charts=charts,
                smartart_items=smartart_items,
                animations=animations,
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def list_masters(self, path: str) -> PresentationMastersResult:
        source = self.resolve_document_path(path)

        with self._open_presentation(source, read_only=True) as presentation:
            masters = [
                self._extract_master_summary(design_index, presentation.Designs(design_index))
                for design_index in range(1, int(presentation.Designs.Count) + 1)
            ]
            return PresentationMastersResult(file_path=str(source), masters=masters)

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def get_master_details(self, path: str, master_index: int) -> MasterDetailsResult:
        source = self.resolve_document_path(path)

        with self._open_presentation(source, read_only=True) as presentation:
            design = self._get_design(presentation, master_index)
            return MasterDetailsResult(
                file_path=str(source),
                master=self._build_master_theme_summary(master_index, design),
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def list_layouts(self, path: str) -> PresentationLayoutsResult:
        source = self.resolve_document_path(path)

        with self._open_presentation(source, read_only=True) as presentation:
            return PresentationLayoutsResult(
                file_path=str(source),
                layouts=list(self._iter_presentation_layouts(presentation)),
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def get_slide_layout(self, path: str, slide_index: int) -> SlideLayoutResult:
        source = self.resolve_document_path(path)

        with self._open_presentation(source, read_only=True) as presentation:
            slide = presentation.Slides(slide_index)
            return SlideLayoutResult(
                file_path=str(source),
                slide_index=slide_index,
                layout=self._extract_layout_summary(slide),
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def apply_layout(self, path: str, slide_index: int, layout: str, create_backup: bool) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            target_kind, target_layout = self._resolve_slide_layout_target(presentation, layout)
            if target_kind == "custom":
                slide.CustomLayout = target_layout
            else:
                slide.Layout = int(target_layout)
            presentation.Save()
            applied_layout = self._extract_layout_summary(slide)

            return OperationResult(
                message="PowerPoint slide layout applied",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "layout": layout,
                    "layout_id": applied_layout.layout_id,
                    "layout_index": applied_layout.layout_index,
                    "layout_name": applied_layout.layout_name,
                    "layout_kind": target_kind,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def reset_slide_to_layout(self, path: str, slide_index: int, create_backup: bool) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            slide.Reset()
            presentation.Save()
            layout = self._extract_layout_summary(slide)

            return OperationResult(
                message="PowerPoint slide reset to layout",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "layout_id": layout.layout_id,
                    "layout_index": layout.layout_index,
                    "layout_name": layout.layout_name,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def list_placeholders(self, path: str, slide_index: int) -> SlidePlaceholdersResult:
        source = self.resolve_document_path(path)

        with self._open_presentation(source, read_only=True) as presentation:
            slide = presentation.Slides(slide_index)
            placeholders = [
                summary
                for index, shape in self._iter_shapes(slide)
                if (summary := self._extract_placeholder_summary(shape, index)) is not None
            ]

            return SlidePlaceholdersResult(file_path=str(source), slide_index=slide_index, placeholders=placeholders)

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def fill_placeholder(
        self,
        path: str,
        slide_index: int,
        shape_index: int | None,
        shape_name: str | None,
        placeholder_type: int | None,
        placeholder_occurrence: int,
        text: str,
        text_color: str | None,
        font_name: str | None,
        font_size: float | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            resolved_index, shape = self._resolve_placeholder_shape(
                slide,
                shape_index=shape_index,
                shape_name=shape_name,
                placeholder_type=placeholder_type,
                placeholder_occurrence=placeholder_occurrence,
            )
            self._apply_shape_text_format(shape, text, text_color, font_name, font_size)
            presentation.Save()
            return OperationResult(
                message="PowerPoint placeholder filled",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": resolved_index,
                    "shape_name": str(shape.Name),
                    "placeholder_type": self._get_placeholder_type(shape),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def replace_placeholder_with_shape(
        self,
        path: str,
        slide_index: int,
        shape_index: int | None,
        shape_name: str | None,
        placeholder_type: int | None,
        placeholder_occurrence: int,
        replacement_kind: str,
        text: str | None,
        image_path: str | None,
        shape_type: str,
        fill_color: str | None,
        line_color: str | None,
        text_color: str | None,
        font_name: str | None,
        font_size: float | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)
        normalized_kind = normalize_powerpoint_token(replacement_kind)

        image = None
        if normalized_kind == "image":
            image = validate_file_path(
                image_path or "",
                allowed_roots=self.settings.allowed_roots,
                allowed_suffixes=IMAGE_SUFFIXES,
                must_exist=True,
            )

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            resolved_index, placeholder = self._resolve_placeholder_shape(
                slide,
                shape_index=shape_index,
                shape_name=shape_name,
                placeholder_type=placeholder_type,
                placeholder_occurrence=placeholder_occurrence,
            )
            previous_name = str(placeholder.Name)
            previous_text = None
            with suppress(Exception):
                previous_text = str(self._get_text_range(placeholder).Text).strip() or None

            left = float(placeholder.Left)
            top = float(placeholder.Top)
            width = float(placeholder.Width)
            height = float(placeholder.Height)
            placeholder.Delete()

            if normalized_kind in {"textbox", "text"}:
                new_shape = slide.Shapes.AddTextbox(1, left, top, width, height)
                replacement_text = text if text is not None else previous_text or ""
                self._apply_shape_text_format(new_shape, replacement_text, text_color, font_name, font_size)
                self._apply_shape_fill_and_line(
                    new_shape,
                    fill_color=fill_color,
                    fill_transparency=None,
                    line_color=line_color,
                    line_weight=None,
                    line_transparency=None,
                    line_visible=None,
                )
            elif normalized_kind == "image":
                new_shape = slide.Shapes.AddPicture(str(image), False, True, left, top, width, height)
            else:
                new_shape = slide.Shapes.AddShape(resolve_shape_type(shape_type), left, top, width, height)
                self._apply_shape_fill_and_line(
                    new_shape,
                    fill_color=fill_color,
                    fill_transparency=None,
                    line_color=line_color,
                    line_weight=None,
                    line_transparency=None,
                    line_visible=None,
                )
                replacement_text = text if text is not None else previous_text
                if replacement_text is not None and self._shape_supports_text(new_shape):
                    self._apply_shape_text_format(new_shape, replacement_text, text_color, font_name, font_size)

            with suppress(Exception):
                new_shape.Name = f"{previous_name}_content"

            presentation.Save()
            return OperationResult(
                message="PowerPoint placeholder replaced with shape",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "placeholder_shape_index": resolved_index,
                    "previous_shape_name": previous_name,
                    "replacement_kind": normalized_kind,
                    "shape_name": str(new_shape.Name),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def restore_placeholder(self, path: str, slide_index: int, create_backup: bool) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            before = [
                summary
                for index, shape in self._iter_shapes(slide)
                if (summary := self._extract_placeholder_summary(shape, index)) is not None
            ]
            slide.Reset()
            after = [
                summary
                for index, shape in self._iter_shapes(slide)
                if (summary := self._extract_placeholder_summary(shape, index)) is not None
            ]
            presentation.Save()
            return OperationResult(
                message="PowerPoint slide placeholders restored",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "placeholder_count_before": len(before),
                    "placeholder_count_after": len(after),
                    "restored_placeholder_count": max(len(after) - len(before), 0),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def find_text(self, path: str, query: str) -> PresentationTextSearchResult:
        source = self.resolve_document_path(path)

        with self._open_presentation(source, read_only=True) as presentation:
            matches: list[TextMatchSummary] = []
            for slide_index in range(1, int(presentation.Slides.Count) + 1):
                slide = presentation.Slides(slide_index)
                for shape_index, shape in self._iter_shapes(slide):
                    with suppress(Exception):
                        if shape.HasTextFrame and shape.TextFrame.HasText:
                            text = str(shape.TextFrame.TextRange.Text)
                            match = self._collect_text_matches(
                                slide_index=slide_index,
                                query=query,
                                text=text,
                                location="slide",
                                shape_index=shape_index,
                                shape_name=str(shape.Name),
                            )
                            if match is not None:
                                matches.append(match)

            return PresentationTextSearchResult(
                file_path=str(source),
                query=query,
                match_count=sum(match.occurrences for match in matches),
                matches=matches,
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def replace_text_all(self, path: str, find_text: str, replace_text: str, create_backup: bool) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            replacement_count = 0
            updated_slides: list[dict[str, object]] = []

            for slide_index in range(1, int(presentation.Slides.Count) + 1):
                slide = presentation.Slides(slide_index)
                slide_replacements, updated_shapes = self._replace_text_in_shapes(slide, find_text, replace_text)
                if slide_replacements:
                    replacement_count += slide_replacements
                    updated_slides.append(
                        {
                            "slide_index": slide_index,
                            "replacement_count": slide_replacements,
                            "shape_indexes": updated_shapes,
                        }
                    )

            presentation.Save()
            return OperationResult(
                message="PowerPoint presentation text replacement completed",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "find_text": find_text,
                    "replace_text": replace_text,
                    "replacement_count": replacement_count,
                    "updated_slide_count": len(updated_slides),
                    "updated_slides": updated_slides,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_slide_title(self, path: str, slide_index: int, title: str, create_backup: bool) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            self._set_title_text(slide, title)
            presentation.Save()
            return OperationResult(
                message="PowerPoint slide title updated",
                file_path=str(source),
                backup_path=backup_path,
                details={"slide_index": slide_index, "title": title},
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def get_shape_text_runs(self, path: str, slide_index: int, shape_index: int) -> ShapeTextRunsResult:
        source = self.resolve_document_path(path)

        with self._open_presentation(source, read_only=True) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            text_range = self._get_text_range(shape)
            runs = text_range.Runs()

            run_summaries: list[TextRunSummary] = []
            cursor = 1
            for run_index in range(1, int(runs.Count) + 1):
                run = text_range.Runs(run_index, 1)
                summary = self._summarize_text_run(run, run_index, cursor)
                run_summaries.append(summary)
                cursor += summary.length

            return ShapeTextRunsResult(
                file_path=str(source),
                slide_index=slide_index,
                shape_index=shape_index,
                run_count=len(run_summaries),
                runs=run_summaries,
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_text_range_style(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        start: int,
        length: int,
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

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            target_range = self._get_text_characters(self._get_text_range(shape), start, length)
            self._apply_text_range_style(
                target_range,
                text=text,
                font_name=font_name,
                font_size=font_size,
                bold=bold,
                italic=italic,
                underline=underline,
                color=color,
                alignment=alignment,
            )

            presentation.Save()
            return OperationResult(
                message="PowerPoint text range style updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "start": start,
                    "length": length,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def insert_bullets(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        items: list[str],
        level: int,
        append: bool,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            text_range = self._get_text_range(shape)
            existing_text = str(text_range.Text) if append else ""
            bullet_text = "\r\n".join(items)
            combined_text = bullet_text if not existing_text else f"{existing_text}\r\n{bullet_text}"
            text_range.Text = combined_text

            paragraph_count = int(text_range.Paragraphs().Count)
            start_index = paragraph_count - len(items) + 1
            for paragraph_index in range(start_index, paragraph_count + 1):
                paragraph_range = text_range.Paragraphs(paragraph_index, 1)
                paragraph_range.ParagraphFormat.Bullet.Visible = -1
                paragraph_range.IndentLevel = level

            presentation.Save()
            return OperationResult(
                message="PowerPoint bullets inserted",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "item_count": len(items),
                    "level": level,
                    "append": append,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_bullet_style(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        paragraph_index: int | None,
        visible: bool | None,
        level: int | None,
        bullet_character: str | None,
        font_name: str | None,
        color: str | None,
        relative_size: float | None,
        left_margin: float | None,
        first_margin: float | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            paragraph_range = self._get_paragraph_range(self._get_text_range(shape), paragraph_index)
            paragraph_format = paragraph_range.ParagraphFormat
            bullet = paragraph_format.Bullet

            if visible is not None:
                bullet.Visible = -1 if visible else 0
            if level is not None:
                paragraph_range.IndentLevel = level
            if bullet_character is not None:
                bullet.Visible = -1
                bullet.Character = ord(bullet_character)
            if font_name is not None:
                bullet.Visible = -1
                bullet.UseTextFont = 0
                bullet.Font.Name = font_name
            if color is not None:
                bullet.Visible = -1
                bullet.UseTextColor = 0
                bullet.Font.Color.RGB = parse_office_color(color)
            if relative_size is not None:
                bullet.RelativeSize = relative_size
            if left_margin is not None:
                paragraph_format.LeftMargin = left_margin
            if first_margin is not None:
                paragraph_format.FirstMargin = first_margin

            presentation.Save()
            return OperationResult(
                message="PowerPoint bullet style updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "paragraph_index": paragraph_index,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_paragraph_spacing(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        paragraph_index: int | None,
        space_before: float | None,
        space_after: float | None,
        space_within: float | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            paragraph_range = self._get_paragraph_range(self._get_text_range(shape), paragraph_index)
            paragraph_format = paragraph_range.ParagraphFormat

            if space_before is not None:
                paragraph_format.SpaceBefore = space_before
            if space_after is not None:
                paragraph_format.SpaceAfter = space_after
            if space_within is not None:
                paragraph_format.SpaceWithin = space_within

            presentation.Save()
            return OperationResult(
                message="PowerPoint paragraph spacing updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "paragraph_index": paragraph_index,
                    "space_before": space_before,
                    "space_after": space_after,
                    "space_within": space_within,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_textbox_margins(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        margin_left: float | None,
        margin_right: float | None,
        margin_top: float | None,
        margin_bottom: float | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            text_frame = self._get_text_frame(shape)

            if margin_left is not None:
                text_frame.MarginLeft = margin_left
            if margin_right is not None:
                text_frame.MarginRight = margin_right
            if margin_top is not None:
                text_frame.MarginTop = margin_top
            if margin_bottom is not None:
                text_frame.MarginBottom = margin_bottom

            presentation.Save()
            return OperationResult(
                message="PowerPoint textbox margins updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "margin_left": margin_left,
                    "margin_right": margin_right,
                    "margin_top": margin_top,
                    "margin_bottom": margin_bottom,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_text_direction(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        direction: str,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        from win32com.client import constants

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            text_frame = self._get_text_frame(shape)
            text_frame.Orientation = self._resolve_office_constant(constants, direction, TEXT_DIRECTIONS)

            presentation.Save()
            return OperationResult(
                message="PowerPoint text direction updated",
                file_path=str(source),
                backup_path=backup_path,
                details={"slide_index": slide_index, "shape_index": shape_index, "direction": direction},
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_autofit(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        mode: str,
        word_wrap: bool | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        from win32com.client import constants

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            text_frame = self._get_text_frame(shape)
            text_frame.AutoSize = self._resolve_office_constant(constants, mode, TEXT_AUTOFIT_MODES)
            if word_wrap is not None:
                text_frame.WordWrap = -1 if word_wrap else 0

            presentation.Save()
            return OperationResult(
                message="PowerPoint text autofit updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "mode": mode,
                    "word_wrap": word_wrap,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_proofing_language(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        language: str,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        from win32com.client import constants

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            text_range = self._get_text_range(shape)
            language_id = self._resolve_office_constant(constants, language, PROOFING_LANGUAGE_ALIASES)
            text_range.LanguageID = language_id

            presentation.Save()
            return OperationResult(
                message="PowerPoint proofing language updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "language": language,
                    "language_id": int(language_id),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def spellcheck_slide(self, path: str, slide_index: int, include_notes: bool) -> SlideSpellcheckResult:
        source = self.resolve_document_path(path)

        with self._open_presentation(source, read_only=True) as presentation:
            slide = presentation.Slides(slide_index)
            fragments = self._iter_slide_spell_fragments(slide, include_notes=include_notes)

        issues: list[SpellingIssueSummary] = []
        with office_application("Word.Application", visible=False) as word:
            for fragment in fragments:
                spelling_issues = self._extract_spelling_issues(
                    word,
                    str(fragment["text"]),
                    fragment["language_id"] if isinstance(fragment["language_id"], int) else None,
                )
                counts: dict[str, int] = {}
                for issue in spelling_issues:
                    counts[issue] = counts.get(issue, 0) + 1
                for issue, occurrences in counts.items():
                    issues.append(
                        SpellingIssueSummary(
                            word=issue,
                            occurrences=occurrences,
                            slide_index=slide_index,
                            location=str(fragment["location"]),
                            shape_index=fragment["shape_index"] if isinstance(fragment["shape_index"], int) else None,
                            shape_name=fragment["shape_name"] if isinstance(fragment["shape_name"], str) else None,
                        )
                    )

        return SlideSpellcheckResult(
            file_path=str(source),
            slide_index=slide_index,
            issue_count=sum(issue.occurrences for issue in issues),
            issues=issues,
        )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def spellcheck_presentation(self, path: str, include_notes: bool) -> PresentationSpellcheckResult:
        source = self.resolve_document_path(path)

        all_fragments: list[tuple[int, list[dict[str, object]]]] = []
        with self._open_presentation(source, read_only=True) as presentation:
            for slide_index in range(1, int(presentation.Slides.Count) + 1):
                slide = presentation.Slides(slide_index)
                all_fragments.append((slide_index, self._iter_slide_spell_fragments(slide, include_notes=include_notes)))

        issues: list[SpellingIssueSummary] = []
        with office_application("Word.Application", visible=False) as word:
            for slide_index, fragments in all_fragments:
                for fragment in fragments:
                    spelling_issues = self._extract_spelling_issues(
                        word,
                        str(fragment["text"]),
                        fragment["language_id"] if isinstance(fragment["language_id"], int) else None,
                    )
                    counts: dict[str, int] = {}
                    for issue in spelling_issues:
                        counts[issue] = counts.get(issue, 0) + 1
                    for issue, occurrences in counts.items():
                        issues.append(
                            SpellingIssueSummary(
                                word=issue,
                                occurrences=occurrences,
                                slide_index=slide_index,
                                location=str(fragment["location"]),
                                shape_index=fragment["shape_index"] if isinstance(fragment["shape_index"], int) else None,
                                shape_name=fragment["shape_name"] if isinstance(fragment["shape_name"], str) else None,
                            )
                        )

        return PresentationSpellcheckResult(
            file_path=str(source),
            issue_count=sum(issue.occurrences for issue in issues),
            issues=issues,
        )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def translate_text(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        target_language: str,
        source_language: str | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        from win32com.client import constants

        with self._open_presentation_session(source, read_only=False, visible=True, with_window=True) as (powerpoint, presentation):
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            text_range = self._get_text_range(shape)
            text = str(text_range.Text).strip()
            if not text:
                raise ValueError("The selected shape does not contain text to translate")

            research = presentation.Research
            source_language_id = None
            if source_language is not None:
                source_language_id = self._resolve_office_constant(constants, source_language, PROOFING_LANGUAGE_ALIASES)
            else:
                with suppress(Exception):
                    source_language_id = int(text_range.LanguageID)
            target_language_id = self._resolve_office_constant(constants, target_language, PROOFING_LANGUAGE_ALIASES)

            with suppress(Exception):
                if source_language_id is not None:
                    research.SetLanguagePair(source_language_id, target_language_id)

            text_range.Select()

            executed_command = None
            last_error: Exception | None = None
            for command_name in ("TranslateSelected", "Translate"):
                try:
                    powerpoint.CommandBars.ExecuteMso(command_name)
                    executed_command = command_name
                    break
                except Exception as exc:
                    last_error = exc

            if executed_command is None:
                raise ValueError(
                    "PowerPoint translation UI is not available through COM on this installation"
                ) from last_error

            presentation.Save()
            return OperationResult(
                message="PowerPoint translation UI launched",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "source_language": source_language,
                    "resolved_source_language_id": source_language_id,
                    "target_language": target_language,
                    "resolved_target_language_id": int(target_language_id),
                    "executed_command": executed_command,
                    "note": "PowerPoint exposes translation through its UI. COM can launch that UI, but does not provide the translated text back programmatically.",
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def get_presenter_notes_all(self, path: str) -> PresentationNotesResult:
        source = self.resolve_document_path(path)

        with self._open_presentation(source, read_only=True) as presentation:
            slides = [
                SlideNotesResult(
                    file_path=str(source),
                    slide_index=slide_index,
                    notes_text=self._get_slide_notes_text(presentation.Slides(slide_index)),
                )
                for slide_index in range(1, int(presentation.Slides.Count) + 1)
            ]
            return PresentationNotesResult(file_path=str(source), slides=slides)

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def find_in_notes(self, path: str, query: str) -> PresentationTextSearchResult:
        source = self.resolve_document_path(path)

        with self._open_presentation(source, read_only=True) as presentation:
            matches: list[TextMatchSummary] = []
            for slide_index in range(1, int(presentation.Slides.Count) + 1):
                slide = presentation.Slides(slide_index)
                notes_text = self._get_slide_notes_text(slide)
                match = self._collect_text_matches(
                    slide_index=slide_index,
                    query=query,
                    text=notes_text,
                    location="notes",
                )
                if match is not None:
                    matches.append(match)

            return PresentationTextSearchResult(
                file_path=str(source),
                query=query,
                match_count=sum(match.occurrences for match in matches),
                matches=matches,
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def replace_notes_text(self, path: str, find_text: str, replace_text: str, create_backup: bool) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            replacement_count = 0
            updated_slides: list[dict[str, object]] = []
            for slide_index in range(1, int(presentation.Slides.Count) + 1):
                slide = presentation.Slides(slide_index)
                shape = self._get_notes_text_shape(slide, create_if_missing=False)
                if shape is None:
                    continue

                with suppress(Exception):
                    text_range = shape.TextFrame.TextRange
                    current_text = str(text_range.Text)
                    occurrences = current_text.count(find_text)
                    if occurrences:
                        text_range.Text = current_text.replace(find_text, replace_text)
                        replacement_count += occurrences
                        updated_slides.append(
                            {
                                "slide_index": slide_index,
                                "replacement_count": occurrences,
                            }
                        )

            presentation.Save()
            return OperationResult(
                message="PowerPoint notes text replacement completed",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "find_text": find_text,
                    "replace_text": replace_text,
                    "replacement_count": replacement_count,
                    "updated_slide_count": len(updated_slides),
                    "updated_slides": updated_slides,
                },
            )

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

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            title_shape_name, body_shape_name = self._apply_style_preset_to_slide(slide, preset)

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
    def add_row_to_table(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        row_index: int | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            table = self._require_table_shape(self._get_shape(slide, shape_index))
            added_row = table.Rows.Add(row_index if row_index is not None else -1)
            presentation.Save()
            return OperationResult(
                message="PowerPoint table row added",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "row_index": int(added_row.Index),
                    "row_count": int(table.Rows.Count),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def add_column_to_table(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        column_index: int | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            table = self._require_table_shape(self._get_shape(slide, shape_index))
            added_column = table.Columns.Add(column_index if column_index is not None else -1)
            presentation.Save()
            return OperationResult(
                message="PowerPoint table column added",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "column_index": int(added_column.Index),
                    "column_count": int(table.Columns.Count),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def delete_row_from_table(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        row_index: int,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            table = self._require_table_shape(self._get_shape(slide, shape_index))
            if int(table.Rows.Count) <= 1:
                raise ValueError("Cannot delete the only remaining row in a table")
            table.Rows(row_index).Delete()
            presentation.Save()
            return OperationResult(
                message="PowerPoint table row deleted",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "row_index": row_index,
                    "row_count": int(table.Rows.Count),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def delete_column_from_table(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        column_index: int,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            table = self._require_table_shape(self._get_shape(slide, shape_index))
            if int(table.Columns.Count) <= 1:
                raise ValueError("Cannot delete the only remaining column in a table")
            table.Columns(column_index).Delete()
            presentation.Save()
            return OperationResult(
                message="PowerPoint table column deleted",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "column_index": column_index,
                    "column_count": int(table.Columns.Count),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def merge_table_cells(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        row_index: int,
        column_index: int,
        merge_to_row_index: int,
        merge_to_column_index: int,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            table = self._require_table_shape(self._get_shape(slide, shape_index))
            source_cell = table.Cell(row_index, column_index)
            target_cell = table.Cell(merge_to_row_index, merge_to_column_index)
            source_cell.Merge(target_cell)
            presentation.Save()
            return OperationResult(
                message="PowerPoint table cells merged",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "row_index": row_index,
                    "column_index": column_index,
                    "merge_to_row_index": merge_to_row_index,
                    "merge_to_column_index": merge_to_column_index,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def split_table_cells(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        row_index: int,
        column_index: int,
        num_rows: int,
        num_columns: int,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            table = self._require_table_shape(self._get_shape(slide, shape_index))
            table.Cell(row_index, column_index).Split(num_rows, num_columns)
            presentation.Save()
            return OperationResult(
                message="PowerPoint table cell split",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "row_index": row_index,
                    "column_index": column_index,
                    "num_rows": num_rows,
                    "num_columns": num_columns,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_table_style(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        style_id: str | None,
        save_formatting: bool,
        first_row: bool | None,
        first_col: bool | None,
        last_row: bool | None,
        last_col: bool | None,
        horiz_banding: bool | None,
        vert_banding: bool | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            table = self._require_table_shape(self._get_shape(slide, shape_index))
            if style_id is not None:
                table.ApplyStyle(style_id, save_formatting)
            if first_row is not None:
                table.FirstRow = -1 if first_row else 0
            if first_col is not None:
                table.FirstCol = -1 if first_col else 0
            if last_row is not None:
                table.LastRow = -1 if last_row else 0
            if last_col is not None:
                table.LastCol = -1 if last_col else 0
            if horiz_banding is not None:
                table.HorizBanding = -1 if horiz_banding else 0
            if vert_banding is not None:
                table.VertBanding = -1 if vert_banding else 0

            presentation.Save()
            style_name = None
            with suppress(Exception):
                style_name = str(table.Style.Id).strip() or None
            return OperationResult(
                message="PowerPoint table style updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "style_id": style_id or style_name,
                    "first_row": bool(table.FirstRow),
                    "first_col": bool(table.FirstCol),
                    "last_row": bool(table.LastRow),
                    "last_col": bool(table.LastCol),
                    "horiz_banding": bool(table.HorizBanding),
                    "vert_banding": bool(table.VertBanding),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_table_cell_style(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        row_index: int,
        column_index: int,
        text: str | None,
        font_name: str | None,
        font_size: float | None,
        bold: bool | None,
        italic: bool | None,
        underline: bool | None,
        color: str | None,
        alignment: str | None,
        fill_color: str | None,
        fill_transparency: float | None,
        line_color: str | None,
        line_weight: float | None,
        line_transparency: float | None,
        line_visible: bool | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            table = self._require_table_shape(self._get_shape(slide, shape_index))
            cell_shape = self._get_table_cell_shape(table, row_index, column_index)
            text_range = self._get_text_range(cell_shape)
            self._apply_text_range_style(
                text_range,
                text=text,
                font_name=font_name,
                font_size=font_size,
                bold=bold,
                italic=italic,
                underline=underline,
                color=color,
                alignment=alignment,
            )
            self._apply_shape_fill_and_line(
                cell_shape,
                fill_color=fill_color,
                fill_transparency=fill_transparency,
                line_color=line_color,
                line_weight=line_weight,
                line_transparency=line_transparency,
                line_visible=line_visible,
            )

            presentation.Save()
            return OperationResult(
                message="PowerPoint table cell style updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "row_index": row_index,
                    "column_index": column_index,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_table_row_style(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        row_index: int,
        text: str | None,
        font_name: str | None,
        font_size: float | None,
        bold: bool | None,
        italic: bool | None,
        underline: bool | None,
        color: str | None,
        alignment: str | None,
        fill_color: str | None,
        fill_transparency: float | None,
        line_color: str | None,
        line_weight: float | None,
        line_transparency: float | None,
        line_visible: bool | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            table = self._require_table_shape(self._get_shape(slide, shape_index))
            column_count = int(table.Columns.Count)
            for current_column in range(1, column_count + 1):
                cell_shape = self._get_table_cell_shape(table, row_index, current_column)
                text_range = self._get_text_range(cell_shape)
                self._apply_text_range_style(
                    text_range,
                    text=text,
                    font_name=font_name,
                    font_size=font_size,
                    bold=bold,
                    italic=italic,
                    underline=underline,
                    color=color,
                    alignment=alignment,
                )
                self._apply_shape_fill_and_line(
                    cell_shape,
                    fill_color=fill_color,
                    fill_transparency=fill_transparency,
                    line_color=line_color,
                    line_weight=line_weight,
                    line_transparency=line_transparency,
                    line_visible=line_visible,
                )

            presentation.Save()
            return OperationResult(
                message="PowerPoint table row style updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "row_index": row_index,
                    "column_count": column_count,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_table_column_style(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        column_index: int,
        text: str | None,
        font_name: str | None,
        font_size: float | None,
        bold: bool | None,
        italic: bool | None,
        underline: bool | None,
        color: str | None,
        alignment: str | None,
        fill_color: str | None,
        fill_transparency: float | None,
        line_color: str | None,
        line_weight: float | None,
        line_transparency: float | None,
        line_visible: bool | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            table = self._require_table_shape(self._get_shape(slide, shape_index))
            row_count = int(table.Rows.Count)
            for current_row in range(1, row_count + 1):
                cell_shape = self._get_table_cell_shape(table, current_row, column_index)
                text_range = self._get_text_range(cell_shape)
                self._apply_text_range_style(
                    text_range,
                    text=text,
                    font_name=font_name,
                    font_size=font_size,
                    bold=bold,
                    italic=italic,
                    underline=underline,
                    color=color,
                    alignment=alignment,
                )
                self._apply_shape_fill_and_line(
                    cell_shape,
                    fill_color=fill_color,
                    fill_transparency=fill_transparency,
                    line_color=line_color,
                    line_weight=line_weight,
                    line_transparency=line_transparency,
                    line_visible=line_visible,
                )

            presentation.Save()
            return OperationResult(
                message="PowerPoint table column style updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "column_index": column_index,
                    "row_count": row_count,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def distribute_table_rows(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            table = self._require_table_shape(self._get_shape(slide, shape_index))
            row_count = int(table.Rows.Count)
            total_height = sum(float(table.Rows(index).Height) for index in range(1, row_count + 1))
            target_height = total_height / row_count
            for index in range(1, row_count + 1):
                table.Rows(index).Height = target_height

            presentation.Save()
            return OperationResult(
                message="PowerPoint table rows distributed",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "row_count": row_count,
                    "height": target_height,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def distribute_table_columns(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            table = self._require_table_shape(self._get_shape(slide, shape_index))
            column_count = int(table.Columns.Count)
            total_width = sum(float(table.Columns(index).Width) for index in range(1, column_count + 1))
            target_width = total_width / column_count
            for index in range(1, column_count + 1):
                table.Columns(index).Width = target_width

            presentation.Save()
            return OperationResult(
                message="PowerPoint table columns distributed",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "column_count": column_count,
                    "width": target_width,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def autofit_table(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            table = self._require_table_shape(shape)
            row_count = int(table.Rows.Count)
            column_count = int(table.Columns.Count)

            column_scores: list[int] = []
            row_values: list[list[str]] = []
            for row_index in range(1, row_count + 1):
                values_for_row: list[str] = []
                for column_index in range(1, column_count + 1):
                    text = ""
                    with suppress(Exception):
                        text = str(table.Cell(row_index, column_index).Shape.TextFrame.TextRange.Text).strip()
                    values_for_row.append(text)
                row_values.append(values_for_row)

            for column_index in range(column_count):
                score = 1
                for current_row in row_values:
                    score = max(score, len(current_row[column_index] or ""))
                column_scores.append(score)

            total_score = sum(column_scores) or column_count
            total_width = float(shape.Width)
            for column_index, score in enumerate(column_scores, start=1):
                table.Columns(column_index).Width = max(36.0, total_width * (score / total_score))

            estimated_heights = [self._estimate_table_row_height(values) for values in row_values]
            for row_index, height in enumerate(estimated_heights, start=1):
                table.Rows(row_index).Height = height

            presentation.Save()
            return OperationResult(
                message="PowerPoint table autofit applied",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "row_count": row_count,
                    "column_count": column_count,
                    "note": "Applied a content-based width and height adjustment heuristic using cell text lengths.",
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def table_from_csv(
        self,
        path: str,
        slide_index: int,
        csv_path: str,
        x: float,
        y: float,
        width: float,
        height: float,
        delimiter: str,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        csv_file = validate_file_path(
            csv_path,
            allowed_roots=self.settings.allowed_roots,
            allowed_suffixes=CSV_SUFFIXES,
            must_exist=True,
        )
        backup_path = self.maybe_create_backup(source, create_backup)

        with csv_file.open("r", encoding="utf-8-sig", newline="") as handle:
            rows = [row for row in csv.reader(handle, delimiter=delimiter)]

        if not rows:
            raise ValueError("The CSV file is empty")

        max_columns = max(len(row) for row in rows)
        normalized_rows = [row + [""] * (max_columns - len(row)) for row in rows]

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = slide.Shapes.AddTable(len(normalized_rows), max_columns, x, y, width, height)
            table = shape.Table
            for row_index, row_values in enumerate(normalized_rows, start=1):
                for column_index, cell_value in enumerate(row_values, start=1):
                    table.Cell(row_index, column_index).Shape.TextFrame.TextRange.Text = cell_value

            presentation.Save()
            return OperationResult(
                message="PowerPoint table created from CSV",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "csv_path": str(csv_file),
                    "shape_name": str(shape.Name),
                    "rows": len(normalized_rows),
                    "columns": max_columns,
                    "delimiter": delimiter,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def sort_table(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        column_index: int,
        descending: bool,
        has_header: bool,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        def sort_key(value: str) -> tuple[int, object]:
            text = value.strip()
            try:
                return (0, float(text.replace(",", ".")))
            except Exception:
                return (1, text.lower())

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            table = self._require_table_shape(self._get_shape(slide, shape_index))
            row_count = int(table.Rows.Count)
            column_count = int(table.Columns.Count)
            if not 1 <= column_index <= column_count:
                raise ValueError(f"The table does not contain column_index={column_index}")

            values: list[list[str]] = []
            for row_index in range(1, row_count + 1):
                row_values: list[str] = []
                for current_column in range(1, column_count + 1):
                    text = ""
                    with suppress(Exception):
                        text = str(table.Cell(row_index, current_column).Shape.TextFrame.TextRange.Text)
                    row_values.append(text)
                values.append(row_values)

            header = values[:1] if has_header and values else []
            body = values[1:] if has_header else values
            body.sort(key=lambda row: sort_key(row[column_index - 1]), reverse=descending)
            sorted_values = header + body

            for row_index, row_values in enumerate(sorted_values, start=1):
                for current_column, cell_value in enumerate(row_values, start=1):
                    table.Cell(row_index, current_column).Shape.TextFrame.TextRange.Text = cell_value

            presentation.Save()
            return OperationResult(
                message="PowerPoint table sorted",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "column_index": column_index,
                    "descending": descending,
                    "has_header": has_header,
                    "row_count": row_count,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def table_from_excel_range(
        self,
        path: str,
        slide_index: int,
        excel_path: str,
        sheet: str,
        cell_range: str,
        shape_index: int | None,
        x: float,
        y: float,
        width: float,
        height: float,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        excel_file = validate_file_path(
            excel_path,
            allowed_roots=self.settings.allowed_roots,
            allowed_suffixes=EXCEL_SUFFIXES,
            must_exist=True,
        )
        backup_path = self.maybe_create_backup(source, create_backup)

        with office_application("Excel.Application", visible=self.settings.office_visible) as excel:
            workbook = None
            try:
                workbook = excel.Workbooks.Open(str(excel_file), ReadOnly=True)
                worksheet = workbook.Worksheets(sheet)
                values = self._normalize_excel_range_values(worksheet.Range(cell_range).Value)
            finally:
                if workbook is not None:
                    with suppress(Exception):
                        workbook.Close(SaveChanges=False)

        if not values:
            raise ValueError("The Excel range is empty")

        row_count = len(values)
        column_count = max(len(row) for row in values)
        normalized_rows = [row + [""] * (column_count - len(row)) for row in values]

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            target_x = x
            target_y = y
            target_width = width
            target_height = height
            target_name = None

            if shape_index is not None:
                existing_shape = self._get_shape(slide, shape_index)
                self._require_table_shape(existing_shape)
                target_x = float(existing_shape.Left)
                target_y = float(existing_shape.Top)
                target_width = float(existing_shape.Width)
                target_height = float(existing_shape.Height)
                target_name = str(existing_shape.Name)
                existing_shape.Delete()

            shape = slide.Shapes.AddTable(row_count, column_count, target_x, target_y, target_width, target_height)
            if target_name is not None:
                with suppress(Exception):
                    shape.Name = target_name

            table = shape.Table
            for row_index, row_values in enumerate(normalized_rows, start=1):
                for column_index, cell_value in enumerate(row_values, start=1):
                    table.Cell(row_index, column_index).Shape.TextFrame.TextRange.Text = cell_value

            presentation.Save()
            return OperationResult(
                message="PowerPoint table created from Excel range",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "excel_path": str(excel_file),
                    "sheet": sheet,
                    "cell_range": cell_range,
                    "shape_name": str(shape.Name),
                    "rows": row_count,
                    "columns": column_count,
                    "refreshed_shape_index": shape_index,
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
    def refresh_chart(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation_session(source, read_only=False, visible=True, with_window=True) as (_, presentation):
            slide = presentation.Slides(slide_index)
            chart = self._require_chart_shape(self._get_shape(slide, shape_index))

            with suppress(Exception):
                chart.ChartData.Activate()
            with suppress(Exception):
                workbook = chart.ChartData.Workbook
                with suppress(Exception):
                    workbook.RefreshAll()
                with suppress(Exception):
                    workbook.Application.CalculateFullRebuild()
                with suppress(Exception):
                    workbook.Application.Calculate()
            with suppress(Exception):
                chart.Refresh()

            presentation.Save()
            return OperationResult(
                message="PowerPoint chart refreshed",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "series_count": int(chart.SeriesCollection().Count),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_chart_axis_scale(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        axis_kind: str,
        minimum_scale: float | int | None,
        maximum_scale: float | int | None,
        major_unit: float | int | None,
        minor_unit: float | int | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation_session(source, read_only=False, visible=True, with_window=True) as (_, presentation):
            slide = presentation.Slides(slide_index)
            chart = self._require_chart_shape(self._get_shape(slide, shape_index))
            axis = self._get_chart_axis(chart, axis_kind)

            if minimum_scale is not None:
                with suppress(Exception):
                    axis.MinimumScaleIsAuto = False
                axis.MinimumScale = float(minimum_scale)
            if maximum_scale is not None:
                with suppress(Exception):
                    axis.MaximumScaleIsAuto = False
                axis.MaximumScale = float(maximum_scale)
            if major_unit is not None:
                with suppress(Exception):
                    axis.MajorUnitIsAuto = False
                axis.MajorUnit = float(major_unit)
            if minor_unit is not None:
                with suppress(Exception):
                    axis.MinorUnitIsAuto = False
                axis.MinorUnit = float(minor_unit)

            presentation.Save()
            return OperationResult(
                message="PowerPoint chart axis scale updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "axis_kind": axis_kind,
                    "minimum_scale": float(axis.MinimumScale) if minimum_scale is not None else None,
                    "maximum_scale": float(axis.MaximumScale) if maximum_scale is not None else None,
                    "major_unit": float(axis.MajorUnit) if major_unit is not None else None,
                    "minor_unit": float(axis.MinorUnit) if minor_unit is not None else None,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_chart_series_order(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        series_index: int,
        position: int,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            chart = self._require_chart_shape(self._get_shape(slide, shape_index))
            series_collection = chart.SeriesCollection()
            series_count = int(series_collection.Count)
            if position > series_count:
                raise ValueError(f"The chart does not support position={position}; it only has {series_count} series")

            series = self._get_chart_series(chart, series_index)
            series.PlotOrder = position

            presentation.Save()
            return OperationResult(
                message="PowerPoint chart series order updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "series_index": series_index,
                    "position": position,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def add_chart_series(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        name: str,
        values: list[float | int],
        categories: list[object],
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation_session(source, read_only=False, visible=True, with_window=True) as (_, presentation):
            slide = presentation.Slides(slide_index)
            chart = self._require_chart_shape(self._get_shape(slide, shape_index))
            created_series = chart.SeriesCollection().NewSeries()
            created_series.Name = name
            if categories:
                created_series.XValues = tuple(categories)
            created_series.Values = tuple(values)

            presentation.Save()
            return OperationResult(
                message="PowerPoint chart series added",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "series_name": name,
                    "series_count": int(chart.SeriesCollection().Count),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def delete_chart_series(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        series_index: int,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            chart = self._require_chart_shape(self._get_shape(slide, shape_index))
            series = self._get_chart_series(chart, series_index)
            series_name = None
            with suppress(Exception):
                series_name = str(series.Name).strip() or None
            series.Delete()

            presentation.Save()
            return OperationResult(
                message="PowerPoint chart series deleted",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "series_index": series_index,
                    "series_name": series_name,
                    "series_count": int(chart.SeriesCollection().Count),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_chart_data_labels(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        series_index: int | None,
        visible: bool | None,
        show_value: bool | None,
        show_category_name: bool | None,
        show_series_name: bool | None,
        show_percentage: bool | None,
        separator: str | None,
        position: str | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            chart = self._require_chart_shape(self._get_shape(slide, shape_index))
            series_items = (
                [self._get_chart_series(chart, series_index)]
                if series_index is not None
                else [chart.SeriesCollection(index) for index in range(1, int(chart.SeriesCollection().Count) + 1)]
            )

            for series in series_items:
                if visible is False and all(
                    field is None
                    for field in (
                        show_value,
                        show_category_name,
                        show_series_name,
                        show_percentage,
                        separator,
                        position,
                    )
                ):
                    series.HasDataLabels = False
                    continue

                if visible is True or any(
                    field is not None
                    for field in (
                        show_value,
                        show_category_name,
                        show_series_name,
                        show_percentage,
                        separator,
                        position,
                    )
                ):
                    series.ApplyDataLabels()

                labels = series.DataLabels()
                if visible is not None:
                    series.HasDataLabels = bool(visible)
                if show_value is not None:
                    labels.ShowValue = bool(show_value)
                if show_category_name is not None:
                    labels.ShowCategoryName = bool(show_category_name)
                if show_series_name is not None:
                    labels.ShowSeriesName = bool(show_series_name)
                if show_percentage is not None:
                    labels.ShowPercentage = bool(show_percentage)
                if separator is not None:
                    labels.Separator = separator
                if position is not None:
                    labels.Position = resolve_chart_data_label_position(position)

            presentation.Save()
            return OperationResult(
                message="PowerPoint chart data labels updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "series_index": series_index,
                    "visible": visible,
                    "show_value": show_value,
                    "show_category_name": show_category_name,
                    "show_series_name": show_series_name,
                    "show_percentage": show_percentage,
                    "separator": separator,
                    "position": position,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_chart_gridlines(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        axis_kind: str,
        major: bool | None,
        minor: bool | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            chart = self._require_chart_shape(self._get_shape(slide, shape_index))
            axis = self._get_chart_axis(chart, axis_kind)

            if major is not None:
                axis.HasMajorGridlines = bool(major)
            if minor is not None:
                axis.HasMinorGridlines = bool(minor)

            presentation.Save()
            return OperationResult(
                message="PowerPoint chart gridlines updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "axis_kind": axis_kind,
                    "major": bool(axis.HasMajorGridlines) if major is not None else None,
                    "minor": bool(axis.HasMinorGridlines) if minor is not None else None,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_chart_colors(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        series_index: int | None,
        fill_color: str | None,
        line_color: str | None,
        series_fill_colors: list[str],
        series_line_colors: list[str],
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            chart = self._require_chart_shape(self._get_shape(slide, shape_index))
            series_count = int(chart.SeriesCollection().Count)

            if series_index is not None:
                self._apply_chart_series_colors(self._get_chart_series(chart, series_index), fill_color, line_color)
            else:
                for index in range(1, series_count + 1):
                    series = chart.SeriesCollection(index)
                    current_fill = fill_color
                    current_line = line_color
                    if index - 1 < len(series_fill_colors):
                        current_fill = series_fill_colors[index - 1]
                    if index - 1 < len(series_line_colors):
                        current_line = series_line_colors[index - 1]
                    self._apply_chart_series_colors(series, current_fill, current_line)

            presentation.Save()
            return OperationResult(
                message="PowerPoint chart colors updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "series_index": series_index,
                    "fill_color": fill_color,
                    "line_color": line_color,
                    "series_fill_colors": series_fill_colors,
                    "series_line_colors": series_line_colors,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def change_chart_type(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        chart_type: str,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation_session(source, read_only=False, visible=True, with_window=True) as (_, presentation):
            slide = presentation.Slides(slide_index)
            chart = self._require_chart_shape(self._get_shape(slide, shape_index))
            chart.ChartType = resolve_chart_type(chart_type)

            presentation.Save()
            return OperationResult(
                message="PowerPoint chart type updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "chart_type": int(chart.ChartType),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def link_chart_to_excel(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        excel_path: str,
        sheet: str,
        cell_range: str,
        plot_by: str | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)
        excel_file = validate_file_path(
            excel_path,
            allowed_roots=self.settings.allowed_roots,
            allowed_suffixes=EXCEL_SUFFIXES,
            must_exist=True,
        )

        workbook_name = excel_file.name
        sheet_name = sheet.replace("'", "''")
        workbook_dir = str(excel_file.parent).replace("'", "''")
        source_reference = f"='{workbook_dir}\\[{workbook_name}]{sheet_name}'!{cell_range}"
        plot_by_value = resolve_chart_plot_by(plot_by) if plot_by is not None else None

        with self._open_presentation_session(source, read_only=False, visible=True, with_window=True) as (_, presentation):
            slide = presentation.Slides(slide_index)
            chart = self._require_chart_shape(self._get_shape(slide, shape_index))

            chart.ChartData.Activate()
            if plot_by_value is None:
                chart.SetSourceData(Source=source_reference)
            else:
                chart.SetSourceData(Source=source_reference, PlotBy=plot_by_value)
            with suppress(Exception):
                chart.Refresh()

            is_linked = None
            with suppress(Exception):
                is_linked = bool(chart.ChartData.IsLinked)

            presentation.Save()
            return OperationResult(
                message="PowerPoint chart linked to Excel",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "excel_path": str(excel_file),
                    "sheet": sheet,
                    "cell_range": cell_range,
                    "plot_by": plot_by,
                    "source_reference": source_reference,
                    "is_linked": is_linked,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def break_chart_link(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation_session(source, read_only=False, visible=True, with_window=True) as (_, presentation):
            slide = presentation.Slides(slide_index)
            chart = self._require_chart_shape(self._get_shape(slide, shape_index))
            chart.ChartData.Activate()

            was_linked = None
            with suppress(Exception):
                was_linked = bool(chart.ChartData.IsLinked)

            if was_linked is False:
                raise ValueError("The selected chart is not linked to an external Excel workbook")

            chart.ChartData.BreakLink()

            is_linked = None
            with suppress(Exception):
                is_linked = bool(chart.ChartData.IsLinked)

            presentation.Save()
            return OperationResult(
                message="PowerPoint chart link removed",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "was_linked": was_linked,
                    "is_linked": is_linked,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def export_chart_data(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        out_path: str | None,
        export_format: str,
    ) -> ChartDataExportResult:
        source = self.resolve_document_path(path)

        with self._open_presentation(source, read_only=True) as presentation:
            slide = presentation.Slides(slide_index)
            chart = self._require_chart_shape(self._get_shape(slide, shape_index))
            series_summaries = self._collect_chart_series_summaries(chart)
            categories = next((series.categories for series in series_summaries if series.categories), [])
            chart_type = None
            with suppress(Exception):
                chart_type = int(chart.ChartType)

            exported_path = None
            if out_path is not None:
                allowed_suffixes = (".json",) if export_format == "json" else (".csv",)
                target = self.resolve_output_path(out_path, allowed_suffixes=allowed_suffixes)
                if export_format == "json":
                    payload = {
                        "file_path": str(source),
                        "slide_index": slide_index,
                        "shape_index": shape_index,
                        "chart_type": chart_type,
                        "categories": categories,
                        "series": [series.model_dump() for series in series_summaries],
                    }
                    target.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
                else:
                    row_count = max(
                        [len(categories), *[len(series.values) for series in series_summaries]],
                        default=0,
                    )
                    with target.open("w", encoding="utf-8", newline="") as handle:
                        writer = csv.writer(handle)
                        writer.writerow(["Category", *[(series.name or f"Series {series.series_index}") for series in series_summaries]])
                        for row_index in range(row_count):
                            row = [categories[row_index] if row_index < len(categories) else ""]
                            for series in series_summaries:
                                row.append(series.values[row_index] if row_index < len(series.values) else "")
                            writer.writerow(row)
                exported_path = str(target)

            return ChartDataExportResult(
                file_path=str(source),
                slide_index=slide_index,
                shape_index=shape_index,
                chart_type=chart_type,
                categories=categories,
                series=series_summaries,
                exported_path=exported_path,
                export_format=export_format,
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
            node = self._get_smartart_node(smartart, node_index)
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
    def add_smartart_node(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        node_index: int,
        position: str,
        node_type: str,
        text: str,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        from win32com.client import constants

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            smartart = self._require_smartart_shape(self._get_shape(slide, shape_index))
            anchor_node = self._get_smartart_node(smartart, node_index)
            new_node = anchor_node.AddNode(
                self._resolve_office_constant(constants, position, SMARTART_NODE_POSITIONS),
                self._resolve_office_constant(constants, node_type, SMARTART_NODE_TYPES),
            )
            new_node.TextFrame2.TextRange.Text = text
            new_level = None
            with suppress(Exception):
                new_level = int(new_node.Level)

            presentation.Save()
            return OperationResult(
                message="PowerPoint SmartArt node added",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "anchor_node_index": node_index,
                    "position": resolve_smartart_node_position(position),
                    "node_type": resolve_smartart_node_type(node_type),
                    "text": text,
                    "level": new_level,
                    "node_count": int(smartart.AllNodes.Count),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def delete_smartart_node(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        node_index: int,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            smartart = self._require_smartart_shape(self._get_shape(slide, shape_index))
            node = self._get_smartart_node(smartart, node_index)
            deleted_text = ""
            with suppress(Exception):
                deleted_text = str(node.TextFrame2.TextRange.Text)
            node.Delete()
            presentation.Save()
            return OperationResult(
                message="PowerPoint SmartArt node deleted",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "node_index": node_index,
                    "text": deleted_text,
                    "node_count": int(smartart.AllNodes.Count),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def promote_smartart_node(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        node_index: int,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            smartart = self._require_smartart_shape(self._get_shape(slide, shape_index))
            node = self._get_smartart_node(smartart, node_index)
            node.Promote()
            new_level = None
            with suppress(Exception):
                new_level = int(node.Level)
            presentation.Save()
            return OperationResult(
                message="PowerPoint SmartArt node promoted",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "node_index": node_index,
                    "level": new_level,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def demote_smartart_node(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        node_index: int,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            smartart = self._require_smartart_shape(self._get_shape(slide, shape_index))
            node = self._get_smartart_node(smartart, node_index)
            node.Demote()
            new_level = None
            with suppress(Exception):
                new_level = int(node.Level)
            presentation.Save()
            return OperationResult(
                message="PowerPoint SmartArt node demoted",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "node_index": node_index,
                    "level": new_level,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def reorder_smartart_node(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        node_index: int,
        direction: str,
        steps: int,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)
        resolved_direction = resolve_smartart_reorder_direction(direction)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            smartart = self._require_smartart_shape(self._get_shape(slide, shape_index))
            node = self._get_smartart_node(smartart, node_index)

            for _ in range(steps):
                if resolved_direction == "up":
                    node.ReorderUp()
                else:
                    node.ReorderDown()

            presentation.Save()
            return OperationResult(
                message="PowerPoint SmartArt node reordered",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "node_index": node_index,
                    "direction": resolved_direction,
                    "steps": steps,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_smartart_style(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        style: str,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation_session(source, read_only=False, visible=True, with_window=True) as (powerpoint, presentation):
            slide = presentation.Slides(slide_index)
            smartart = self._require_smartart_shape(self._get_shape(slide, shape_index))
            quick_style = self._resolve_smartart_quickstyle(powerpoint, style)
            smartart.QuickStyle = quick_style
            presentation.Save()
            return OperationResult(
                message="PowerPoint SmartArt style updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "style": self._describe_smartart_collection_item(quick_style),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_smartart_color_theme(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        color_theme: str,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation_session(source, read_only=False, visible=True, with_window=True) as (powerpoint, presentation):
            slide = presentation.Slides(slide_index)
            smartart = self._require_smartart_shape(self._get_shape(slide, shape_index))
            color_style = self._resolve_smartart_color_theme(powerpoint, color_theme)
            smartart.Color = color_style
            presentation.Save()
            return OperationResult(
                message="PowerPoint SmartArt color theme updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "color_theme": self._describe_smartart_collection_item(color_style),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def convert_smartart_to_shapes(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation_session(source, read_only=False, visible=True, with_window=True) as (powerpoint, presentation):
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            self._require_smartart_shape(shape)

            conversion_method = None
            converted_shape_names: list[str] = []

            with suppress(Exception):
                shape_range = shape.Ungroup()
                converted_shape_names = self._collect_shape_range_names(shape_range)
                if converted_shape_names:
                    conversion_method = "ungroup"

            if conversion_method is None:
                with suppress(Exception):
                    presentation.Windows(1).View.GotoSlide(slide_index)
                shape.Select()

                last_error: Exception | None = None
                for command_name in SMARTART_CONVERT_COMMANDS:
                    try:
                        powerpoint.CommandBars.ExecuteMso(command_name)
                        selection = presentation.Windows(1).Selection.ShapeRange
                        converted_shape_names = self._collect_shape_range_names(selection)
                        conversion_method = f"execute_mso:{command_name}"
                        break
                    except Exception as exc:
                        last_error = exc

                if conversion_method is None:
                    raise ValueError(
                        "PowerPoint could not convert the SmartArt to editable shapes on this installation"
                    ) from last_error

            presentation.Save()
            return OperationResult(
                message="PowerPoint SmartArt converted to shapes",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "conversion_method": conversion_method,
                    "shape_count": len(converted_shape_names),
                    "shape_names": converted_shape_names,
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
    def reroute_connectors(
        self,
        path: str,
        slide_index: int,
        shape_indexes: list[int],
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)

            connector_shapes: list[tuple[int, object]] = []
            if shape_indexes:
                for current_shape_index in shape_indexes:
                    connector = self._get_shape(slide, current_shape_index)
                    if not self._shape_is_connector(connector):
                        raise ValueError(f"The selected shape is not a connector: shape_index={current_shape_index}")
                    connector_shapes.append((current_shape_index, connector))
            else:
                connector_shapes = [
                    (current_shape_index, connector)
                    for current_shape_index, connector in self._iter_shapes(slide)
                    if self._shape_is_connector(connector)
                ]

            rerouted_indexes: list[int] = []
            skipped_indexes: list[int] = []
            skipped_names: list[str] = []
            for current_shape_index, connector in connector_shapes:
                try:
                    connector.RerouteConnections()
                    rerouted_indexes.append(current_shape_index)
                except Exception:
                    skipped_indexes.append(current_shape_index)
                    with suppress(Exception):
                        skipped_names.append(str(connector.Name))

            presentation.Save()
            return OperationResult(
                message="PowerPoint connectors rerouted",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "requested_shape_indexes": shape_indexes,
                    "rerouted_indexes": rerouted_indexes,
                    "rerouted_count": len(rerouted_indexes),
                    "skipped_indexes": skipped_indexes,
                    "skipped_names": skipped_names,
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
    def duplicate_shape(self, path: str, slide_index: int, shape_index: int, create_backup: bool) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            duplicated_range = shape.Duplicate()
            duplicated_shape = duplicated_range(1)
            presentation.Save()

            return OperationResult(
                message="PowerPoint shape duplicated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "source_shape_index": shape_index,
                    "shape_name": str(duplicated_shape.Name),
                    "left": float(duplicated_shape.Left),
                    "top": float(duplicated_shape.Top),
                    "width": float(duplicated_shape.Width),
                    "height": float(duplicated_shape.Height),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def delete_shape(
        self,
        path: str,
        slide_index: int,
        shape_index: int | None,
        shape_name: str | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            resolved_index, shape = self._resolve_shape_selector(
                slide,
                shape_index=shape_index,
                shape_name=shape_name,
            )
            resolved_name = str(shape.Name)
            shape.Delete()
            presentation.Save()

            return OperationResult(
                message="PowerPoint shape deleted",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": resolved_index,
                    "shape_name": resolved_name,
                    "remaining_shape_count": int(slide.Shapes.Count),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def rename_shape(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        name: str,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            previous_name = str(shape.Name)
            shape.Name = name
            presentation.Save()

            return OperationResult(
                message="PowerPoint shape renamed",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "previous_name": previous_name,
                    "name": name,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def find_shapes(
        self,
        path: str,
        slide_index: int,
        shape_name_contains: str | None,
        text_contains: str | None,
        shape_type: str | None,
    ) -> ShapeSearchResult:
        source = self.resolve_document_path(path)
        normalized_name = shape_name_contains.strip().lower() if shape_name_contains is not None else None
        normalized_text = text_contains.strip().lower() if text_contains is not None else None
        resolved_shape_type = resolve_shape_type(shape_type) if shape_type is not None else None

        with self._open_presentation(source, read_only=True) as presentation:
            slide = presentation.Slides(slide_index)
            matches: list[ShapeSummary] = []
            for index, shape in self._iter_shapes(slide):
                summary = self._shape_summary(shape, index)
                if normalized_name is not None and normalized_name not in summary.name.lower():
                    continue
                if normalized_text is not None and normalized_text not in (summary.text_preview or "").lower():
                    continue
                if resolved_shape_type is not None:
                    current_shape_type = summary.shape_type
                    with suppress(Exception):
                        current_shape_type = int(shape.AutoShapeType)
                    if current_shape_type != resolved_shape_type:
                        continue
                matches.append(summary)

            return ShapeSearchResult(
                file_path=str(source),
                slide_index=slide_index,
                match_count=len(matches),
                shapes=matches,
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def group_shapes(
        self,
        path: str,
        slide_index: int,
        shape_indexes: list[int],
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            grouped_shape = self._get_shape_range(slide, shape_indexes).Group()
            presentation.Save()

            return OperationResult(
                message="PowerPoint shapes grouped",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_indexes": shape_indexes,
                    "shape_name": str(grouped_shape.Name),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def ungroup_shapes(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            previous_name = str(shape.Name)
            shape_range = shape.Ungroup()
            ungrouped_names: list[str] = []
            with suppress(Exception):
                for index in range(1, int(shape_range.Count) + 1):
                    ungrouped_names.append(str(shape_range(index).Name))
            presentation.Save()

            return OperationResult(
                message="PowerPoint shapes ungrouped",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "previous_group_name": previous_name,
                    "ungrouped_count": len(ungrouped_names),
                    "shape_names": ungrouped_names,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def align_shapes(
        self,
        path: str,
        slide_index: int,
        shape_indexes: list[int],
        alignment: str,
        relative_to_slide: bool,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape_range = self._get_shape_range(slide, shape_indexes)
            shape_range.Align(resolve_shape_alignment(alignment), -1 if relative_to_slide else 0)
            presentation.Save()

            return OperationResult(
                message="PowerPoint shapes aligned",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_indexes": shape_indexes,
                    "alignment": normalize_powerpoint_token(alignment),
                    "relative_to_slide": relative_to_slide,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def distribute_shapes(
        self,
        path: str,
        slide_index: int,
        shape_indexes: list[int],
        direction: str,
        relative_to_slide: bool,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape_range = self._get_shape_range(slide, shape_indexes)
            shape_range.Distribute(resolve_shape_distribution(direction), -1 if relative_to_slide else 0)
            presentation.Save()

            return OperationResult(
                message="PowerPoint shapes distributed",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_indexes": shape_indexes,
                    "direction": normalize_powerpoint_token(direction),
                    "relative_to_slide": relative_to_slide,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def resize_shapes(
        self,
        path: str,
        slide_index: int,
        shape_indexes: list[int],
        mode: str,
        reference_shape_index: int | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)
        normalized_mode = resolve_shape_resize_mode(mode)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            reference_index = reference_shape_index or shape_indexes[0]
            reference_shape = self._get_shape(slide, reference_index)
            reference_width = float(reference_shape.Width)
            reference_height = float(reference_shape.Height)

            updated_shapes: list[int] = []
            for current_index in shape_indexes:
                if current_index == reference_index:
                    continue
                shape = self._get_shape(slide, current_index)
                if normalized_mode in {"width", "both"}:
                    shape.Width = reference_width
                if normalized_mode in {"height", "both"}:
                    shape.Height = reference_height
                updated_shapes.append(current_index)

            presentation.Save()
            return OperationResult(
                message="PowerPoint shapes resized",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_indexes": shape_indexes,
                    "updated_shapes": updated_shapes,
                    "mode": normalized_mode,
                    "reference_shape_index": reference_index,
                    "reference_width": reference_width,
                    "reference_height": reference_height,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def rotate_shape(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        rotation: float,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            shape.Rotation = rotation
            presentation.Save()
            return OperationResult(
                message="PowerPoint shape rotation updated",
                file_path=str(source),
                backup_path=backup_path,
                details={"slide_index": slide_index, "shape_index": shape_index, "rotation": float(shape.Rotation)},
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def flip_shape(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        direction: str,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            shape.Flip(resolve_shape_flip(direction))
            presentation.Save()
            return OperationResult(
                message="PowerPoint shape flipped",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "direction": normalize_powerpoint_token(direction),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_shape_position(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        x: float,
        y: float,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            shape.Left = x
            shape.Top = y
            presentation.Save()
            return OperationResult(
                message="PowerPoint shape position updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "left": float(shape.Left),
                    "top": float(shape.Top),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_shape_size(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        width: float | None,
        height: float | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            if width is not None:
                shape.Width = width
            if height is not None:
                shape.Height = height
            presentation.Save()
            return OperationResult(
                message="PowerPoint shape size updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "width": float(shape.Width),
                    "height": float(shape.Height),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def lock_aspect_ratio(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        lock: bool,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            shape.LockAspectRatio = -1 if lock else 0
            presentation.Save()
            return OperationResult(
                message="PowerPoint shape aspect ratio updated",
                file_path=str(source),
                backup_path=backup_path,
                details={"slide_index": slide_index, "shape_index": shape_index, "lock": lock},
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_shape_visibility(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        visible: bool,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            shape.Visible = -1 if visible else 0
            presentation.Save()
            return OperationResult(
                message="PowerPoint shape visibility updated",
                file_path=str(source),
                backup_path=backup_path,
                details={"slide_index": slide_index, "shape_index": shape_index, "visible": visible},
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_shape_z_order(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        command: str,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)
        normalized_command = normalize_powerpoint_token(command)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            shape.ZOrder(resolve_shape_z_order(command))
            presentation.Save()
            return OperationResult(
                message="PowerPoint shape z-order updated",
                file_path=str(source),
                backup_path=backup_path,
                details={"slide_index": slide_index, "shape_index": shape_index, "command": normalized_command},
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def merge_shapes(
        self,
        path: str,
        slide_index: int,
        shape_indexes: list[int],
        mode: str,
        primary_shape_index: int | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)
        merge_mode = resolve_shape_merge_mode(mode)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape_range = self._get_shape_range(slide, shape_indexes)
            primary_shape = self._get_shape(slide, primary_shape_index or shape_indexes[0])
            merged_shape = shape_range.MergeShapes(merge_mode, primary_shape)
            presentation.Save()

            return OperationResult(
                message="PowerPoint shapes merged",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_indexes": shape_indexes,
                    "mode": normalize_powerpoint_token(mode),
                    "primary_shape_index": primary_shape_index or shape_indexes[0],
                    "shape_name": str(merged_shape.Name),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def crop_shape_to_content(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            picture_format = shape.PictureFormat

            crop_left = float(picture_format.CropLeft)
            crop_right = float(picture_format.CropRight)
            crop_top = float(picture_format.CropTop)
            crop_bottom = float(picture_format.CropBottom)

            new_left = float(shape.Left) + crop_left
            new_top = float(shape.Top) + crop_top
            new_width = float(shape.Width) - crop_left - crop_right
            new_height = float(shape.Height) - crop_top - crop_bottom
            if new_width <= 0 or new_height <= 0:
                raise ValueError("The current crop margins would produce a non-positive shape size")

            picture_format.CropLeft = 0
            picture_format.CropRight = 0
            picture_format.CropTop = 0
            picture_format.CropBottom = 0
            shape.Left = new_left
            shape.Top = new_top
            shape.Width = new_width
            shape.Height = new_height

            presentation.Save()
            return OperationResult(
                message="PowerPoint shape cropped to visible content",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "previous_crop": {
                        "left": crop_left,
                        "right": crop_right,
                        "top": crop_top,
                        "bottom": crop_bottom,
                    },
                    "left": float(shape.Left),
                    "top": float(shape.Top),
                    "width": float(shape.Width),
                    "height": float(shape.Height),
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
    def insert_svg(
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
            allowed_suffixes=SVG_SUFFIXES,
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
                message="SVG inserted into PowerPoint slide",
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
    def add_video(
        self,
        path: str,
        slide_index: int,
        media_path: str,
        x: float,
        y: float,
        width: float | None,
        height: float | None,
        link_to_file: bool,
        save_with_document: bool,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        media = validate_file_path(
            media_path,
            allowed_roots=self.settings.allowed_roots,
            allowed_suffixes=VIDEO_SUFFIXES,
            must_exist=True,
        )
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = slide.Shapes.AddMediaObject2(
                str(media),
                -1 if link_to_file else 0,
                -1 if save_with_document else 0,
                x,
                y,
                width if width is not None else -1,
                height if height is not None else -1,
            )
            presentation.Save()
            return OperationResult(
                message="Video inserted into PowerPoint slide",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "media_path": str(media),
                    "shape_name": str(shape.Name),
                    "left": float(shape.Left),
                    "top": float(shape.Top),
                    "width": float(shape.Width),
                    "height": float(shape.Height),
                    "link_to_file": link_to_file,
                    "save_with_document": save_with_document,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_video_playback(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        autoplay: bool | None,
        loop_until_stopped: bool | None,
        pause_animation: bool | None,
        hide_while_not_playing: bool | None,
        stop_after_slides: int | None,
        volume: float | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            play_settings = self._get_play_settings(shape)
            media_format = self._get_media_format(shape)

            if autoplay is not None:
                play_settings.PlayOnEntry = -1 if autoplay else 0
            if loop_until_stopped is not None:
                play_settings.LoopUntilStopped = -1 if loop_until_stopped else 0
            if pause_animation is not None:
                play_settings.PauseAnimation = -1 if pause_animation else 0
            if hide_while_not_playing is not None:
                play_settings.HideWhileNotPlaying = -1 if hide_while_not_playing else 0
            if stop_after_slides is not None:
                play_settings.StopAfterSlides = stop_after_slides
            if volume is not None:
                media_format.Volume = volume

            presentation.Save()
            return OperationResult(
                message="PowerPoint video playback updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "autoplay": autoplay,
                    "loop_until_stopped": loop_until_stopped,
                    "pause_animation": pause_animation,
                    "hide_while_not_playing": hide_while_not_playing,
                    "stop_after_slides": stop_after_slides,
                    "volume": float(media_format.Volume),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def trim_video(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        start_point: int | None,
        end_point: int | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            media_format = self._get_media_format(shape)

            if start_point is not None:
                media_format.StartPoint = start_point
            if end_point is not None:
                media_format.EndPoint = end_point
            if int(media_format.StartPoint) >= int(media_format.EndPoint):
                raise ValueError("The media trim start point must be less than the end point")

            presentation.Save()
            return OperationResult(
                message="PowerPoint video trim updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "start_point": int(media_format.StartPoint),
                    "end_point": int(media_format.EndPoint),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def add_audio(
        self,
        path: str,
        slide_index: int,
        media_path: str,
        x: float,
        y: float,
        width: float | None,
        height: float | None,
        link_to_file: bool,
        save_with_document: bool,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        media = validate_file_path(
            media_path,
            allowed_roots=self.settings.allowed_roots,
            allowed_suffixes=AUDIO_SUFFIXES,
            must_exist=True,
        )
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = slide.Shapes.AddMediaObject2(
                str(media),
                -1 if link_to_file else 0,
                -1 if save_with_document else 0,
                x,
                y,
                width if width is not None else -1,
                height if height is not None else -1,
            )
            presentation.Save()
            return OperationResult(
                message="Audio inserted into PowerPoint slide",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "media_path": str(media),
                    "shape_name": str(shape.Name),
                    "left": float(shape.Left),
                    "top": float(shape.Top),
                    "width": float(shape.Width),
                    "height": float(shape.Height),
                    "link_to_file": link_to_file,
                    "save_with_document": save_with_document,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_audio_playback(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        autoplay: bool | None,
        loop_until_stopped: bool | None,
        pause_animation: bool | None,
        hide_while_not_playing: bool | None,
        stop_after_slides: int | None,
        volume: float | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            play_settings = self._get_play_settings(shape)
            media_format = self._get_media_format(shape)

            if autoplay is not None:
                play_settings.PlayOnEntry = -1 if autoplay else 0
            if loop_until_stopped is not None:
                play_settings.LoopUntilStopped = -1 if loop_until_stopped else 0
            if pause_animation is not None:
                play_settings.PauseAnimation = -1 if pause_animation else 0
            if hide_while_not_playing is not None:
                play_settings.HideWhileNotPlaying = -1 if hide_while_not_playing else 0
            if stop_after_slides is not None:
                play_settings.StopAfterSlides = stop_after_slides
            if volume is not None:
                media_format.Volume = volume

            presentation.Save()
            return OperationResult(
                message="PowerPoint audio playback updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "autoplay": autoplay,
                    "loop_until_stopped": loop_until_stopped,
                    "pause_animation": pause_animation,
                    "hide_while_not_playing": hide_while_not_playing,
                    "stop_after_slides": stop_after_slides,
                    "volume": float(media_format.Volume),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def extract_media_inventory(self, path: str) -> PresentationMediaInventoryResult:
        source = self.resolve_document_path(path)

        with self._open_presentation(source, read_only=True) as presentation:
            items: list[MediaSummary] = []
            for slide_index in range(1, int(presentation.Slides.Count) + 1):
                slide = presentation.Slides(slide_index)
                for shape_index, shape in self._iter_shapes(slide):
                    summary = self._media_summary(shape, slide_index, shape_index)
                    if summary is not None:
                        items.append(summary)

            return PresentationMediaInventoryResult(
                file_path=str(source),
                media_count=len(items),
                items=items,
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def replace_image(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        image_path: str,
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
            shape = self._get_shape(slide, shape_index)
            previous_name = str(shape.Name)
            left = float(shape.Left)
            top = float(shape.Top)
            width = float(shape.Width)
            height = float(shape.Height)
            rotation = float(shape.Rotation)

            new_shape = slide.Shapes.AddPicture(str(image), False, True, left, top, width, height)
            with suppress(Exception):
                new_shape.Rotation = rotation
            with suppress(Exception):
                new_shape.Name = previous_name

            shape.Delete()
            presentation.Save()
            return OperationResult(
                message="PowerPoint image replaced",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "shape_name": str(new_shape.Name),
                    "image_path": str(image),
                    "left": float(new_shape.Left),
                    "top": float(new_shape.Top),
                    "width": float(new_shape.Width),
                    "height": float(new_shape.Height),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def crop_image(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        crop_left: float | None,
        crop_right: float | None,
        crop_top: float | None,
        crop_bottom: float | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            picture_format = self._get_picture_format(shape)
            if crop_left is not None:
                picture_format.CropLeft = crop_left
            if crop_right is not None:
                picture_format.CropRight = crop_right
            if crop_top is not None:
                picture_format.CropTop = crop_top
            if crop_bottom is not None:
                picture_format.CropBottom = crop_bottom

            presentation.Save()
            return OperationResult(
                message="PowerPoint image crop updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "crop_left": float(picture_format.CropLeft),
                    "crop_right": float(picture_format.CropRight),
                    "crop_top": float(picture_format.CropTop),
                    "crop_bottom": float(picture_format.CropBottom),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def reset_image(self, path: str, slide_index: int, shape_index: int, create_backup: bool) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            picture_format = self._get_picture_format(shape)
            picture_format.CropLeft = 0
            picture_format.CropRight = 0
            picture_format.CropTop = 0
            picture_format.CropBottom = 0
            with suppress(Exception):
                picture_format.Brightness = 0.5
            with suppress(Exception):
                picture_format.Contrast = 0.5
            with suppress(Exception):
                picture_format.TransparentBackground = 0

            presentation.Save()
            return OperationResult(
                message="PowerPoint image reset",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "crop_left": float(picture_format.CropLeft),
                    "crop_right": float(picture_format.CropRight),
                    "crop_top": float(picture_format.CropTop),
                    "crop_bottom": float(picture_format.CropBottom),
                    "brightness": float(getattr(picture_format, 'Brightness', 0.5)),
                    "contrast": float(getattr(picture_format, 'Contrast', 0.5)),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_image_transparency(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        value: float,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            # PowerPoint COM does not expose a reliable uniform picture transparency property.
            # Apply fill transparency as a best-effort fallback for compatible picture shapes/fills.
            shape.Fill.Transparency = value
            presentation.Save()
            return OperationResult(
                message="PowerPoint image transparency updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "transparency": value,
                    "note": "Applied via shape fill transparency as a COM best-effort fallback.",
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_image_brightness(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        value: float,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            picture_format = self._get_picture_format(shape)
            picture_format.Brightness = value
            presentation.Save()
            return OperationResult(
                message="PowerPoint image brightness updated",
                file_path=str(source),
                backup_path=backup_path,
                details={"slide_index": slide_index, "shape_index": shape_index, "brightness": float(picture_format.Brightness)},
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_image_contrast(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        value: float,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            picture_format = self._get_picture_format(shape)
            picture_format.Contrast = value
            presentation.Save()
            return OperationResult(
                message="PowerPoint image contrast updated",
                file_path=str(source),
                backup_path=backup_path,
                details={"slide_index": slide_index, "shape_index": shape_index, "contrast": float(picture_format.Contrast)},
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
    def apply_builtin_theme(self, path: str, theme_name: str, create_backup: bool) -> OperationResult:
        source = self.resolve_document_path(path)
        theme_path = self._resolve_builtin_theme_path(theme_name)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            presentation.ApplyTheme(str(theme_path))
            presentation.Save()
            return OperationResult(
                message="PowerPoint built-in theme applied",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "theme_name": theme_name,
                    "theme_path": str(theme_path),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def apply_design_ideas(
        self,
        path: str,
        slide_index: int | None,
        fallback_preset: str | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)
        invoked_command = None
        applied_slides: list[int] = []

        with self._open_presentation_session(source, read_only=False, visible=True, with_window=True) as (powerpoint, presentation):
            if slide_index is not None:
                with suppress(Exception):
                    presentation.Windows(1).View.GotoSlide(slide_index)

            for command_name in DESIGN_IDEAS_COMMANDS:
                with suppress(Exception):
                    powerpoint.CommandBars.ExecuteMso(command_name)
                    invoked_command = command_name
                    break

            if fallback_preset is not None:
                if slide_index is None:
                    for current_index in range(1, int(presentation.Slides.Count) + 1):
                        self._apply_style_preset_to_slide(presentation.Slides(current_index), fallback_preset)
                        applied_slides.append(current_index)
                else:
                    self._apply_style_preset_to_slide(presentation.Slides(slide_index), fallback_preset)
                    applied_slides.append(slide_index)
                presentation.Save()

            return OperationResult(
                message="PowerPoint design ideas workflow executed",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "designer_command": invoked_command,
                    "fallback_preset": fallback_preset,
                    "applied_slide_indexes": applied_slides,
                    "manual_review_recommended": invoked_command is not None,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def apply_theme_variant(self, path: str, master_index: int, variant: str, create_backup: bool) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)
        normalized_variant = normalize_powerpoint_token(variant)

        with self._open_presentation(source, read_only=False) as presentation:
            design = self._get_design(presentation, master_index)
            target_variant = None
            target_variant_index = None
            target_variant_name = None

            variants = list(self._iter_theme_variants(design))
            variant_collection = self._get_design_variants(design)
            if variant_collection is not None:
                for item in variants:
                    candidate_name = normalize_powerpoint_token(item.name or "") if item.name else None
                    if normalized_variant.isdigit() and item.variant_index == int(normalized_variant):
                        target_variant = variant_collection(item.variant_index)
                        target_variant_index = item.variant_index
                        target_variant_name = item.name
                        break
                    if candidate_name and candidate_name == normalized_variant:
                        target_variant = variant_collection(item.variant_index)
                        target_variant_index = item.variant_index
                        target_variant_name = item.name
                        break

            applied_via = None
            if target_variant is not None:
                for applier_name, applier in (
                    ("presentation.ApplyThemeVariant", lambda: presentation.ApplyThemeVariant(target_variant)),
                    ("design.ApplyThemeVariant", lambda: design.ApplyThemeVariant(target_variant)),
                ):
                    try:
                        applier()
                        applied_via = applier_name
                        break
                    except Exception:
                        continue

            if applied_via is None:
                if normalized_variant in STYLE_PRESETS:
                    for slide_position in range(1, int(presentation.Slides.Count) + 1):
                        self._apply_style_preset_to_slide(presentation.Slides(slide_position), normalized_variant)
                    applied_via = "style_preset_fallback"
                    target_variant_name = normalized_variant
                else:
                    supported = ", ".join(sorted({normalize_powerpoint_token(item.name or "") for item in variants if item.name}))
                    fallbacks = ", ".join(sorted(STYLE_PRESETS))
                    raise ValueError(
                        "Theme variant could not be applied via COM. "
                        f"Available COM variants: {supported or 'none'}. Supported fallback presets: {fallbacks}"
                    )

            presentation.Save()
            return OperationResult(
                message="PowerPoint theme variant applied",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "master_index": master_index,
                    "variant": variant,
                    "variant_index": target_variant_index,
                    "variant_name": target_variant_name,
                    "applied_via": applied_via,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def extract_theme(self, path: str, master_index: int | None = None) -> PresentationThemeResult:
        source = self.resolve_document_path(path)

        with self._open_presentation(source, read_only=True) as presentation:
            if master_index is None:
                masters = [
                    self._build_master_theme_summary(index, presentation.Designs(index))
                    for index in range(1, int(presentation.Designs.Count) + 1)
                ]
            else:
                design = self._get_design(presentation, master_index)
                masters = [self._build_master_theme_summary(master_index, design)]

            theme_name = masters[0].theme_name if masters else None
            return PresentationThemeResult(file_path=str(source), theme_name=theme_name, masters=masters)

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_master_background(self, path: str, master_index: int, color: str, create_backup: bool) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            design = self._get_design(presentation, master_index)
            master = design.SlideMaster
            master.Background.Fill.Visible = -1
            master.Background.Fill.Solid()
            master.Background.Fill.ForeColor.RGB = parse_office_color(color)
            self._set_theme_scheme_color(design, "msoThemeBackground1", color)
            presentation.Save()
            return OperationResult(
                message="PowerPoint master background updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "master_index": master_index,
                    "color": color,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_master_fonts(
        self,
        path: str,
        master_index: int,
        title_font_name: str | None,
        title_font_size: float | None,
        body_font_name: str | None,
        body_font_size: float | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)
        updated_shape_count = 0

        with self._open_presentation(source, read_only=False) as presentation:
            design = self._get_design(presentation, master_index)
            for shape in self._iter_master_text_shapes(design):
                if self._shape_is_title_placeholder(shape):
                    if title_font_name is None and title_font_size is None:
                        continue
                    self._apply_shape_text_format(shape, None, None, title_font_name, title_font_size)
                    updated_shape_count += 1
                else:
                    if body_font_name is None and body_font_size is None:
                        continue
                    self._apply_shape_text_format(shape, None, None, body_font_name, body_font_size)
                    updated_shape_count += 1

            presentation.Save()
            summary = self._build_master_theme_summary(master_index, design)
            return OperationResult(
                message="PowerPoint master fonts updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "master_index": master_index,
                    "updated_shape_count": updated_shape_count,
                    "title_font_name": summary.fonts.get("title_font_name"),
                    "body_font_name": summary.fonts.get("body_font_name"),
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_master_colors(
        self,
        path: str,
        master_index: int,
        background_color: str | None,
        title_text_color: str | None,
        body_text_color: str | None,
        accent_color: str | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)
        updated_shape_count = 0

        with self._open_presentation(source, read_only=False) as presentation:
            design = self._get_design(presentation, master_index)
            master = design.SlideMaster

            if background_color is not None:
                master.Background.Fill.Visible = -1
                master.Background.Fill.Solid()
                master.Background.Fill.ForeColor.RGB = parse_office_color(background_color)
                self._set_theme_scheme_color(design, "msoThemeBackground1", background_color)
            if title_text_color is not None:
                self._set_theme_scheme_color(design, "msoThemeText1", title_text_color)
            if body_text_color is not None:
                self._set_theme_scheme_color(design, "msoThemeText2", body_text_color)
            if accent_color is not None:
                self._set_theme_scheme_color(design, "msoThemeAccent1", accent_color)

            for shape in self._iter_master_text_shapes(design):
                if self._shape_is_title_placeholder(shape):
                    if title_text_color is not None:
                        self._apply_text_range_style(
                            self._get_text_range(shape),
                            text=None,
                            font_name=None,
                            font_size=None,
                            bold=None,
                            italic=None,
                            underline=None,
                            color=title_text_color,
                            alignment=None,
                        )
                        updated_shape_count += 1
                    if accent_color is not None:
                        self._apply_shape_fill_and_line(
                            shape,
                            fill_color=accent_color,
                            fill_transparency=None,
                            line_color=accent_color,
                            line_weight=None,
                            line_transparency=None,
                            line_visible=None,
                        )
                elif body_text_color is not None:
                    self._apply_text_range_style(
                        self._get_text_range(shape),
                        text=None,
                        font_name=None,
                        font_size=None,
                        bold=None,
                        italic=None,
                        underline=None,
                        color=body_text_color,
                        alignment=None,
                    )
                    updated_shape_count += 1

            presentation.Save()
            summary = self._build_master_theme_summary(master_index, design)
            return OperationResult(
                message="PowerPoint master colors updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "master_index": master_index,
                    "updated_shape_count": updated_shape_count,
                    "background_color": summary.background_color,
                    "title_text_color": summary.colors.get("title_text_color"),
                    "body_text_color": summary.colors.get("body_text_color"),
                    "accent_color": summary.colors.get("accent_color"),
                },
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
    def set_text_gradient(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        start_color: str,
        end_color: str,
        style: str,
        variant: int,
        text: str | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            text_range = self._get_text_range(shape)
            if text is not None:
                text_range.Text = text

            gradient_fill = self._get_text_gradient_fill(shape)
            self._apply_two_color_gradient(
                gradient_fill,
                start_color=start_color,
                end_color=end_color,
                style=style,
                variant=variant,
            )

            presentation.Save()
            return OperationResult(
                message="PowerPoint text gradient updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "start_color": start_color,
                    "end_color": end_color,
                    "style": normalize_powerpoint_token(style),
                    "variant": variant,
                    "text_updated": text is not None,
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
    def set_shape_shadow(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        visible: bool | None,
        color: str | None,
        transparency: float | None,
        blur: float | None,
        offset_x: float | None,
        offset_y: float | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            shadow = shape.Shadow
            if visible is not None:
                shadow.Visible = -1 if visible else 0
            if color is not None:
                shadow.ForeColor.RGB = parse_office_color(color)
            if transparency is not None:
                shadow.Transparency = transparency
            if blur is not None:
                shadow.Blur = blur
            if offset_x is not None:
                shadow.OffsetX = offset_x
            if offset_y is not None:
                shadow.OffsetY = offset_y

            presentation.Save()
            return OperationResult(
                message="PowerPoint shape shadow updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "visible": visible,
                    "color": color,
                    "transparency": transparency,
                    "blur": blur,
                    "offset_x": offset_x,
                    "offset_y": offset_y,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_shape_glow(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        color: str | None,
        radius: float | None,
        transparency: float | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            glow = shape.Glow
            if color is not None:
                glow.Color.RGB = parse_office_color(color)
            if radius is not None:
                glow.Radius = radius
            if transparency is not None:
                glow.Transparency = transparency

            presentation.Save()
            return OperationResult(
                message="PowerPoint shape glow updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "color": color,
                    "radius": radius,
                    "transparency": transparency,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_shape_reflection(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        preset_type: int | None,
        blur: float | None,
        size: float | None,
        offset: float | None,
        transparency: float | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            reflection = shape.Reflection
            if preset_type is not None:
                reflection.Type = preset_type
            if blur is not None:
                reflection.Blur = blur
            if size is not None:
                reflection.Size = size
            if offset is not None:
                reflection.Offset = offset
            if transparency is not None:
                reflection.Transparency = transparency

            presentation.Save()
            return OperationResult(
                message="PowerPoint shape reflection updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "preset_type": preset_type,
                    "blur": blur,
                    "size": size,
                    "offset": offset,
                    "transparency": transparency,
                },
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_shape_soft_edges(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        radius: float,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            shape.SoftEdge.Radius = radius

            presentation.Save()
            return OperationResult(
                message="PowerPoint shape soft edges updated",
                file_path=str(source),
                backup_path=backup_path,
                details={"slide_index": slide_index, "shape_index": shape_index, "radius": radius},
            )

    @retry(stop=stop_after_attempt(3), wait=wait_fixed(1), reraise=True)
    def set_shape_3d(
        self,
        path: str,
        slide_index: int,
        shape_index: int,
        visible: bool | None,
        depth: float | None,
        rotation_x: float | None,
        rotation_y: float | None,
        rotation_z: float | None,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            shape = self._get_shape(slide, shape_index)
            three_d = shape.ThreeD
            if visible is not None:
                three_d.Visible = -1 if visible else 0
            if depth is not None:
                three_d.Depth = depth
            if rotation_x is not None:
                three_d.RotationX = rotation_x
            if rotation_y is not None:
                three_d.RotationY = rotation_y
            if rotation_z is not None:
                three_d.RotationZ = rotation_z

            presentation.Save()
            return OperationResult(
                message="PowerPoint shape 3D updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "visible": visible,
                    "depth": depth,
                    "rotation_x": rotation_x,
                    "rotation_y": rotation_y,
                    "rotation_z": rotation_z,
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
    def set_slide_background_gradient(
        self,
        path: str,
        slide_index: int,
        start_color: str,
        end_color: str,
        style: str,
        variant: int,
        create_backup: bool,
    ) -> OperationResult:
        source = self.resolve_document_path(path)
        backup_path = self.maybe_create_backup(source, create_backup)

        with self._open_presentation(source, read_only=False) as presentation:
            slide = presentation.Slides(slide_index)
            slide.FollowMasterBackground = 0
            self._apply_two_color_gradient(
                slide.Background.Fill,
                start_color=start_color,
                end_color=end_color,
                style=style,
                variant=variant,
            )

            presentation.Save()
            return OperationResult(
                message="PowerPoint slide background gradient updated",
                file_path=str(source),
                backup_path=backup_path,
                details={
                    "slide_index": slide_index,
                    "start_color": start_color,
                    "end_color": end_color,
                    "style": normalize_powerpoint_token(style),
                    "variant": variant,
                    "follow_master": False,
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
