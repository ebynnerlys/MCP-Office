from __future__ import annotations

from typing import Any

from pydantic import BaseModel, Field


class OperationResult(BaseModel):
    ok: bool = True
    message: str
    file_path: str
    backup_path: str | None = None
    details: dict[str, Any] = Field(default_factory=dict)


class SectionSummary(BaseModel):
    index: int
    text: str
    style_name: str | None = None


class WordStructureResult(BaseModel):
    file_path: str
    paragraph_count: int
    table_count: int
    comment_count: int
    track_changes_enabled: bool
    headings: list[SectionSummary] = Field(default_factory=list)


class WorksheetSummary(BaseModel):
    name: str
    rows: int
    columns: int


class WorkbookSheetsResult(BaseModel):
    file_path: str
    sheets: list[WorksheetSummary] = Field(default_factory=list)


class ExcelRangeResult(BaseModel):
    file_path: str
    sheet: str
    cell_range: str
    values: Any


class DocumentPropertiesResult(BaseModel):
    file_path: str
    built_in: dict[str, Any] = Field(default_factory=dict)
    custom: dict[str, Any] = Field(default_factory=dict)


class FileLinksResult(BaseModel):
    file_path: str
    links: list[Any] = Field(default_factory=list)


class SlideSummary(BaseModel):
    slide_index: int
    title: str | None = None
    shape_count: int


class SlideMetadataSummary(BaseModel):
    slide_index: int
    slide_id: int | None = None
    name: str | None = None
    title: str | None = None
    shape_count: int = 0
    hidden: bool = False
    layout_id: int | None = None
    layout_index: int | None = None
    layout_name: str | None = None
    section_index: int | None = None
    section_name: str | None = None


class SlideMetadataResult(BaseModel):
    file_path: str
    slide: SlideMetadataSummary


class ShapeSummary(BaseModel):
    shape_index: int
    name: str
    shape_type: int
    placeholder_type: int | None = None
    has_text: bool = False
    text_preview: str | None = None
    left: float
    top: float
    width: float
    height: float
    fill_color: str | None = None
    fill_transparency: float | None = None
    line_color: str | None = None
    line_transparency: float | None = None
    line_weight: float | None = None
    text_color: str | None = None
    font_name: str | None = None
    font_size: float | None = None
    font_bold: bool | None = None
    font_italic: bool | None = None


class PresentationSummary(BaseModel):
    file_path: str
    slide_count: int
    slides: list[SlideSummary] = Field(default_factory=list)


class SlideTextResult(BaseModel):
    file_path: str
    slide_index: int
    texts: list[str] = Field(default_factory=list)


class TextRunSummary(BaseModel):
    run_index: int
    start: int
    length: int
    text: str
    font_name: str | None = None
    font_size: float | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    color: str | None = None
    language_id: int | None = None


class ShapeTextRunsResult(BaseModel):
    file_path: str
    slide_index: int
    shape_index: int
    run_count: int
    runs: list[TextRunSummary] = Field(default_factory=list)


class TextMatchSummary(BaseModel):
    slide_index: int
    location: str
    occurrences: int = 1
    shape_index: int | None = None
    shape_name: str | None = None
    text_preview: str | None = None


class PresentationTextSearchResult(BaseModel):
    file_path: str
    query: str
    match_count: int
    matches: list[TextMatchSummary] = Field(default_factory=list)


class SpellingIssueSummary(BaseModel):
    word: str
    occurrences: int = 1
    slide_index: int
    location: str
    shape_index: int | None = None
    shape_name: str | None = None


class SlideSpellcheckResult(BaseModel):
    file_path: str
    slide_index: int
    issue_count: int
    issues: list[SpellingIssueSummary] = Field(default_factory=list)


class PresentationSpellcheckResult(BaseModel):
    file_path: str
    issue_count: int
    issues: list[SpellingIssueSummary] = Field(default_factory=list)


class SlideNotesResult(BaseModel):
    file_path: str
    slide_index: int
    notes_text: str = ""


class PresentationNotesResult(BaseModel):
    file_path: str
    slides: list[SlideNotesResult] = Field(default_factory=list)


class SlideShapesResult(BaseModel):
    file_path: str
    slide_index: int
    shapes: list[ShapeSummary] = Field(default_factory=list)


class ShapeSearchResult(BaseModel):
    file_path: str
    slide_index: int
    match_count: int
    shapes: list[ShapeSummary] = Field(default_factory=list)


class PlaceholderSummary(BaseModel):
    shape_index: int
    name: str
    placeholder_type: int | None = None
    has_text: bool = False
    text_preview: str | None = None
    left: float
    top: float
    width: float
    height: float


class SlidePlaceholdersResult(BaseModel):
    file_path: str
    slide_index: int
    placeholders: list[PlaceholderSummary] = Field(default_factory=list)


class LayoutSummary(BaseModel):
    layout_id: int | None = None
    layout_index: int | None = None
    layout_name: str | None = None
    design_index: int | None = None
    master_name: str | None = None


class PresentationLayoutsResult(BaseModel):
    file_path: str
    layouts: list[LayoutSummary] = Field(default_factory=list)


class MasterSummary(BaseModel):
    master_index: int
    master_name: str | None = None
    theme_name: str | None = None
    layout_count: int = 0


class ThemeVariantSummary(BaseModel):
    variant_index: int
    name: str | None = None
    color_scheme_name: str | None = None
    font_scheme_name: str | None = None


class PresentationMastersResult(BaseModel):
    file_path: str
    masters: list[MasterSummary] = Field(default_factory=list)


class MasterThemeSummary(BaseModel):
    master_index: int
    master_name: str | None = None
    theme_name: str | None = None
    background_color: str | None = None
    fonts: dict[str, Any] = Field(default_factory=dict)
    colors: dict[str, Any] = Field(default_factory=dict)
    layouts: list[LayoutSummary] = Field(default_factory=list)
    placeholders: list[PlaceholderSummary] = Field(default_factory=list)
    variants: list[ThemeVariantSummary] = Field(default_factory=list)


class MasterDetailsResult(BaseModel):
    file_path: str
    master: MasterThemeSummary


class PresentationThemeResult(BaseModel):
    file_path: str
    theme_name: str | None = None
    masters: list[MasterThemeSummary] = Field(default_factory=list)


class SlideLayoutResult(BaseModel):
    file_path: str
    slide_index: int
    layout: LayoutSummary


class SlideTransitionResult(BaseModel):
    file_path: str
    slide_index: int
    effect_id: int
    effect_name: str | None = None
    speed_id: int
    speed_name: str | None = None
    advance_on_click: bool
    advance_after_seconds: float | None = None


class AnimationSummary(BaseModel):
    animation_index: int
    shape_name: str | None = None
    effect_id: int | None = None
    effect_name: str | None = None
    trigger_id: int | None = None
    trigger_name: str | None = None
    duration_seconds: float | None = None
    delay_seconds: float | None = None


class SlideAnimationsResult(BaseModel):
    file_path: str
    slide_index: int
    animation_count: int
    animations: list[AnimationSummary] = Field(default_factory=list)


class MediaSummary(BaseModel):
    slide_index: int
    shape_index: int
    shape_name: str
    media_kind: str | None = None
    source_path: str | None = None
    linked: bool | None = None
    left: float
    top: float
    width: float
    height: float
    volume: float | None = None
    trim_start: int | None = None
    trim_end: int | None = None


class PresentationMediaInventoryResult(BaseModel):
    file_path: str
    media_count: int
    items: list[MediaSummary] = Field(default_factory=list)


class TableSummary(BaseModel):
    shape_index: int
    shape_name: str
    rows: int
    columns: int
    cells: list[list[str]] = Field(default_factory=list)


class SlideTablesResult(BaseModel):
    file_path: str
    slide_index: int
    tables: list[TableSummary] = Field(default_factory=list)


class ChartSeriesSummary(BaseModel):
    series_index: int
    name: str | None = None
    values: list[Any] = Field(default_factory=list)
    categories: list[Any] = Field(default_factory=list)


class ChartSummary(BaseModel):
    shape_index: int
    shape_name: str
    chart_type: int | None = None
    chart_title: str | None = None
    legend_visible: bool | None = None
    legend_position_id: int | None = None
    legend_position_name: str | None = None
    category_axis_title: str | None = None
    value_axis_title: str | None = None
    series: list[ChartSeriesSummary] = Field(default_factory=list)


class SlideChartsResult(BaseModel):
    file_path: str
    slide_index: int
    charts: list[ChartSummary] = Field(default_factory=list)


class ChartDataExportResult(BaseModel):
    file_path: str
    slide_index: int
    shape_index: int
    chart_type: int | None = None
    categories: list[Any] = Field(default_factory=list)
    series: list[ChartSeriesSummary] = Field(default_factory=list)
    exported_path: str | None = None
    export_format: str | None = None


class SmartArtNodeSummary(BaseModel):
    node_index: int
    text: str
    level: int | None = None


class SmartArtSummary(BaseModel):
    shape_index: int
    shape_name: str
    node_count: int
    nodes: list[SmartArtNodeSummary] = Field(default_factory=list)


class SlideSmartArtResult(BaseModel):
    file_path: str
    slide_index: int
    smartart_items: list[SmartArtSummary] = Field(default_factory=list)


class ExtendedSlideSummaryResult(BaseModel):
    file_path: str
    slide_index: int
    metadata: SlideMetadataSummary
    texts: list[str] = Field(default_factory=list)
    notes_text: str = ""
    shapes: list[ShapeSummary] = Field(default_factory=list)
    tables: list[TableSummary] = Field(default_factory=list)
    charts: list[ChartSummary] = Field(default_factory=list)
    smartart_items: list[SmartArtSummary] = Field(default_factory=list)
    animations: list[AnimationSummary] = Field(default_factory=list)
