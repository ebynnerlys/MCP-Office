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


class SlideSummary(BaseModel):
    slide_index: int
    title: str | None = None
    shape_count: int


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


class SlideNotesResult(BaseModel):
    file_path: str
    slide_index: int
    notes_text: str = ""


class SlideShapesResult(BaseModel):
    file_path: str
    slide_index: int
    shapes: list[ShapeSummary] = Field(default_factory=list)


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
