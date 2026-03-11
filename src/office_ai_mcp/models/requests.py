from __future__ import annotations

from typing import Any

from pydantic import BaseModel, Field, field_validator, model_validator


class DocumentPathRequest(BaseModel):
    path: str
    create_backup: bool = True

    @field_validator("path")
    @classmethod
    def validate_path(cls, value: str) -> str:
        cleaned = value.strip()
        if not cleaned:
            raise ValueError("path cannot be empty")
        return cleaned


class ReplaceTextRequest(DocumentPathRequest):
    find_text: str
    replace_text: str

    @field_validator("find_text")
    @classmethod
    def validate_find_text(cls, value: str) -> str:
        if not value:
            raise ValueError("find_text cannot be empty")
        return value


class ExportPdfRequest(DocumentPathRequest):
    out_path: str
    create_backup: bool = False


class SaveAsRequest(DocumentPathRequest):
    out_path: str
    create_backup: bool = False


class ExcelRangeRequest(DocumentPathRequest):
    sheet: str
    cell_range: str = Field(min_length=1)


class ExcelWriteRangeRequest(ExcelRangeRequest):
    values: Any


class PowerPointSlideRequest(DocumentPathRequest):
    slide_index: int = Field(ge=1)


class PowerPointShapeRequest(PowerPointSlideRequest):
    shape_index: int = Field(ge=1)


class PowerPointPositionedSlideRequest(PowerPointSlideRequest):
    x: float = Field(default=72.0, ge=0)
    y: float = Field(default=72.0, ge=0)
    width: float = Field(default=360.0, gt=0)
    height: float = Field(default=220.0, gt=0)


class PowerPointPresetRequest(PowerPointSlideRequest):
    preset: str = Field(min_length=1)


class PowerPointTextReplaceRequest(PowerPointSlideRequest):
    find_text: str
    replace_text: str

    @field_validator("find_text")
    @classmethod
    def validate_find_text(cls, value: str) -> str:
        if not value:
            raise ValueError("find_text cannot be empty")
        return value


class PowerPointNotesRequest(PowerPointSlideRequest):
    text: str = ""
    append: bool = False


class PowerPointAddSlideRequest(DocumentPathRequest):
    layout: str = Field(default="title_and_text", min_length=1)
    position: int | None = Field(default=None, ge=1)
    title: str | None = None
    body_text: str | None = None


class PowerPointImageRequest(PowerPointSlideRequest):
    image_path: str
    x: float = Field(default=72.0, ge=0)
    y: float = Field(default=72.0, ge=0)
    width: float | None = Field(default=None, gt=0)
    height: float | None = Field(default=None, gt=0)

    @field_validator("image_path")
    @classmethod
    def validate_image_path(cls, value: str) -> str:
        cleaned = value.strip()
        if not cleaned:
            raise ValueError("image_path cannot be empty")
        return cleaned


class PowerPointThemeRequest(DocumentPathRequest):
    theme_path: str

    @field_validator("theme_path")
    @classmethod
    def validate_theme_path(cls, value: str) -> str:
        cleaned = value.strip()
        if not cleaned:
            raise ValueError("theme_path cannot be empty")
        return cleaned


class PowerPointExportSlideImagesRequest(DocumentPathRequest):
    out_dir: str
    image_format: str = Field(default="png", min_length=1)
    width: int | None = Field(default=None, ge=1)
    height: int | None = Field(default=None, ge=1)
    create_backup: bool = False

    @model_validator(mode="after")
    def validate_dimensions(self) -> "PowerPointExportSlideImagesRequest":
        if (self.width is None) != (self.height is None):
            raise ValueError("width and height must be provided together")
        return self


class PowerPointTransitionRequest(PowerPointSlideRequest):
    effect: str = Field(default="fade", min_length=1)
    speed: str = Field(default="medium", min_length=1)
    advance_on_click: bool = True
    advance_after_seconds: float | None = Field(default=None, ge=0)


class PowerPointTextStyleRequest(PowerPointShapeRequest):
    text: str | None = None
    font_name: str | None = None
    font_size: float | None = Field(default=None, gt=0)
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    color: str | None = None
    alignment: str | None = None


class PowerPointShapeFillRequest(PowerPointShapeRequest):
    color: str = Field(min_length=1)
    transparency: float | None = Field(default=None, ge=0, le=1)


class PowerPointShapeLineRequest(PowerPointShapeRequest):
    color: str | None = None
    weight: float | None = Field(default=None, gt=0)
    transparency: float | None = Field(default=None, ge=0, le=1)
    visible: bool | None = None


class PowerPointBackgroundRequest(PowerPointSlideRequest):
    color: str = Field(min_length=1)
    follow_master: bool = False


class PowerPointAnimationRequest(PowerPointShapeRequest):
    effect: str = Field(default="fade", min_length=1)
    trigger: str = Field(default="on_click", min_length=1)
    duration_seconds: float | None = Field(default=None, ge=0)
    delay_seconds: float | None = Field(default=None, ge=0)
    target_kind: str = Field(default="shape", min_length=1)
    animation_level: str | None = None
    row_index: int | None = Field(default=None, ge=1)
    column_index: int | None = Field(default=None, ge=1)

    @model_validator(mode="after")
    def validate_animation_target(self) -> "PowerPointAnimationRequest":
        normalized_target = self.target_kind.strip().lower().replace("-", "_").replace(" ", "_")
        normalized_target = {
            "cell": "table_cell",
            "cells": "table_cells",
            "row": "table_row",
            "column": "table_column",
        }.get(normalized_target, normalized_target)

        if normalized_target == "table_cell":
            if self.row_index is None or self.column_index is None:
                raise ValueError("row_index and column_index are required for table_cell animations")
        elif normalized_target == "table_row":
            if self.row_index is None:
                raise ValueError("row_index is required for table_row animations")
            if self.column_index is not None:
                raise ValueError("column_index is not used for table_row animations")
        elif normalized_target == "table_column":
            if self.column_index is None:
                raise ValueError("column_index is required for table_column animations")
            if self.row_index is not None:
                raise ValueError("row_index is not used for table_column animations")
        elif normalized_target in {"shape", "text", "table", "table_cells"}:
            if self.row_index is not None or self.column_index is not None:
                raise ValueError(
                    "row_index and column_index are only valid for table_cell, table_row, and table_column animations"
                )
        else:
            raise ValueError(
                "target_kind must be one of shape, text, table, table_cell, table_row, table_column, or table_cells"
            )

        if self.animation_level is not None:
            normalized_level = self.animation_level.strip().lower().replace("-", "_").replace(" ", "_")
            if normalized_target != "text" and normalized_level not in {"", "none", "0"}:
                raise ValueError("animation_level is only supported for text animations")

        return self


class PowerPointTableCellRequest(PowerPointShapeRequest):
    row_index: int = Field(ge=1)
    column_index: int = Field(ge=1)
    text: str


class PowerPointChartTitleRequest(PowerPointShapeRequest):
    title: str


class PowerPointChartDataRequest(PowerPointShapeRequest):
    chart_type: str | None = None
    title: str | None = None
    categories: list[str | int | float] = Field(default_factory=list)
    series: list[PowerPointChartSeriesInput] = Field(default_factory=list)
    replace_existing: bool = True

    @model_validator(mode="after")
    def validate_chart_data(self) -> "PowerPointChartDataRequest":
        if self.categories:
            expected = len(self.categories)
            for series in self.series:
                if len(series.values) != expected:
                    raise ValueError("each series must contain the same number of values as categories")
        if self.chart_type is None and self.title is None and not self.categories and not self.series:
            raise ValueError("at least one chart update field must be provided")
        return self


class PowerPointChartSeriesStyleRequest(PowerPointShapeRequest):
    series_index: int = Field(ge=1)
    fill_color: str | None = None
    line_color: str | None = None
    show_data_labels: bool | None = None


class PowerPointChartLayoutRequest(PowerPointShapeRequest):
    legend_visible: bool | None = None
    legend_position: str | None = None
    category_axis_title: str | None = None
    value_axis_title: str | None = None

    @model_validator(mode="after")
    def validate_layout_changes(self) -> "PowerPointChartLayoutRequest":
        if (
            self.legend_visible is None
            and self.legend_position is None
            and self.category_axis_title is None
            and self.value_axis_title is None
        ):
            raise ValueError("at least one chart layout field must be provided")
        return self


class PowerPointSmartArtNodeRequest(PowerPointShapeRequest):
    node_index: int = Field(ge=1)
    text: str


class PowerPointAddTableRequest(PowerPointPositionedSlideRequest):
    rows: int = Field(ge=1)
    columns: int = Field(ge=1)
    values: list[list[Any]] = Field(default_factory=list)

    @model_validator(mode="after")
    def validate_values(self) -> "PowerPointAddTableRequest":
        if len(self.values) > self.rows:
            raise ValueError("values cannot contain more rows than the table definition")
        for row in self.values:
            if len(row) > self.columns:
                raise ValueError("values cannot contain more columns than the table definition")
        return self


class PowerPointChartSeriesInput(BaseModel):
    name: str = Field(min_length=1)
    values: list[float | int] = Field(min_length=1)


class PowerPointAddChartRequest(PowerPointPositionedSlideRequest):
    chart_type: str = Field(default="column_clustered", min_length=1)
    title: str | None = None
    categories: list[str | int | float] = Field(default_factory=list)
    series: list[PowerPointChartSeriesInput] = Field(default_factory=list)

    @model_validator(mode="after")
    def validate_chart_data(self) -> "PowerPointAddChartRequest":
        if self.categories and not self.series:
            raise ValueError("categories require at least one series")
        if self.categories:
            expected = len(self.categories)
            for series in self.series:
                if len(series.values) != expected:
                    raise ValueError("each series must contain the same number of values as categories")
        return self


class PowerPointAddSmartArtRequest(PowerPointPositionedSlideRequest):
    layout: str = Field(default="basic_list", min_length=1)
    node_texts: list[str] = Field(default_factory=list)


class PowerPointAddShapeRequest(PowerPointPositionedSlideRequest):
    shape_type: str = Field(default="rectangle", min_length=1)
    text: str | None = None
    fill_color: str | None = None
    line_color: str | None = None
    line_weight: float | None = Field(default=None, gt=0)
    text_color: str | None = None
    font_name: str | None = None
    font_size: float | None = Field(default=None, gt=0)


class PowerPointAddConnectorRequest(PowerPointSlideRequest):
    connector_type: str = Field(default="straight", min_length=1)
    begin_x: float = Field(ge=0)
    begin_y: float = Field(ge=0)
    end_x: float = Field(ge=0)
    end_y: float = Field(ge=0)
    line_color: str | None = None
    line_weight: float | None = Field(default=None, gt=0)


class PowerPointConnectShapesRequest(PowerPointSlideRequest):
    start_shape_index: int = Field(ge=1)
    end_shape_index: int = Field(ge=1)
    begin_connection_site: int = Field(default=1, ge=1)
    end_connection_site: int = Field(default=1, ge=1)
    connector_type: str = Field(default="straight", min_length=1)
    line_color: str | None = None
    line_weight: float | None = Field(default=None, gt=0)
