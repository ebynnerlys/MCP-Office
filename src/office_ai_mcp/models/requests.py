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


class PowerPointSlidePositionRequest(PowerPointSlideRequest):
    position: int = Field(ge=1)


class PowerPointSlideNameRequest(PowerPointSlideRequest):
    name: str = Field(min_length=1)


class PowerPointSlideLayoutRequest(PowerPointSlideRequest):
    layout: str = Field(min_length=1)


class PowerPointMasterRequest(DocumentPathRequest):
    master_index: int = Field(ge=1)


class PowerPointOptionalMasterRequest(DocumentPathRequest):
    master_index: int | None = Field(default=None, ge=1)
    create_backup: bool = False


class PowerPointMasterBackgroundRequest(PowerPointMasterRequest):
    color: str = Field(min_length=1)


class PowerPointMasterFontsRequest(PowerPointMasterRequest):
    title_font_name: str | None = None
    title_font_size: float | None = Field(default=None, gt=0)
    body_font_name: str | None = None
    body_font_size: float | None = Field(default=None, gt=0)

    @model_validator(mode="after")
    def validate_changes(self) -> "PowerPointMasterFontsRequest":
        if self.title_font_name is not None and not self.title_font_name.strip():
            raise ValueError("title_font_name cannot be empty")
        if self.body_font_name is not None and not self.body_font_name.strip():
            raise ValueError("body_font_name cannot be empty")
        if not any(
            value is not None
            for value in (
                self.title_font_name,
                self.title_font_size,
                self.body_font_name,
                self.body_font_size,
            )
        ):
            raise ValueError("at least one master font field must be provided")
        return self


class PowerPointMasterColorsRequest(PowerPointMasterRequest):
    background_color: str | None = None
    title_text_color: str | None = None
    body_text_color: str | None = None
    accent_color: str | None = None

    @model_validator(mode="after")
    def validate_changes(self) -> "PowerPointMasterColorsRequest":
        if not any(
            value is not None
            for value in (
                self.background_color,
                self.title_text_color,
                self.body_text_color,
                self.accent_color,
            )
        ):
            raise ValueError("at least one master color field must be provided")
        return self


class PowerPointBuiltinThemeRequest(DocumentPathRequest):
    theme_name: str = Field(min_length=1)

    @field_validator("theme_name")
    @classmethod
    def validate_theme_name(cls, value: str) -> str:
        cleaned = value.strip()
        if not cleaned:
            raise ValueError("theme_name cannot be empty")
        return cleaned


class PowerPointDesignIdeasRequest(DocumentPathRequest):
    slide_index: int | None = Field(default=None, ge=1)
    fallback_preset: str | None = "executive"

    @model_validator(mode="after")
    def validate_fallback(self) -> "PowerPointDesignIdeasRequest":
        if self.fallback_preset is not None and not self.fallback_preset.strip():
            raise ValueError("fallback_preset cannot be empty when provided")
        return self


class PowerPointThemeVariantRequest(PowerPointMasterRequest):
    variant: str = Field(min_length=1)

    @field_validator("variant")
    @classmethod
    def validate_variant(cls, value: str) -> str:
        cleaned = value.strip()
        if not cleaned:
            raise ValueError("variant cannot be empty")
        return cleaned


class PowerPointPlaceholderRequest(PowerPointSlideRequest):
    shape_index: int | None = Field(default=None, ge=1)
    shape_name: str | None = None
    placeholder_type: int | None = Field(default=None, ge=1)
    placeholder_occurrence: int = Field(default=1, ge=1)

    @model_validator(mode="after")
    def validate_placeholder_selector(self) -> "PowerPointPlaceholderRequest":
        selector_count = sum(
            value is not None for value in (self.shape_index, self.shape_name, self.placeholder_type)
        )
        if selector_count != 1:
            raise ValueError("provide exactly one of shape_index, shape_name or placeholder_type")
        if self.shape_name is not None and not self.shape_name.strip():
            raise ValueError("shape_name cannot be empty")
        return self


class PowerPointFillPlaceholderRequest(PowerPointPlaceholderRequest):
    text: str = ""
    text_color: str | None = None
    font_name: str | None = None
    font_size: float | None = Field(default=None, gt=0)


class PowerPointReplacePlaceholderRequest(PowerPointPlaceholderRequest):
    replacement_kind: str = Field(default="textbox", min_length=1)
    text: str | None = None
    image_path: str | None = None
    shape_type: str = Field(default="rectangle", min_length=1)
    fill_color: str | None = None
    line_color: str | None = None
    text_color: str | None = None
    font_name: str | None = None
    font_size: float | None = Field(default=None, gt=0)

    @model_validator(mode="after")
    def validate_replacement(self) -> "PowerPointReplacePlaceholderRequest":
        normalized = self.replacement_kind.strip().lower().replace("-", "_").replace(" ", "_")
        if normalized not in {"textbox", "text", "image", "shape"}:
            raise ValueError("replacement_kind must be one of textbox, text, image or shape")
        if normalized == "image":
            if self.image_path is None or not self.image_path.strip():
                raise ValueError("image_path is required when replacement_kind=image")
        return self


class PowerPointSlideTitleRequest(PowerPointSlideRequest):
    title: str = ""


class PowerPointSearchRequest(DocumentPathRequest):
    query: str
    create_backup: bool = False

    @field_validator("query")
    @classmethod
    def validate_query(cls, value: str) -> str:
        cleaned = value.strip()
        if not cleaned:
            raise ValueError("query cannot be empty")
        return cleaned


class PowerPointDocumentPropertiesRequest(DocumentPathRequest):
    author: str | None = None
    title: str | None = None
    subject: str | None = None
    keywords: str | None = None
    comments: str | None = None
    category: str | None = None
    company: str | None = None
    manager: str | None = None

    @model_validator(mode="after")
    def validate_changes(self) -> "PowerPointDocumentPropertiesRequest":
        if not any(
            value is not None
            for value in (
                self.author,
                self.title,
                self.subject,
                self.keywords,
                self.comments,
                self.category,
                self.company,
                self.manager,
            )
        ):
            raise ValueError("at least one document property must be provided")
        return self


class PowerPointShapeRequest(PowerPointSlideRequest):
    shape_index: int = Field(ge=1)


class PowerPointShapeSelectorRequest(PowerPointSlideRequest):
    shape_index: int | None = Field(default=None, ge=1)
    shape_name: str | None = None

    @model_validator(mode="after")
    def validate_selector(self) -> "PowerPointShapeSelectorRequest":
        if (self.shape_index is None) == (self.shape_name is None):
            raise ValueError("provide exactly one of shape_index or shape_name")
        if self.shape_name is not None and not self.shape_name.strip():
            raise ValueError("shape_name cannot be empty")
        return self


class PowerPointShapeNameRequest(PowerPointShapeRequest):
    name: str = Field(min_length=1)


class PowerPointShapeSearchRequest(PowerPointSlideRequest):
    shape_name_contains: str | None = None
    text_contains: str | None = None
    shape_type: str | None = None
    create_backup: bool = False

    @model_validator(mode="after")
    def validate_filters(self) -> "PowerPointShapeSearchRequest":
        if self.shape_name_contains is not None and not self.shape_name_contains.strip():
            raise ValueError("shape_name_contains cannot be empty")
        if self.text_contains is not None and not self.text_contains.strip():
            raise ValueError("text_contains cannot be empty")
        if self.shape_type is not None and not self.shape_type.strip():
            raise ValueError("shape_type cannot be empty")
        if self.shape_name_contains is None and self.text_contains is None and self.shape_type is None:
            raise ValueError("at least one shape search filter must be provided")
        return self


class PowerPointShapeCollectionRequest(PowerPointSlideRequest):
    shape_indexes: list[int] = Field(min_length=2)

    @model_validator(mode="after")
    def validate_shape_indexes(self) -> "PowerPointShapeCollectionRequest":
        if len(set(self.shape_indexes)) != len(self.shape_indexes):
            raise ValueError("shape_indexes must not contain duplicates")
        return self


class PowerPointShapesAlignRequest(PowerPointShapeCollectionRequest):
    alignment: str = Field(min_length=1)
    relative_to_slide: bool = False


class PowerPointShapesDistributeRequest(PowerPointShapeCollectionRequest):
    direction: str = Field(min_length=1)
    relative_to_slide: bool = False


class PowerPointShapesResizeRequest(PowerPointShapeCollectionRequest):
    mode: str = Field(min_length=1)
    reference_shape_index: int | None = Field(default=None, ge=1)


class PowerPointShapeMergeRequest(PowerPointShapeCollectionRequest):
    mode: str = Field(min_length=1)
    primary_shape_index: int | None = Field(default=None, ge=1)

    @model_validator(mode="after")
    def validate_primary_shape(self) -> "PowerPointShapeMergeRequest":
        if self.primary_shape_index is not None and self.primary_shape_index not in self.shape_indexes:
            raise ValueError("primary_shape_index must be included in shape_indexes")
        return self


class PowerPointShapeRotateRequest(PowerPointShapeRequest):
    rotation: float


class PowerPointShapeFlipRequest(PowerPointShapeRequest):
    direction: str = Field(min_length=1)


class PowerPointShapePositionRequest(PowerPointShapeRequest):
    x: float = Field(ge=0)
    y: float = Field(ge=0)


class PowerPointShapeSizeRequest(PowerPointShapeRequest):
    width: float | None = Field(default=None, gt=0)
    height: float | None = Field(default=None, gt=0)

    @model_validator(mode="after")
    def validate_dimensions(self) -> "PowerPointShapeSizeRequest":
        if self.width is None and self.height is None:
            raise ValueError("at least one of width or height must be provided")
        return self


class PowerPointShapeAspectRatioRequest(PowerPointShapeRequest):
    lock: bool


class PowerPointShapeVisibilityRequest(PowerPointShapeRequest):
    visible: bool


class PowerPointShapeCropRequest(PowerPointShapeRequest):
    pass


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


class PowerPointCreatePresentationRequest(DocumentPathRequest):
    layout: str = Field(default="title", min_length=1)
    title: str | None = None
    body_text: str | None = None
    create_backup: bool = False


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


class PowerPointReplaceImageRequest(PowerPointShapeRequest):
    image_path: str

    @field_validator("image_path")
    @classmethod
    def validate_image_path(cls, value: str) -> str:
        cleaned = value.strip()
        if not cleaned:
            raise ValueError("image_path cannot be empty")
        return cleaned


class PowerPointCropImageRequest(PowerPointShapeRequest):
    crop_left: float | None = None
    crop_right: float | None = None
    crop_top: float | None = None
    crop_bottom: float | None = None

    @model_validator(mode="after")
    def validate_crop_fields(self) -> "PowerPointCropImageRequest":
        if all(value is None for value in (self.crop_left, self.crop_right, self.crop_top, self.crop_bottom)):
            raise ValueError("at least one crop field must be provided")
        return self


class PowerPointImageFormatRequest(PowerPointShapeRequest):
    value: float = Field(ge=0, le=1)


class PowerPointMediaRequest(PowerPointSlideRequest):
    media_path: str
    x: float = Field(default=72.0, ge=0)
    y: float = Field(default=72.0, ge=0)
    width: float | None = Field(default=None, gt=0)
    height: float | None = Field(default=None, gt=0)
    link_to_file: bool = False
    save_with_document: bool = True

    @field_validator("media_path")
    @classmethod
    def validate_media_path(cls, value: str) -> str:
        cleaned = value.strip()
        if not cleaned:
            raise ValueError("media_path cannot be empty")
        return cleaned

    @model_validator(mode="after")
    def validate_linking(self) -> "PowerPointMediaRequest":
        if not self.link_to_file and not self.save_with_document:
            raise ValueError("at least one of link_to_file or save_with_document must be true")
        return self


class PowerPointMediaPlaybackRequest(PowerPointShapeRequest):
    autoplay: bool | None = None
    loop_until_stopped: bool | None = None
    pause_animation: bool | None = None
    hide_while_not_playing: bool | None = None
    stop_after_slides: int | None = Field(default=None, ge=0)
    volume: float | None = Field(default=None, ge=0, le=1)

    @model_validator(mode="after")
    def validate_changes(self) -> "PowerPointMediaPlaybackRequest":
        if all(
            value is None
            for value in (
                self.autoplay,
                self.loop_until_stopped,
                self.pause_animation,
                self.hide_while_not_playing,
                self.stop_after_slides,
                self.volume,
            )
        ):
            raise ValueError("at least one playback field must be provided")
        return self


class PowerPointMediaTrimRequest(PowerPointShapeRequest):
    start_point: int | None = Field(default=None, ge=0)
    end_point: int | None = Field(default=None, ge=0)

    @model_validator(mode="after")
    def validate_trim_points(self) -> "PowerPointMediaTrimRequest":
        if self.start_point is None and self.end_point is None:
            raise ValueError("at least one of start_point or end_point must be provided")
        if self.start_point is not None and self.end_point is not None and self.start_point >= self.end_point:
            raise ValueError("start_point must be less than end_point")
        return self


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


class PowerPointTextGradientRequest(PowerPointShapeRequest):
    start_color: str = Field(min_length=1)
    end_color: str = Field(min_length=1)
    style: str = Field(default="horizontal", min_length=1)
    variant: int = Field(default=1, ge=1, le=4)
    text: str | None = None


class PowerPointShapeTextRunsRequest(PowerPointShapeRequest):
    pass


class PowerPointTextRangeStyleRequest(PowerPointShapeRequest):
    start: int = Field(ge=1)
    length: int = Field(ge=1)
    text: str | None = None
    font_name: str | None = None
    font_size: float | None = Field(default=None, gt=0)
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    color: str | None = None
    alignment: str | None = None

    @model_validator(mode="after")
    def validate_changes(self) -> "PowerPointTextRangeStyleRequest":
        if all(
            value is None
            for value in (
                self.text,
                self.font_name,
                self.font_size,
                self.bold,
                self.italic,
                self.underline,
                self.color,
                self.alignment,
            )
        ):
            raise ValueError("at least one text range field must be provided")
        return self


class PowerPointInsertBulletsRequest(PowerPointShapeRequest):
    items: list[str] = Field(min_length=1)
    level: int = Field(default=1, ge=1)
    append: bool = False

    @field_validator("items")
    @classmethod
    def validate_items(cls, value: list[str]) -> list[str]:
        cleaned = [item.strip() for item in value]
        if not all(cleaned):
            raise ValueError("bullet items cannot be empty")
        return cleaned


class PowerPointBulletStyleRequest(PowerPointShapeRequest):
    paragraph_index: int | None = Field(default=None, ge=1)
    visible: bool | None = None
    level: int | None = Field(default=None, ge=1)
    bullet_character: str | None = None
    font_name: str | None = None
    color: str | None = None
    relative_size: float | None = Field(default=None, gt=0)
    left_margin: float | None = None
    first_margin: float | None = None

    @field_validator("bullet_character")
    @classmethod
    def validate_bullet_character(cls, value: str | None) -> str | None:
        if value is None:
            return None
        if len(value) != 1:
            raise ValueError("bullet_character must be a single character")
        return value

    @model_validator(mode="after")
    def validate_changes(self) -> "PowerPointBulletStyleRequest":
        if all(
            value is None
            for value in (
                self.visible,
                self.level,
                self.bullet_character,
                self.font_name,
                self.color,
                self.relative_size,
                self.left_margin,
                self.first_margin,
            )
        ):
            raise ValueError("at least one bullet style field must be provided")
        return self


class PowerPointParagraphSpacingRequest(PowerPointShapeRequest):
    paragraph_index: int | None = Field(default=None, ge=1)
    space_before: float | None = Field(default=None, ge=0)
    space_after: float | None = Field(default=None, ge=0)
    space_within: float | None = Field(default=None, ge=0)

    @model_validator(mode="after")
    def validate_changes(self) -> "PowerPointParagraphSpacingRequest":
        if self.space_before is None and self.space_after is None and self.space_within is None:
            raise ValueError("at least one paragraph spacing field must be provided")
        return self


class PowerPointTextboxMarginsRequest(PowerPointShapeRequest):
    margin_left: float | None = Field(default=None, ge=0)
    margin_right: float | None = Field(default=None, ge=0)
    margin_top: float | None = Field(default=None, ge=0)
    margin_bottom: float | None = Field(default=None, ge=0)

    @model_validator(mode="after")
    def validate_changes(self) -> "PowerPointTextboxMarginsRequest":
        if (
            self.margin_left is None
            and self.margin_right is None
            and self.margin_top is None
            and self.margin_bottom is None
        ):
            raise ValueError("at least one textbox margin field must be provided")
        return self


class PowerPointTextDirectionRequest(PowerPointShapeRequest):
    direction: str = Field(min_length=1)


class PowerPointAutofitRequest(PowerPointShapeRequest):
    mode: str = Field(default="shape_to_fit_text", min_length=1)
    word_wrap: bool | None = None


class PowerPointProofingLanguageRequest(PowerPointShapeRequest):
    language: str = Field(min_length=1)


class PowerPointSpellcheckSlideRequest(PowerPointSlideRequest):
    include_notes: bool = False


class PowerPointSpellcheckPresentationRequest(DocumentPathRequest):
    include_notes: bool = False


class PowerPointTranslateTextRequest(PowerPointShapeRequest):
    target_language: str = Field(min_length=1)
    source_language: str | None = None


class PowerPointShapeFillRequest(PowerPointShapeRequest):
    color: str = Field(min_length=1)
    transparency: float | None = Field(default=None, ge=0, le=1)


class PowerPointShapeLineRequest(PowerPointShapeRequest):
    color: str | None = None
    weight: float | None = Field(default=None, gt=0)
    transparency: float | None = Field(default=None, ge=0, le=1)
    visible: bool | None = None


class PowerPointShapeShadowRequest(PowerPointShapeRequest):
    visible: bool | None = None
    color: str | None = None
    transparency: float | None = Field(default=None, ge=0, le=1)
    blur: float | None = Field(default=None, ge=0)
    offset_x: float | None = None
    offset_y: float | None = None

    @model_validator(mode="after")
    def validate_changes(self) -> "PowerPointShapeShadowRequest":
        if all(
            value is None
            for value in (
                self.visible,
                self.color,
                self.transparency,
                self.blur,
                self.offset_x,
                self.offset_y,
            )
        ):
            raise ValueError("at least one shadow field must be provided")
        return self


class PowerPointShapeGlowRequest(PowerPointShapeRequest):
    color: str | None = None
    radius: float | None = Field(default=None, ge=0)
    transparency: float | None = Field(default=None, ge=0, le=1)

    @model_validator(mode="after")
    def validate_changes(self) -> "PowerPointShapeGlowRequest":
        if self.color is None and self.radius is None and self.transparency is None:
            raise ValueError("at least one glow field must be provided")
        return self


class PowerPointShapeReflectionRequest(PowerPointShapeRequest):
    preset_type: int | None = Field(default=None, ge=0)
    blur: float | None = Field(default=None, ge=0)
    size: float | None = Field(default=None, ge=0)
    offset: float | None = None
    transparency: float | None = Field(default=None, ge=0, le=1)

    @model_validator(mode="after")
    def validate_changes(self) -> "PowerPointShapeReflectionRequest":
        if all(
            value is None
            for value in (self.preset_type, self.blur, self.size, self.offset, self.transparency)
        ):
            raise ValueError("at least one reflection field must be provided")
        return self


class PowerPointShapeSoftEdgesRequest(PowerPointShapeRequest):
    radius: float = Field(ge=0)


class PowerPointShape3DRequest(PowerPointShapeRequest):
    visible: bool | None = None
    depth: float | None = Field(default=None, ge=0)
    rotation_x: float | None = None
    rotation_y: float | None = None
    rotation_z: float | None = None

    @model_validator(mode="after")
    def validate_changes(self) -> "PowerPointShape3DRequest":
        if all(value is None for value in (self.visible, self.depth, self.rotation_x, self.rotation_y, self.rotation_z)):
            raise ValueError("at least one 3D field must be provided")
        return self


class PowerPointBackgroundRequest(PowerPointSlideRequest):
    color: str = Field(min_length=1)
    follow_master: bool = False


class PowerPointBackgroundGradientRequest(PowerPointSlideRequest):
    start_color: str = Field(min_length=1)
    end_color: str = Field(min_length=1)
    style: str = Field(default="horizontal", min_length=1)
    variant: int = Field(default=1, ge=1, le=4)


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


class PowerPointTableRowRequest(PowerPointShapeRequest):
    row_index: int | None = Field(default=None, ge=1)


class PowerPointTableColumnRequest(PowerPointShapeRequest):
    column_index: int | None = Field(default=None, ge=1)


class PowerPointTableMergeCellsRequest(PowerPointShapeRequest):
    row_index: int = Field(ge=1)
    column_index: int = Field(ge=1)
    merge_to_row_index: int = Field(ge=1)
    merge_to_column_index: int = Field(ge=1)


class PowerPointTableSplitCellsRequest(PowerPointShapeRequest):
    row_index: int = Field(ge=1)
    column_index: int = Field(ge=1)
    num_rows: int = Field(ge=1)
    num_columns: int = Field(ge=1)

    @model_validator(mode="after")
    def validate_split_target(self) -> "PowerPointTableSplitCellsRequest":
        if self.num_rows == 1 and self.num_columns == 1:
            raise ValueError("splitting into 1 row and 1 column would not change the table")
        return self


class PowerPointTableStyleRequest(PowerPointShapeRequest):
    style_id: str | None = None
    save_formatting: bool = True
    first_row: bool | None = None
    first_col: bool | None = None
    last_row: bool | None = None
    last_col: bool | None = None
    horiz_banding: bool | None = None
    vert_banding: bool | None = None

    @model_validator(mode="after")
    def validate_changes(self) -> "PowerPointTableStyleRequest":
        if self.style_id is not None and not self.style_id.strip():
            raise ValueError("style_id cannot be empty")
        if all(
            value is None
            for value in (
                self.style_id,
                self.first_row,
                self.first_col,
                self.last_row,
                self.last_col,
                self.horiz_banding,
                self.vert_banding,
            )
        ):
            raise ValueError("at least one table style field must be provided")
        return self


class PowerPointTableFormatRequest(PowerPointShapeRequest):
    text: str | None = None
    font_name: str | None = None
    font_size: float | None = Field(default=None, gt=0)
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    color: str | None = None
    alignment: str | None = None
    fill_color: str | None = None
    fill_transparency: float | None = Field(default=None, ge=0, le=1)
    line_color: str | None = None
    line_weight: float | None = Field(default=None, gt=0)
    line_transparency: float | None = Field(default=None, ge=0, le=1)
    line_visible: bool | None = None

    @model_validator(mode="after")
    def validate_changes(self) -> "PowerPointTableFormatRequest":
        if all(
            value is None
            for value in (
                self.text,
                self.font_name,
                self.font_size,
                self.bold,
                self.italic,
                self.underline,
                self.color,
                self.alignment,
                self.fill_color,
                self.fill_transparency,
                self.line_color,
                self.line_weight,
                self.line_transparency,
                self.line_visible,
            )
        ):
            raise ValueError("at least one table formatting field must be provided")
        return self


class PowerPointTableCellStyleRequest(PowerPointTableFormatRequest):
    row_index: int = Field(ge=1)
    column_index: int = Field(ge=1)


class PowerPointTableRowStyleRequest(PowerPointTableFormatRequest):
    row_index: int = Field(ge=1)


class PowerPointTableColumnStyleRequest(PowerPointTableFormatRequest):
    column_index: int = Field(ge=1)


class PowerPointTableFromCsvRequest(PowerPointPositionedSlideRequest):
    csv_path: str
    delimiter: str = ","

    @field_validator("csv_path")
    @classmethod
    def validate_csv_path(cls, value: str) -> str:
        cleaned = value.strip()
        if not cleaned:
            raise ValueError("csv_path cannot be empty")
        return cleaned

    @field_validator("delimiter")
    @classmethod
    def validate_delimiter(cls, value: str) -> str:
        if len(value) != 1:
            raise ValueError("delimiter must be a single character")
        return value


class PowerPointTableSortRequest(PowerPointShapeRequest):
    column_index: int = Field(ge=1)
    descending: bool = False
    has_header: bool = False


class PowerPointTableFromExcelRangeRequest(PowerPointPositionedSlideRequest):
    excel_path: str
    sheet: str
    cell_range: str = Field(min_length=1)
    shape_index: int | None = Field(default=None, ge=1)

    @field_validator("excel_path", "sheet")
    @classmethod
    def validate_non_empty_text(cls, value: str) -> str:
        cleaned = value.strip()
        if not cleaned:
            raise ValueError("value cannot be empty")
        return cleaned


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


class PowerPointChartAxisScaleRequest(PowerPointShapeRequest):
    axis_kind: str = Field(default="value", min_length=1)
    minimum_scale: float | int | None = None
    maximum_scale: float | int | None = None
    major_unit: float | int | None = None
    minor_unit: float | int | None = None

    @model_validator(mode="after")
    def validate_scale_changes(self) -> "PowerPointChartAxisScaleRequest":
        if (
            self.minimum_scale is None
            and self.maximum_scale is None
            and self.major_unit is None
            and self.minor_unit is None
        ):
            raise ValueError("at least one chart axis scale field must be provided")
        if (
            self.minimum_scale is not None
            and self.maximum_scale is not None
            and float(self.minimum_scale) >= float(self.maximum_scale)
        ):
            raise ValueError("minimum_scale must be less than maximum_scale")
        return self


class PowerPointChartSeriesOrderRequest(PowerPointShapeRequest):
    series_index: int = Field(ge=1)
    position: int = Field(ge=1)


class PowerPointChartDataLabelsRequest(PowerPointShapeRequest):
    series_index: int | None = Field(default=None, ge=1)
    visible: bool | None = None
    show_value: bool | None = None
    show_category_name: bool | None = None
    show_series_name: bool | None = None
    show_percentage: bool | None = None
    separator: str | None = None
    position: str | None = None

    @model_validator(mode="after")
    def validate_label_changes(self) -> "PowerPointChartDataLabelsRequest":
        if (
            self.visible is None
            and self.show_value is None
            and self.show_category_name is None
            and self.show_series_name is None
            and self.show_percentage is None
            and self.separator is None
            and self.position is None
        ):
            raise ValueError("at least one chart data label field must be provided")
        return self


class PowerPointChartGridlinesRequest(PowerPointShapeRequest):
    axis_kind: str = Field(default="value", min_length=1)
    major: bool | None = None
    minor: bool | None = None

    @model_validator(mode="after")
    def validate_gridline_changes(self) -> "PowerPointChartGridlinesRequest":
        if self.major is None and self.minor is None:
            raise ValueError("at least one chart gridline field must be provided")
        return self


class PowerPointChartColorsRequest(PowerPointShapeRequest):
    series_index: int | None = Field(default=None, ge=1)
    fill_color: str | None = None
    line_color: str | None = None
    series_fill_colors: list[str] = Field(default_factory=list)
    series_line_colors: list[str] = Field(default_factory=list)

    @model_validator(mode="after")
    def validate_chart_colors(self) -> "PowerPointChartColorsRequest":
        if (
            self.fill_color is None
            and self.line_color is None
            and not self.series_fill_colors
            and not self.series_line_colors
        ):
            raise ValueError("at least one chart color field must be provided")
        if self.series_index is not None and (self.series_fill_colors or self.series_line_colors):
            raise ValueError("series_index cannot be combined with series palette colors")
        return self


class PowerPointChartTypeChangeRequest(PowerPointShapeRequest):
    chart_type: str = Field(min_length=1)


class PowerPointChartLinkRequest(PowerPointShapeRequest):
    excel_path: str
    sheet: str
    cell_range: str = Field(min_length=1)
    plot_by: str | None = None

    @field_validator("excel_path", "sheet")
    @classmethod
    def validate_required_text(cls, value: str) -> str:
        cleaned = value.strip()
        if not cleaned:
            raise ValueError("value cannot be empty")
        return cleaned


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


class PowerPointAddChartSeriesRequest(PowerPointShapeRequest):
    name: str = Field(min_length=1)
    values: list[float | int] = Field(min_length=1)
    categories: list[str | int | float] = Field(default_factory=list)

    @model_validator(mode="after")
    def validate_series_values(self) -> "PowerPointAddChartSeriesRequest":
        if self.categories and len(self.categories) != len(self.values):
            raise ValueError("values must contain the same number of items as categories")
        return self


class PowerPointDeleteChartSeriesRequest(PowerPointShapeRequest):
    series_index: int = Field(ge=1)


class PowerPointChartExportDataRequest(PowerPointShapeRequest):
    out_path: str | None = None
    export_format: str = Field(default="json", min_length=1)
    create_backup: bool = False

    @field_validator("out_path")
    @classmethod
    def validate_optional_output_path(cls, value: str | None) -> str | None:
        if value is None:
            return None
        cleaned = value.strip()
        if not cleaned:
            raise ValueError("out_path cannot be empty")
        return cleaned

    @field_validator("export_format")
    @classmethod
    def validate_export_format(cls, value: str) -> str:
        normalized = value.strip().lower()
        if normalized not in {"json", "csv"}:
            raise ValueError("export_format must be one of: csv, json")
        return normalized


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
