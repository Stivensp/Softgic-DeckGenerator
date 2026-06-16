from __future__ import annotations

from pydantic import BaseModel, ConfigDict


class ColorPalette(BaseModel):
    primary: str
    secondary: str
    accent: str
    background: str
    text_dark: str
    text_light: str
    text_muted: str
    divider: str


class FontConfig(BaseModel):
    family: str
    size_heading: int
    size_subheading: int
    size_body: int
    size_caption: int
    size_stat: int
    bold_headings: bool = True


class AssetPaths(BaseModel):
    template: str
    logo_white: str
    logo_dark: str
    placeholder_profile: str


class Limits(BaseModel):
    title_max_chars: int = 80
    subtitle_max_chars: int = 120
    bullet_max_chars: int = 100
    bullets_per_column_max: int = 6
    skills_max: int = 8
    pricing_rows_max: int = 10
    profile_image_min_width: int = 200
    profile_image_min_height: int = 200


class SlideLayoutIndex(BaseModel):
    cover: int
    section_divider: int
    content_two_col: int
    pricing_table: int
    profile_card: int
    stat_callout: int
    closing: int


class ThemeModel(BaseModel):
    model_config = ConfigDict(frozen=True)

    version: str
    colors: ColorPalette
    fonts: FontConfig
    assets: AssetPaths
    limits: Limits
    layout_indices: SlideLayoutIndex
