from __future__ import annotations

from pathlib import Path

import pytest

from src.exceptions import ThemeLoadError
from src.models.theme import ThemeModel
from src.theme.theme_loader import ThemeLoader


VALID_THEME = """\
version: "1.0.0"
colors:
  primary: "#0A1628"
  secondary: "#1E3A5F"
  accent: "#00AEEF"
  background: "#FFFFFF"
  text_dark: "#0A1628"
  text_light: "#FFFFFF"
  text_muted: "#6B7280"
  divider: "#E5E7EB"
fonts:
  family: "Calibri"
  size_heading: 36
  size_subheading: 24
  size_body: 14
  size_caption: 10
  size_stat: 72
  bold_headings: true
assets:
  template: "assets/template.pptx"
  logo_white: "images/logo_white.png"
  logo_dark: "images/logo_dark.png"
  placeholder_profile: "images/placeholder_profile.png"
limits:
  title_max_chars: 80
  subtitle_max_chars: 120
  bullet_max_chars: 100
  bullets_per_column_max: 6
  skills_max: 8
  pricing_rows_max: 10
  profile_image_min_width: 200
  profile_image_min_height: 200
layout_indices:
  cover: 6
  section_divider: 6
  content_two_col: 6
  pricing_table: 6
  profile_card: 6
  stat_callout: 6
  closing: 6
"""


@pytest.fixture
def loader() -> ThemeLoader:
    return ThemeLoader()


def test_load_valid_theme(tmp_path: Path, loader: ThemeLoader) -> None:
    f = tmp_path / "theme.yaml"
    f.write_text(VALID_THEME, encoding="utf-8")
    theme = loader.load(f)
    assert isinstance(theme, ThemeModel)
    assert theme.version == "1.0.0"
    assert theme.colors.primary == "#0A1628"
    assert theme.fonts.family == "Calibri"
    assert theme.limits.title_max_chars == 80


def test_theme_is_frozen(tmp_path: Path, loader: ThemeLoader) -> None:
    f = tmp_path / "theme.yaml"
    f.write_text(VALID_THEME, encoding="utf-8")
    theme = loader.load(f)
    with pytest.raises(Exception):
        theme.version = "mutated"  # type: ignore[misc]


def test_missing_file_raises_error(tmp_path: Path, loader: ThemeLoader) -> None:
    with pytest.raises(ThemeLoadError, match="not found"):
        loader.load(tmp_path / "ghost_theme.yaml")


def test_invalid_yaml_raises_error(tmp_path: Path, loader: ThemeLoader) -> None:
    f = tmp_path / "bad.yaml"
    f.write_text("key: [unclosed", encoding="utf-8")
    with pytest.raises(ThemeLoadError, match="YAML"):
        loader.load(f)


def test_missing_required_key_raises_error(tmp_path: Path, loader: ThemeLoader) -> None:
    f = tmp_path / "incomplete.yaml"
    f.write_text("version: '1.0'\n", encoding="utf-8")
    with pytest.raises(ThemeLoadError, match="validation"):
        loader.load(f)


def test_non_mapping_yaml_raises_error(tmp_path: Path, loader: ThemeLoader) -> None:
    f = tmp_path / "list.yaml"
    f.write_text("- item\n", encoding="utf-8")
    with pytest.raises(ThemeLoadError, match="mapping"):
        loader.load(f)
