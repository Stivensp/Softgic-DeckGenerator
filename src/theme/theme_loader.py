from __future__ import annotations

import logging
from pathlib import Path

import yaml
from pydantic import ValidationError

from src.exceptions import ThemeLoadError
from src.models.theme import ThemeModel

logger = logging.getLogger("softgic.theme_loader")


class ThemeLoader:
    def load(self, theme_path: Path) -> ThemeModel:
        logger.debug("Loading theme from: %s", theme_path)

        if not theme_path.exists():
            raise ThemeLoadError(f"Theme file not found: {theme_path}")

        try:
            raw = yaml.safe_load(theme_path.read_text(encoding="utf-8"))
        except yaml.YAMLError as exc:
            raise ThemeLoadError(f"Invalid YAML in theme file '{theme_path}': {exc}") from exc

        if not isinstance(raw, dict):
            raise ThemeLoadError(f"Theme file must be a YAML mapping, got {type(raw).__name__}")

        try:
            theme = ThemeModel.model_validate(raw)
        except ValidationError as exc:
            raise ThemeLoadError(f"Theme validation failed: {exc}") from exc

        logger.info("Theme loaded — version %s", theme.version)
        return theme
