"""
Lee config/layout.yaml y provee posiciones a los renderers.
Si el archivo no existe, los renderers usan sus valores por defecto.
"""
from __future__ import annotations

from pathlib import Path

import yaml

_ROOT = Path(__file__).resolve().parent.parent.parent
_CACHE: dict | None = None


def _load() -> dict:
    global _CACHE
    if _CACHE is None:
        p = _ROOT / "config" / "layout.yaml"
        _CACHE = yaml.safe_load(p.read_text(encoding="utf-8")) if p.exists() else {}
    return _CACHE


def reload() -> None:
    """Fuerza la recarga del archivo (util despues de apply_visual_config)."""
    global _CACHE
    _CACHE = None


def pos(slide_type: str, element: str) -> dict:
    """Retorna el dict {left, top, width, height} de un elemento, o {} si no existe."""
    return _load().get(slide_type, {}).get(element, {})
