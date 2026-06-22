"""
Lee config/generated/layout.yaml y provee posiciones a los renderers.
Si el archivo no existe, los renderers usan sus valores por defecto.
"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

import yaml

logger = logging.getLogger("softgic.layout.layout_config")

_ROOT = Path(__file__).resolve().parent.parent.parent
_CACHE: dict[str, Any] | None = None


def _load() -> dict[str, Any]:
    global _CACHE
    if _CACHE is None:
        p = _ROOT / "config" / "generated" / "layout.yaml"
        if p.exists():
            try:
                loaded = yaml.safe_load(p.read_text(encoding="utf-8"))
                # yaml.safe_load devuelve None si el archivo esta vacio
                _CACHE = loaded if isinstance(loaded, dict) else {}
            except Exception as exc:
                logger.warning("No se pudo leer layout.yaml: %s — usando defaults", exc)
                _CACHE = {}
        else:
            _CACHE = {}
    return _CACHE


def reload() -> None:
    """Fuerza la recarga del archivo (util despues de apply_visual_config)."""
    global _CACHE
    _CACHE = None


def pos(slide_type: str, element: str) -> dict[str, Any] | None:
    """
    Retorna las propiedades del elemento desde layout.yaml.
    - Elemento ausente en layout.yaml         → {} (renderer usa _D por defecto)
    - Elemento marcado {hidden: true}          → None (eliminado del editor visual)
    - Elemento con propiedades almacenadas     → dict con sus propiedades
    - Valor invalido en YAML (no es dict)      → {} (tratar como ausente)
    """
    slide_data = _load().get(slide_type, {})
    if not isinstance(slide_data, dict):
        return {}
    el_data = slide_data.get(element, {})
    if not isinstance(el_data, dict):
        # Valor corrupto (ej: "logo: true" editado a mano) — tratar como ausente
        logger.warning("Valor invalido para '%s.%s' en layout.yaml — ignorado", slide_type, element)
        return {}
    if el_data.get("hidden"):
        return None
    return {k: v for k, v in el_data.items() if k != "hidden"}
