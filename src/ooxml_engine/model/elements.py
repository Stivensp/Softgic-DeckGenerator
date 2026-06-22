"""Modelo de elementos del shape tree (M1: Shape/Text/Image + UNKNOWN).

Group/Table/Chart/Connector se agregan en M2 — hasta entonces, cualquier
shape de esos tipos se modela como OoxmlElement base con kind=UNKNOWN (ver
shape_parser.py), sin perder posicion/nombre, pero sin contenido detallado.

Orden de campos sin default antes de campos con default se respeta en cada
dataclass (requisito de herencia de dataclasses): OoxmlElement declara sus
4 campos obligatorios primero, is_placeholder/placeholder_type despues con
default — cualquier subclase solo agrega campos con default.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum

from src.ooxml_engine.model.style import Fill, Paragraph


class ElementKind(str, Enum):
    SHAPE = "shape"
    TEXT = "text"
    IMAGE = "image"
    UNKNOWN = "unknown"


@dataclass(slots=True)
class Transform:
    left_emu: int
    top_emu: int
    width_emu: int
    height_emu: int
    rotation_deg: float = 0.0
    flip_h: bool = False
    flip_v: bool = False


@dataclass(slots=True)
class OoxmlElement:
    shape_id: int
    name: str
    kind: ElementKind
    transform: Transform
    is_placeholder: bool = False
    placeholder_type: str | None = None


@dataclass(slots=True)
class TextElement(OoxmlElement):
    paragraphs: list[Paragraph] = field(default_factory=list)
    fill: Fill = field(default_factory=Fill)
    autofit_font_scale: float = 1.0
    margin_left_emu: int = 0
    margin_top_emu: int = 0
    margin_right_emu: int = 0
    margin_bottom_emu: int = 0
    word_wrap: bool = True


@dataclass(slots=True)
class ShapeElement(OoxmlElement):
    geometry_preset: str = "rect"
    fill: Fill = field(default_factory=Fill)
    # Forma con relleno Y texto (ej. boton/CTA solido con etiqueta) — nunca
    # se descarta uno por el otro, a diferencia de la heuristica de
    # apply_visual_config.py (que reclasifica el "kind" segun haya fill o
    # no, porque ahi el destino son 9 campos pydantic fijos). Aqui ambos
    # coexisten porque el modelo no esta atado a esos campos.
    text: TextElement | None = None


@dataclass(slots=True)
class ImageElement(OoxmlElement):
    blob: bytes = b""
    image_format: str = "png"
    crop_left: float = 0.0
    crop_right: float = 0.0
    crop_top: float = 0.0
    crop_bottom: float = 0.0
