"""Dataclasses de estilo compartidas por elementos de texto y formas.

Alcance M1: solo relleno solido (FillKind.SOLID/NONE) y propiedades basicas
de fuente/parrafo. Gradiente, patron, imagen como relleno, linea extendida
y efectos (sombra/glow/reflejo/soft-edge) se agregan en M2 — los valores
de FillKind ya existen en el enum para no romper compatibilidad de campo
cuando se agreguen, pero el parser de M1 nunca produce esos kinds.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum


class FillKind(str, Enum):
    NONE = "none"
    SOLID = "solid"
    GRADIENT_LINEAR = "gradient_linear"
    GRADIENT_RADIAL = "gradient_radial"
    PATTERN = "pattern"
    PICTURE = "picture"


@dataclass(slots=True)
class Fill:
    kind: FillKind = FillKind.NONE
    color_hex: str | None = None  # "RRGGBB", ya resuelto (sin schemeClr crudo)
    alpha: float = 1.0


@dataclass(slots=True)
class FontSpec:
    family: str | None = None
    size_pt: float | None = None
    bold: bool = False
    italic: bool = False
    underline: bool = False
    color_hex: str | None = None


@dataclass(slots=True)
class TextRun:
    text: str
    font: FontSpec = field(default_factory=FontSpec)


@dataclass(slots=True)
class Paragraph:
    runs: list[TextRun] = field(default_factory=list)
    alignment: str = "left"  # "left" | "center" | "right" | "justify"
    bullet_char: str | None = None
