"""Raiz del modelo interno: OoxmlPresentation -> OoxmlSlide -> elementos.

Decision deliberada: la herencia de placeholders (slide hereda de layout
hereda de master) se RESUELVE en el momento del parseo — cada OoxmlElement
queda con sus valores ya concretos (lo que verias si abrieras el shape en
PowerPoint), no con una cadena slide_value or layout_value or master_value.

Esto evita que el renderer futuro tenga que reimplementar la logica de
resolucion de PowerPoint (que depende del tipo de placeholder y de si
hereda geometria, texto, o ambos), y hace trivial el comparador de
fidelidad: compara valores finales campo a campo, sin necesidad de saber
de donde vino cada valor. El costo aceptado es que un slide regenerado no
se actualiza automaticamente si cambia el master — no es un requisito de
esta fase (snapshot fiel, no theming dinamico).
"""
from __future__ import annotations

from dataclasses import dataclass, field

from src.ooxml_engine.model.elements import OoxmlElement
from src.ooxml_engine.model.style import Fill


@dataclass(slots=True)
class ColorScheme:
    """Los 12 slots de a:clrScheme, ya resueltos a hex (sysClr resuelto a
    su lastClr)."""
    dk1: str
    lt1: str
    dk2: str
    lt2: str
    accent1: str
    accent2: str
    accent3: str
    accent4: str
    accent5: str
    accent6: str
    hlink: str
    fol_hlink: str


@dataclass(slots=True)
class Theme:
    name: str
    colors: ColorScheme
    fonts_major: str
    fonts_minor: str


@dataclass(slots=True)
class OoxmlSlide:
    """elements en orden de z-order real (fondo -> frente) — es el mismo
    orden de iteracion de slide.shapes, nunca se reordena. index es la
    posicion 0-based dentro de OoxmlPresentation.slides (para reportes
    legibles: 'slide 3, shape "Title 1"')."""
    index: int
    name: str
    background: Fill = field(default_factory=Fill)
    elements: list[OoxmlElement] = field(default_factory=list)


@dataclass(slots=True)
class OoxmlPresentation:
    slide_width_emu: int
    slide_height_emu: int
    theme: Theme
    slides: list[OoxmlSlide] = field(default_factory=list)
    source_path: str | None = None
