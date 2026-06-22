"""slide (python-pptx) -> OoxmlSlide. Recorre slide.shapes en orden — ese
orden ES el z-order real (fondo -> frente), no se reordena en ningun lado.
"""
from __future__ import annotations

from pptx.oxml.ns import qn

from src.ooxml_engine.model.document import OoxmlSlide
from src.ooxml_engine.model.style import Fill, FillKind
from src.ooxml_engine.parser.shape_parser import parse_shape
from src.ooxml_engine.parser.xml_readers import read_solid_fill_from_parent


def _parse_background(slide) -> Fill:
    """Lee el fondo del slide (Format Background -> Solid Fill en
    PowerPoint). Alcance M1: solo solido — fondo en gradiente se agrega en
    M2 junto con el resto de gradientes (ver style.py).
    """
    try:
        bg_pr = slide._element.find(".//" + qn("p:bgPr"))
        hex_val = read_solid_fill_from_parent(bg_pr, slide)
        if hex_val:
            return Fill(kind=FillKind.SOLID, color_hex=hex_val)
    except AttributeError:
        pass
    return Fill()


def parse_slide(slide, index: int) -> OoxmlSlide:
    return OoxmlSlide(
        index=index,
        name=slide.name if hasattr(slide, "name") and slide.name else f"Slide {index + 1}",
        background=_parse_background(slide),
        elements=[parse_shape(shape) for shape in slide.shapes],
    )
