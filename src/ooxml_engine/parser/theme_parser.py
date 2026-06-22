"""Lee el theme XML (a:clrScheme + a:fontScheme) de un slide master y
construye un Theme del modelo interno. Generaliza el acceso a tema que
apply_visual_config.py hace de forma incidental dentro de
_resolve_scheme_color — aqui se lee el esquema completo de una vez, no
color por color.
"""
from __future__ import annotations

import logging

import lxml.etree as etree
from pptx.oxml.ns import qn

from src.ooxml_engine.model.document import ColorScheme, Theme

logger = logging.getLogger("softgic.ooxml_engine.theme_parser")

_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"

# slot del modelo -> nombre del elemento OOXML dentro de a:clrScheme
_COLOR_SLOTS: tuple[tuple[str, str], ...] = (
    ("dk1", "dk1"), ("lt1", "lt1"), ("dk2", "dk2"), ("lt2", "lt2"),
    ("accent1", "accent1"), ("accent2", "accent2"), ("accent3", "accent3"),
    ("accent4", "accent4"), ("accent5", "accent5"), ("accent6", "accent6"),
    ("hlink", "hlink"), ("fol_hlink", "folHlink"),
)

_FALLBACK_COLORS = {
    "dk1": "000000", "lt1": "FFFFFF", "dk2": "44546A", "lt2": "E7E6E6",
    "accent1": "4472C4", "accent2": "ED7D31", "accent3": "A5A5A5",
    "accent4": "FFC000", "accent5": "5B9BD5", "accent6": "70AD47",
    "hlink": "0563C1", "fol_hlink": "954F72",
}
_FALLBACK_FONT = "Calibri"


def _read_color_slot(scheme_elem, xml_name: str) -> str | None:
    clr = scheme_elem.find(qn("a:" + xml_name))
    if clr is None:
        return None
    srgb = clr.find(qn("a:srgbClr"))
    if srgb is not None:
        val = srgb.get("val", "")
        return val.upper() if len(val) == 6 else None
    sys_c = clr.find(qn("a:sysClr"))
    if sys_c is not None:
        val = sys_c.get("lastClr", "")
        return val.upper() if len(val) == 6 else None
    return None


def parse_theme(slide_master) -> Theme:
    """slide_master: objeto SlideMaster de python-pptx (prs.slide_masters[0]).
    Nunca lanza — si el theme no se puede leer (XML inesperado, parte
    faltante), retorna un Theme con la paleta default de Office y Calibri,
    dejando un log de aviso en vez de abortar el parseo de toda la
    presentacion por un problema de tema.
    """
    try:
        theme_part = slide_master.part.part_related_by(_RT + "theme")
        theme_xml = etree.fromstring(theme_part.blob)
    except (AttributeError, KeyError, etree.XMLSyntaxError) as exc:
        logger.warning("No se pudo leer el theme part — usando paleta default (%s)", exc)
        return Theme(
            name="default",
            colors=ColorScheme(**_FALLBACK_COLORS),
            fonts_major=_FALLBACK_FONT,
            fonts_minor=_FALLBACK_FONT,
        )

    theme_elem_name = theme_xml.get("name", "default")

    scheme = theme_xml.find(".//" + qn("a:clrScheme"))
    colors_kwargs: dict[str, str] = {}
    for model_field, xml_name in _COLOR_SLOTS:
        value = _read_color_slot(scheme, xml_name) if scheme is not None else None
        colors_kwargs[model_field] = value or _FALLBACK_COLORS[model_field]

    font_scheme = theme_xml.find(".//" + qn("a:fontScheme"))
    fonts_major = _FALLBACK_FONT
    fonts_minor = _FALLBACK_FONT
    if font_scheme is not None:
        major_latin = font_scheme.find(qn("a:majorFont") + "/" + qn("a:latin"))
        minor_latin = font_scheme.find(qn("a:minorFont") + "/" + qn("a:latin"))
        if major_latin is not None and major_latin.get("typeface"):
            fonts_major = major_latin.get("typeface")
        if minor_latin is not None and minor_latin.get("typeface"):
            fonts_minor = minor_latin.get("typeface")

    return Theme(
        name=theme_elem_name,
        colors=ColorScheme(**colors_kwargs),
        fonts_major=fonts_major,
        fonts_minor=fonts_minor,
    )
