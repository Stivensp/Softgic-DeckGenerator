"""Tests de integracion contra config/template_master.pptx — el unico
archivo real usado para validar el motor OOXML generico. Los conteos de
shapes por slide (187 AUTO_SHAPE + 3 PICTURE, 0 grupos/tablas/charts/
conectores) ya fueron confirmados manualmente inspeccionando el archivo
con python-pptx directamente antes de escribir el parser.
"""
from __future__ import annotations

from collections import Counter

from src.ooxml_engine.model.elements import ElementKind
from src.ooxml_engine.parser.presentation_parser import parse_pptx

_EXPECTED_ELEMENTS_PER_SLIDE = [19, 7, 23, 27, 36, 61, 17]


def test_parse_template_master_no_exceptions(template_master_path):
    doc = parse_pptx(template_master_path)
    assert doc.source_path == str(template_master_path)


def test_parse_template_master_slide_dimensions(template_master_path):
    doc = parse_pptx(template_master_path)

    assert len(doc.slides) == 7
    assert doc.slide_width_emu == 18288000
    assert doc.slide_height_emu == 10287000


def test_parse_template_master_element_counts_per_slide(template_master_path):
    doc = parse_pptx(template_master_path)

    counts = [len(s.elements) for s in doc.slides]
    assert counts == _EXPECTED_ELEMENTS_PER_SLIDE


def test_parse_template_master_no_unknown_elements(template_master_path):
    """template_master.pptx contiene solo AUTO_SHAPE y PICTURE — ningun
    elemento deberia caer en UNKNOWN (eso significaria una regresion en el
    dispatch de shape_parser.py, o que el archivo cambio de contenido)."""
    doc = parse_pptx(template_master_path)

    kinds = Counter(el.kind for s in doc.slides for el in s.elements)

    assert kinds[ElementKind.UNKNOWN] == 0
    assert kinds[ElementKind.IMAGE] == 3
    assert kinds[ElementKind.SHAPE] + kinds[ElementKind.TEXT] == 187


def test_parse_template_master_theme_is_softgic(template_master_path):
    doc = parse_pptx(template_master_path)

    assert doc.theme.fonts_major == "Calibri Light"
    assert doc.theme.fonts_minor == "Calibri"
