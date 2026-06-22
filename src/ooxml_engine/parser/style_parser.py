"""Parseo de relleno (alcance M1: solo solido — gradiente/patron/imagen
se agregan en M2, ver FillKind en model/style.py)."""
from __future__ import annotations

from pptx.enum.dml import MSO_FILL_TYPE

from src.ooxml_engine.model.style import Fill, FillKind
from src.ooxml_engine.parser.xml_readers import read_fill_hex


def parse_fill(shape) -> Fill:
    """Lee el relleno solido de un shape. Cualquier otro tipo (gradiente,
    patron, imagen, fondo heredado) se modela como FillKind.NONE en esta
    fase — no es lo mismo que 'no tiene relleno', es 'este parser de M1
    no extrae ese tipo todavia'. M2 distingue ambos casos.
    """
    try:
        if shape.fill.type == MSO_FILL_TYPE.SOLID:
            hex_val = read_fill_hex(shape)
            if hex_val:
                return Fill(kind=FillKind.SOLID, color_hex=hex_val)
    except (AttributeError, ValueError, NotImplementedError):
        pass
    return Fill()
