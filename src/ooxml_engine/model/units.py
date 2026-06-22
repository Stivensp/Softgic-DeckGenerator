"""Conversion de unidades EMU <-> pulgadas/puntos.

El modelo interno (ver elements.py, document.py) almacena SIEMPRE en EMU
(igual que python-pptx internamente) — estas funciones son solo para
presentacion (CLI de inspeccion, reportes de fidelidad), nunca se usan
dentro del parser/render para decidir logica.
"""
from __future__ import annotations

EMU_PER_INCH: int = 914400
EMU_PER_PT: int = 12700


def emu_to_inches(value: int) -> float:
    return round(value / EMU_PER_INCH, 4)


def emu_to_pt(value: int) -> float:
    return round(value / EMU_PER_PT, 2)


def inches_to_emu(value: float) -> int:
    return round(value * EMU_PER_INCH)


def pt_to_emu(value: float) -> int:
    return round(value * EMU_PER_PT)
