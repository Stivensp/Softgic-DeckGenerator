from __future__ import annotations

from src.ooxml_engine.model.units import emu_to_inches, emu_to_pt, inches_to_emu, pt_to_emu


def test_emu_to_inches_one_inch():
    assert emu_to_inches(914400) == 1.0


def test_emu_to_pt_one_point():
    assert emu_to_pt(12700) == 1.0


def test_inches_to_emu_roundtrip():
    assert inches_to_emu(2.5) == 2286000
    assert emu_to_inches(inches_to_emu(2.5)) == 2.5


def test_pt_to_emu_roundtrip():
    assert pt_to_emu(18) == 228600
    assert emu_to_pt(pt_to_emu(18)) == 18.0


def test_emu_to_inches_zero():
    assert emu_to_inches(0) == 0.0
