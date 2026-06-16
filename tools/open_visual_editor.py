"""
Genera assets/config_visual.pptx y lo abre en PowerPoint.

Cada slide = un tipo de layout.
Cada caja = un elemento posicionable (logo, titulo, etc.)

COMO USAR:
  1. Corre este script (Menu -> [3] -> [1])
  2. PowerPoint se abre con las cajas
  3. Mueve y redimensiona las cajas donde quieras
  4. Guarda el archivo (Ctrl+S) y cierra PowerPoint
  5. Menu -> [3] -> [2] para aplicar los cambios
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

import yaml
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parent.parent
LAYOUT_YAML = ROOT / "config" / "layout.yaml"
OUT_PPTX = ROOT / "assets" / "config_visual.pptx"

W, H = 13.33, 7.5

# Color de fondo por tipo de slide
BG_COLORS = {
    "cover":            "0A1628",
    "section_divider":  "1E3A5F",
    "content_two_col":  "F0F4F8",
    "content_one_col":  "F0F4F8",
    "pricing_table":    "F0F4F8",
    "profile_card":     "F0F4F8",
    "stat_callout":     "0A1628",
    "closing":          "0A1628",
    "grafico":          "F0F4F8",
}

# Color de cada caja segun categoria del nombre
def box_color(name: str) -> str:
    if "logo" in name:           return "F4A261"
    if "title" in name or "headline" in name or "number" in name: return "2A9D8F"
    if "bar" in name or "line" in name or "divider" in name:      return "E9C46A"
    if "photo" in name or "image" in name or "panel" in name:     return "E76F51"
    if "table" in name:          return "457B9D"
    if "cta" in name or "tag" in name: return "F15BB5"
    return "52B788"


def rgb(h: str) -> RGBColor:
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def add_element_box(slide, name: str, pos: dict) -> None:
    shape = slide.shapes.add_shape(
        1,
        Inches(pos["left"]), Inches(pos["top"]),
        Inches(pos["width"]), Inches(pos["height"]),
    )
    shape.name = name
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(box_color(name))
    shape.line.color.rgb = RGBColor(255, 255, 255)
    shape.line.width = Pt(1.5)

    tf = shape.text_frame
    tf.word_wrap = True
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.font.size = Pt(7)
    run.font.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    run.text = name


def add_slide_label(slide, slide_type: str) -> None:
    shape = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(W), Inches(0.32))
    shape.name = "_SLIDE_TYPE_"
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0, 0, 0)
    shape.line.fill.background()
    tf = shape.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = f"LAYOUT: {slide_type}   |   Mueve las cajas y guarda (Ctrl+S)"
    run.font.size = Pt(9)
    run.font.bold = True
    run.font.color.rgb = RGBColor(255, 220, 50)


def main() -> None:
    if not LAYOUT_YAML.exists():
        print(f"ERROR: No se encontro {LAYOUT_YAML}")
        sys.exit(1)

    layout_data: dict = yaml.safe_load(LAYOUT_YAML.read_text(encoding="utf-8"))

    prs = Presentation()
    prs.slide_width  = Inches(W)
    prs.slide_height = Inches(H)
    blank = prs.slide_layouts[6]

    slide_order = [
        "cover", "section_divider", "content_two_col", "content_one_col",
        "pricing_table", "profile_card", "stat_callout", "closing", "grafico",
    ]

    for slide_type in slide_order:
        elements = layout_data.get(slide_type, {})
        slide = prs.slides.add_slide(blank)

        # Fondo
        bg = slide.background.fill
        bg.solid()
        bg.fore_color.rgb = rgb(BG_COLORS.get(slide_type, "CCCCCC"))

        add_slide_label(slide, slide_type)

        for name, pos in elements.items():
            if isinstance(pos, dict) and all(k in pos for k in ("left", "top", "width", "height")):
                add_element_box(slide, name, pos)

    OUT_PPTX.parent.mkdir(exist_ok=True)
    prs.save(str(OUT_PPTX))
    print(f"Editor visual guardado en: {OUT_PPTX}")
    print()
    print("Instrucciones:")
    print("  1. Mueve y redimensiona las cajas de colores")
    print("  2. NO cambies el nombre de las cajas (aparece en el panel de seleccion)")
    print("  3. Guarda el archivo con Ctrl+S")
    print("  4. Cierra PowerPoint")
    print("  5. Vuelve al menu y selecciona 'Aplicar cambios del editor visual'")

    if os.name == "nt":
        os.startfile(str(OUT_PPTX))  # type: ignore[attr-defined]


if __name__ == "__main__":
    main()
