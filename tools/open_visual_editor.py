"""
Genera assets/config_visual.pptx y lo abre en PowerPoint.

Cada slide muestra el diseno REAL del tipo de layout:
- Fondo del color correcto (navy, blanco, etc.)
- Texto de ejemplo con la fuente, tamano y color reales
- Barras y paneles con sus colores de tema reales

COMO USAR:
  1. Menu -> [3] Herramientas de diseno -> [1] Abrir editor visual
  2. PowerPoint se abre mostrando como quedara cada slide
  3. Mueve, redimensiona o cambia colores/fuentes a tu gusto
  4. Guarda (Ctrl+S) y cierra PowerPoint
  5. Menu -> [3] -> [2] Aplicar cambios  (lee posiciones, tamanoss y colores)
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
THEME_YAML = ROOT / "config" / "theme.yaml"
OUT_PPTX = ROOT / "assets" / "config_visual.pptx"

W, H = 13.33, 7.5
_RECT = 1  # MSO_AUTO_SHAPE_TYPE.RECTANGLE


def _rgb(hex_color: str) -> RGBColor:
    h = hex_color.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _load_theme() -> dict:
    if THEME_YAML.exists():
        return yaml.safe_load(THEME_YAML.read_text(encoding="utf-8")) or {}
    return {}


def _label_color(bg_hex: str) -> RGBColor:
    """White label on dark backgrounds, dark on light."""
    h = bg_hex.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    luminance = 0.299 * r + 0.587 * g + 0.114 * b
    return RGBColor(255, 255, 255) if luminance < 140 else RGBColor(30, 30, 30)


# ──────────────────────────────────────────────────────────────
# Definicion de cada elemento por slide type
# t: "text" | "fill" | "image"
# fill / text elements se stylizan con los valores del tema
# ──────────────────────────────────────────────────────────────
def _build_defs(c: dict, f: dict) -> dict:
    """Construye las definiciones de diseno usando los colores/fuentes del tema."""
    return {
        "cover": {
            "_bg": c["primary"],
            "accent_bar":  {"t": "fill", "fill": c["accent"]},
            "logo":        {"t": "image", "label": "LOGOTIPO"},
            "overlay":     {"t": "fill", "fill": c["primary"]},
            "title":       {"t": "text", "text": "Titulo Principal del Deck", "pt": f["size_heading"] + 4, "color": c["text_light"], "bold": True},
            "subtitle":    {"t": "text", "text": "Subtitulo descriptivo de la propuesta", "pt": f["size_subheading"] - 4, "color": c["accent"]},
            "divider":     {"t": "fill", "fill": c["accent"]},
            "client":      {"t": "text", "text": "CLIENTE ABC S.A.", "pt": f["size_body"], "color": c["text_muted"]},
            "meta":        {"t": "text", "text": "Autor  |  Fecha", "pt": f["size_caption"], "color": c["text_muted"], "align": "right"},
        },
        "section_divider": {
            "_bg": c["secondary"],
            "top_line":       {"t": "fill", "fill": c["accent"]},
            "right_bar":      {"t": "fill", "fill": c["accent"]},
            "section_number": {"t": "text", "text": "01", "pt": f["size_stat"], "color": c["accent"], "bold": True},
            "section_title":  {"t": "text", "text": "Nombre de la Seccion", "pt": f["size_heading"], "color": c["text_light"], "bold": True},
            "divider":        {"t": "fill", "fill": c["accent"]},
            "tagline":        {"t": "text", "text": "Descripcion breve de esta seccion", "pt": f["size_body"], "color": c["text_muted"], "italic": True},
        },
        "content_two_col": {
            "_bg": c["background"],
            "header_bar":        {"t": "fill", "fill": c["primary"]},
            "title":             {"t": "text", "text": "Titulo del Slide de Contenido", "pt": f["size_subheading"], "color": c["text_light"], "bold": True},
            "accent_line":       {"t": "fill", "fill": c["accent"]},
            "left_col_title":    {"t": "text", "text": "Columna Izquierda", "pt": f["size_body"] + 1, "color": c["accent"], "bold": True},
            "left_col_bullets":  {"t": "text", "text": "Punto 1\nPunto 2\nPunto 3", "pt": f["size_body"], "color": c["text_dark"]},
            "divider":           {"t": "fill", "fill": c["divider"]},
            "right_col_title":   {"t": "text", "text": "Columna Derecha", "pt": f["size_body"] + 1, "color": c["accent"], "bold": True},
            "right_col_bullets": {"t": "text", "text": "Punto A\nPunto B\nPunto C", "pt": f["size_body"], "color": c["text_dark"]},
        },
        "content_one_col": {
            "_bg": c["background"],
            "header_bar":  {"t": "fill", "fill": c["primary"]},
            "title":       {"t": "text", "text": "Titulo del Slide", "pt": f["size_subheading"], "color": c["text_light"], "bold": True},
            "accent_line": {"t": "fill", "fill": c["accent"]},
            "body_title":  {"t": "text", "text": "Encabezado de Contenido", "pt": f["size_body"] + 2, "color": c["primary"], "bold": True},
            "bullets":     {"t": "text", "text": "Primer punto del contenido\nSegundo punto del contenido\nTercer punto del contenido", "pt": f["size_body"], "color": c["text_dark"]},
        },
        "pricing_table": {
            "_bg": c["background"],
            "header_bar":  {"t": "fill", "fill": c["primary"]},
            "title":       {"t": "text", "text": "Titulo de la Tabla de Precios", "pt": f["size_subheading"], "color": c["text_light"], "bold": True},
            "accent_line": {"t": "fill", "fill": c["accent"]},
            "table":       {"t": "fill", "fill": c["divider"], "label": "TABLA DE PRECIOS (Descripcion | Cant | Unidad | P.Unit | Total)"},
            "notes":       {"t": "text", "text": "* Notas adicionales sobre los precios", "pt": f["size_caption"], "color": c["text_muted"], "italic": True},
        },
        "profile_card": {
            "_bg": c["background"],
            "left_panel":      {"t": "fill", "fill": c["primary"]},
            "left_accent_bar": {"t": "fill", "fill": c["accent"]},
            "photo":           {"t": "image", "label": "FOTO\nPERFIL"},
            "name":            {"t": "text", "text": "Nombre del Perfil", "pt": f["size_body"] + 3, "color": c["text_light"], "bold": True},
            "role":            {"t": "text", "text": "Rol Profesional", "pt": f["size_body"], "color": c["accent"]},
            "seniority":       {"t": "text", "text": "Senior  |  8 anos de exp.", "pt": f["size_caption"] + 1, "color": c["text_muted"]},
            "panel_divider":   {"t": "fill", "fill": c["accent"]},
            "skills_label":    {"t": "text", "text": "HABILIDADES TECNICAS", "pt": f["size_caption"], "color": c["accent"], "bold": True},
            "skills_area":     {"t": "fill", "fill": c["secondary"], "label": "[ TAGS DE HABILIDADES ]"},
        },
        "stat_callout": {
            "_bg": c["primary"],
            "accent_bar_top":    {"t": "fill", "fill": c["accent"]},
            "accent_bar_bottom": {"t": "fill", "fill": c["accent"]},
            "big_number":        {"t": "text", "text": "98%", "pt": f["size_stat"], "color": c["accent"], "bold": True, "align": "center"},
            "label":             {"t": "text", "text": "Descripcion del indicador clave", "pt": f["size_subheading"], "color": c["text_light"], "bold": True, "align": "center"},
            "divider":           {"t": "fill", "fill": c["accent"]},
            "context":           {"t": "text", "text": "Contexto adicional del dato estadistico", "pt": f["size_body"], "color": c["text_muted"], "align": "center"},
            "source":            {"t": "text", "text": "Fuente: Area de Datos", "pt": f["size_caption"], "color": c["text_muted"], "align": "right", "italic": True},
        },
        "closing": {
            "_bg": c["primary"],
            "accent_bar_top":    {"t": "fill", "fill": c["accent"]},
            "accent_bar_bottom": {"t": "fill", "fill": c["accent"]},
            "logo":              {"t": "image", "label": "LOGO"},
            "headline":          {"t": "text", "text": "Listo para transformar tu operacion?", "pt": f["size_heading"], "color": c["text_light"], "bold": True, "align": "center"},
            "divider":           {"t": "fill", "fill": c["accent"]},
            "contact_label":     {"t": "text", "text": "CONTACTO", "pt": f["size_caption"], "color": c["accent"], "bold": True, "align": "center"},
            "contact_name":      {"t": "text", "text": "Nombre del Contacto", "pt": f["size_body"], "color": c["text_light"], "bold": True, "align": "center"},
            "contact_email":     {"t": "text", "text": "email@softgic.com", "pt": f["size_body"], "color": c["accent"], "align": "center"},
            "contact_phone":     {"t": "text", "text": "+57 300 000 0000", "pt": f["size_body"], "color": c["text_light"], "align": "center"},
            "cta_box":           {"t": "fill", "fill": c["accent"]},
            "cta_text":          {"t": "text", "text": "Llamado a la accion del cierre", "pt": f["size_body"], "color": c["primary"], "bold": True, "align": "center"},
        },
        "grafico": {
            "_bg": c["background"],
            "header_bar":  {"t": "fill", "fill": c["primary"]},
            "title":       {"t": "text", "text": "Titulo del Grafico", "pt": f["size_subheading"], "color": c["text_light"], "bold": True},
            "accent_line": {"t": "fill", "fill": c["accent"]},
            "chart_area":  {"t": "fill", "fill": c["divider"], "label": "AREA DEL GRAFICO  (chart_type: column / bar / line / pie)"},
            "source":      {"t": "text", "text": "Fuente: Nombre de la fuente", "pt": f["size_caption"], "color": c["text_muted"], "italic": True},
        },
    }


_SLIDE_ORDER = [
    "cover", "section_divider", "content_two_col", "content_one_col",
    "pricing_table", "profile_card", "stat_callout", "closing", "grafico",
]


def _add_label_bar(slide, slide_type: str) -> None:
    """Barra negra en la parte superior con instrucciones."""
    shape = slide.shapes.add_shape(_RECT, Inches(0), Inches(0), Inches(W), Inches(0.32))
    shape.name = "_LABEL_"
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0, 0, 0)
    shape.line.fill.background()
    tf = shape.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = f"LAYOUT: {slide_type}   |   Mueve, redimensiona o cambia colores y fuentes. Guarda (Ctrl+S) y cierra."
    run.font.size = Pt(8)
    run.font.bold = True
    run.font.color.rgb = RGBColor(255, 220, 50)


def _add_fill_element(slide, name: str, pos: dict, el_def: dict) -> None:
    """Rectangulo de color (barra, panel, linea divisora)."""
    shape = slide.shapes.add_shape(
        _RECT,
        Inches(pos["left"]), Inches(pos["top"]),
        Inches(pos["width"]), Inches(pos["height"]),
    )
    shape.name = name

    fill_hex = el_def.get("fill", "CCCCCC").lstrip("#")
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(fill_hex)
    shape.line.fill.background()

    tf = shape.text_frame
    tf.word_wrap = True
    tf.clear()
    label = el_def.get("label", name)
    if pos["height"] >= 0.25:
        para = tf.paragraphs[0]
        para.alignment = PP_ALIGN.CENTER
        run = para.add_run()
        run.text = label
        run.font.size = Pt(7)
        run.font.bold = True
        run.font.color.rgb = _label_color(fill_hex)


def _add_text_element(slide, name: str, pos: dict, el_def: dict, font_family: str) -> None:
    """Caja de texto con contenido de muestra en fuente/tamano/color reales."""
    txBox = slide.shapes.add_textbox(
        Inches(pos["left"]), Inches(pos["top"]),
        Inches(pos["width"]), Inches(pos["height"]),
    )
    txBox.name = name

    align_map = {
        "left":   PP_ALIGN.LEFT,
        "center": PP_ALIGN.CENTER,
        "right":  PP_ALIGN.RIGHT,
    }
    align = align_map.get(el_def.get("align", "left"), PP_ALIGN.LEFT)
    color_hex = el_def.get("color", "000000").lstrip("#")
    font_pt = el_def.get("pt", 12)
    bold = el_def.get("bold", False)
    italic = el_def.get("italic", False)

    tf = txBox.text_frame
    tf.word_wrap = True
    tf.clear()

    sample_text = el_def.get("text", name)
    lines = sample_text.split("\n")
    for i, line in enumerate(lines):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        run = para.add_run()
        run.text = line
        run.font.name = font_family
        run.font.size = Pt(font_pt)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = _rgb(color_hex)


def _add_image_placeholder(slide, name: str, pos: dict, el_def: dict) -> None:
    """Rectangulo gris para logotipos o fotos."""
    shape = slide.shapes.add_shape(
        _RECT,
        Inches(pos["left"]), Inches(pos["top"]),
        Inches(pos["width"]), Inches(pos["height"]),
    )
    shape.name = name
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(170, 170, 170)
    shape.line.fill.background()

    label = el_def.get("label", name.upper())
    tf = shape.text_frame
    tf.clear()
    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.CENTER
    run = para.add_run()
    run.text = label
    run.font.size = Pt(8)
    run.font.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)


def _render_slide(prs: Presentation, slide_type: str, layout_data: dict, defs: dict, font_family: str) -> None:
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)

    slide_def = defs.get(slide_type, {})
    bg_color = slide_def.get("_bg", "FFFFFF").lstrip("#")
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = _rgb(bg_color)

    _add_label_bar(slide, slide_type)

    elements = layout_data.get(slide_type, {})
    for el_name, pos in elements.items():
        if not isinstance(pos, dict) or not all(k in pos for k in ("left", "top", "width", "height")):
            continue

        el_def = slide_def.get(el_name, {"t": "fill", "fill": "CCCCCC"})
        el_type = el_def.get("t", "fill")

        if el_type == "text":
            _add_text_element(slide, el_name, pos, el_def, font_family)
        elif el_type == "image":
            _add_image_placeholder(slide, el_name, pos, el_def)
        else:
            _add_fill_element(slide, el_name, pos, el_def)


def main() -> None:
    if not LAYOUT_YAML.exists():
        print(f"ERROR: No se encontro {LAYOUT_YAML}")
        sys.exit(1)

    layout_data: dict = yaml.safe_load(LAYOUT_YAML.read_text(encoding="utf-8")) or {}
    theme = _load_theme()
    c_raw = theme.get("colors", {})
    f_raw = theme.get("fonts", {})

    c = {
        "primary":    c_raw.get("primary",    "#0A1628").lstrip("#"),
        "secondary":  c_raw.get("secondary",  "#1E3A5F").lstrip("#"),
        "accent":     c_raw.get("accent",     "#00AEEF").lstrip("#"),
        "background": c_raw.get("background", "#FFFFFF").lstrip("#"),
        "text_dark":  c_raw.get("text_dark",  "#0A1628").lstrip("#"),
        "text_light": c_raw.get("text_light", "#FFFFFF").lstrip("#"),
        "text_muted": c_raw.get("text_muted", "#6B7280").lstrip("#"),
        "divider":    c_raw.get("divider",    "#E5E7EB").lstrip("#"),
    }
    f = {
        "family":          f_raw.get("family",         "Calibri"),
        "size_heading":    f_raw.get("size_heading",   36),
        "size_subheading": f_raw.get("size_subheading",24),
        "size_body":       f_raw.get("size_body",      14),
        "size_caption":    f_raw.get("size_caption",   10),
        "size_stat":       f_raw.get("size_stat",      72),
    }

    defs = _build_defs(c, f)

    prs = Presentation()
    prs.slide_width  = Inches(W)
    prs.slide_height = Inches(H)

    for slide_type in _SLIDE_ORDER:
        _render_slide(prs, slide_type, layout_data, defs, f["family"])
        print(f"  OK {slide_type}")

    OUT_PPTX.parent.mkdir(exist_ok=True)
    prs.save(str(OUT_PPTX))

    print()
    print(f"Editor visual guardado en: {OUT_PPTX}")
    print()
    print("Instrucciones:")
    print("  - Mueve y redimensiona los elementos donde quieras")
    print("  - Cambia el color de relleno de las barras y paneles")
    print("  - Cambia el tamano de fuente del texto de ejemplo")
    print("  - Cambia el color del texto de ejemplo")
    print("  - NO cambies el nombre de los elementos (visible en 'Selection Pane')")
    print("  - Guarda con Ctrl+S y cierra PowerPoint")
    print("  - Luego usa 'Aplicar cambios del editor visual' para actualizar el proyecto")

    if os.name == "nt":
        os.startfile(str(OUT_PPTX))  # type: ignore[attr-defined]


if __name__ == "__main__":
    main()
