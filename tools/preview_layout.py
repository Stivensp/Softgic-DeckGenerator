"""
Genera output/PREVIEW_layouts.pptx

Abre ese archivo en PowerPoint para ver exactamente donde va cada
elemento en cada tipo de slide. Cada caja muestra su nombre y posicion
en pulgadas (left, top, width, height).

Para cambiar una posicion:
  1. Mide donde quieres el elemento en tu template maestro:
       PowerPoint -> click en el shape -> Format Shape -> Size & Properties -> Position
  2. Anota los valores en cm, dividellos entre 2.54 para convertir a pulgadas
  3. Edita el renderer correspondiente en src/layout/renderers/
  4. Vuelve a correr este script para verificar visualmente
"""
from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parent.parent

# Paleta de colores para cada tipo de elemento
COLORS = {
    "fondo":       "1E3A5F",  # azul oscuro — areas de fondo
    "logo":        "F4A261",  # naranja      — logo
    "titulo":      "2A9D8F",  # verde azulado — titulos
    "texto":       "52B788",  # verde         — texto/bullets
    "decoracion":  "E9C46A",  # amarillo      — barras y lineas
    "imagen":      "E76F51",  # rojo suave    — areas de imagen
    "tabla":       "457B9D",  # azul medio    — tablas
    "destacado":   "9B5DE5",  # morado        — numeros grandes
    "cta":         "F15BB5",  # rosa          — llamadas a accion
}

# Slide: 13.33" x 7.5"
W = 13.33
H = 7.5


def rgb(hex_color: str) -> RGBColor:
    h = hex_color.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def add_box(
    slide,
    left: float, top: float, width: float, height: float,
    color_key: str,
    label: str,
) -> None:
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(left), Inches(top), Inches(width), Inches(height),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(COLORS[color_key])
    shape.line.color.rgb = RGBColor(255, 255, 255)
    shape.line.width = Pt(1)

    tf = shape.text_frame
    tf.word_wrap = True
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.font.size = Pt(8)
    run.font.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)

    pos = f"L:{left}  T:{top}\nW:{width}  H:{height}"
    run.text = f"{label}\n{pos}"


def add_bg(slide, color_key: str) -> None:
    fill = slide.background.fill
    fill.solid()
    hex_map = {
        "primary":    "0A1628",
        "secondary":  "1E3A5F",
        "background": "F0F4F8",
    }
    fill.fore_color.rgb = rgb(hex_map.get(color_key, "CCCCCC"))


def add_slide_title(slide, text: str) -> None:
    shape = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(W), Inches(0.4))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0, 0, 0)
    shape.line.fill.background()
    tf = shape.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = f"LAYOUT: {text.upper()}"
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)


def new_slide(prs: Presentation) -> object:
    layout = prs.slide_layouts[6]  # blank
    return prs.slides.add_slide(layout)


def preview_cover(prs: Presentation) -> None:
    slide = new_slide(prs)
    add_bg(slide, "primary")
    add_slide_title(slide, "cover — portada")

    add_box(slide, 0,    0.4,  W,    0.12, "decoracion", "accent_bar (barra superior)")
    add_box(slide, 0.5,  0.55, 2.2,  0.7,  "logo",       "logo_white")
    add_box(slide, 0,    4.5,  W,    3.0,  "fondo",      "overlay (franja inferior)")
    add_box(slide, 0.8,  2.3,  11.5, 1.5,  "titulo",     "title\nfuente: heading+4 (40pt)")
    add_box(slide, 0.8,  3.85, 9.0,  0.65, "texto",      "subtitle\ncolor: accent")
    add_box(slide, 0.8,  4.55, 2.5,  0.04, "decoracion", "divider line")
    add_box(slide, 0.8,  4.7,  6.0,  0.45, "texto",      "client\ncolor: text_muted")
    add_box(slide, 8.5,  6.6,  4.5,  0.5,  "texto",      "author | date\nalineado derecha")


def preview_section_divider(prs: Presentation) -> None:
    slide = new_slide(prs)
    add_bg(slide, "secondary")
    add_slide_title(slide, "section_divider")

    add_box(slide, 0,     0.4,  W,    0.1,  "decoracion", "linea superior (accent)")
    add_box(slide, 12.53, 0.4,  0.8,  7.1,  "decoracion", "barra vertical derecha (accent)")
    add_box(slide, 0.8,   0.8,  4.0,  2.0,  "destacado",  "section_number\n72pt, accent color")
    add_box(slide, 0.8,   2.8,  11.3, 1.5,  "titulo",     "section_title\n36pt, blanco")
    add_box(slide, 0.8,   4.4,  3.5,  0.05, "decoracion", "divider (accent)")
    add_box(slide, 0.8,   4.6,  11.3, 0.6,  "texto",      "tagline (opcional)\ncursiva, muted")


def preview_content_two_col(prs: Presentation) -> None:
    slide = new_slide(prs)
    add_bg(slide, "background")
    add_slide_title(slide, "content_two_col")

    add_box(slide, 0,    0.4,  W,    1.0,  "fondo",      "header bar (primary)")
    add_box(slide, 0.5,  0.5,  12.3, 0.8,  "titulo",     "title\n24pt, texto blanco")
    add_box(slide, 0.5,  1.55, 5.8,  0.45, "titulo",     "left.title (opcional)\n14pt")
    add_box(slide, 0.5,  2.1,  5.8,  4.5,  "texto",      "left.bullets\n•  item 1\n•  item 2")
    add_box(slide, 6.63, 1.55, 0.07, 5.5,  "decoracion", "divider (accent)")
    add_box(slide, 6.9,  1.55, 5.8,  0.45, "titulo",     "right.title (opcional)\n14pt")
    add_box(slide, 6.9,  2.1,  5.8,  4.5,  "texto",      "right.bullets\n•  item 1\n•  item 2")


def preview_pricing_table(prs: Presentation) -> None:
    slide = new_slide(prs)
    add_bg(slide, "background")
    add_slide_title(slide, "pricing_table")

    add_box(slide, 0,   0.4,  W,    1.0,  "fondo",   "header bar (primary)")
    add_box(slide, 0.5, 0.5,  12.3, 0.8,  "titulo",  "title\n24pt, blanco")
    add_box(slide, 0,   1.4,  W,    0.06, "decoracion", "linea accent")

    # Tabla
    add_box(slide, 0.5, 1.55, 12.33, 0.42, "fondo",    "FILA HEADER (primary)\nDescripcion | Cant | P.Unit | Total")
    add_box(slide, 0.5, 1.97, 12.33, 0.42, "texto",    "fila 1 (background)")
    add_box(slide, 0.5, 2.39, 12.33, 0.42, "decoracion", "fila 2 (divider — gris claro alternado)")
    add_box(slide, 0.5, 2.81, 12.33, 0.42, "texto",    "fila 3 (background)")
    add_box(slide, 0.5, 3.23, 12.33, 0.42, "decoracion", "...")
    add_box(slide, 0.5, 3.65, 12.33, 0.42, "cta",      "FILA TOTAL (accent)\nTOTAL | USD X,XXX.XX")
    add_box(slide, 0.5, 4.2,  12.33, 0.45, "texto",    "notes (opcional)\ncursiva, caption 10pt")


def preview_profile_card(prs: Presentation) -> None:
    slide = new_slide(prs)
    add_bg(slide, "background")
    add_slide_title(slide, "profile_card")

    # Panel izquierdo
    add_box(slide, 0,   0.4,  4.0,  7.1,  "fondo",   "panel izquierdo (secondary)")
    add_box(slide, 0.3, 0.7,  3.4,  3.0,  "imagen",  "photo / placeholder_profile")
    add_box(slide, 0.3, 3.85, 3.4,  0.55, "titulo",  "name\n18pt, blanco, negrita")
    add_box(slide, 0.3, 4.45, 3.4,  0.45, "texto",   "role\n12pt, accent")
    add_box(slide, 0.3, 4.95, 3.4,  0.3,  "texto",   "seniority | X anos\n11pt, muted")

    # Panel derecho
    add_box(slide, 4.3, 0.7,  2.6,  0.45, "titulo",  "Skills\n14pt, primary")
    add_box(slide, 4.3, 1.2,  2.3,  0.35, "cta",     "skill tag (accent)")
    add_box(slide, 6.8, 1.2,  2.3,  0.35, "cta",     "skill tag (accent)")
    add_box(slide, 4.3, 1.65, 2.3,  0.35, "cta",     "skill tag (accent)")
    add_box(slide, 6.8, 1.65, 2.3,  0.35, "cta",     "skill tag (accent)")
    add_box(slide, 4.3, 2.5,  8.5,  0.45, "titulo",  "Highlights\n14pt, primary")
    add_box(slide, 4.3, 3.1,  8.5,  3.5,  "texto",   "highlights (opcional)\n•  logro 1\n•  logro 2")


def preview_stat_callout(prs: Presentation) -> None:
    slide = new_slide(prs)
    add_bg(slide, "primary")
    add_slide_title(slide, "stat_callout")

    add_box(slide, 0,   0.4,  W,    0.12, "decoracion", "accent_bar")
    add_box(slide, 1.5, 1.5,  10.3, 2.5,  "destacado",  "big_number\n72pt, accent, centrado")
    add_box(slide, 1.5, 4.1,  10.3, 1.2,  "titulo",     "label\n36pt, blanco, centrado")
    add_box(slide, 1.5, 5.4,  10.3, 0.55, "texto",      "context (opcional)\nmuted, centrado")
    add_box(slide, 1.5, 6.0,  10.3, 0.4,  "texto",      "source (opcional)\ncaption, muted")


def preview_closing(prs: Presentation) -> None:
    slide = new_slide(prs)
    add_bg(slide, "primary")
    add_slide_title(slide, "closing")

    add_box(slide, 0,   0.4,  W,    0.12, "decoracion", "accent_bar")
    add_box(slide, 0.5, 0.55, 2.2,  0.7,  "logo",       "logo_white")
    add_box(slide, 1.0, 1.5,  11.3, 1.6,  "titulo",     "headline\n40pt, blanco, centrado")
    add_box(slide, 3.0, 3.2,  7.3,  0.65, "cta",        "CTA (opcional)\nacent bg, blanco, centrado")
    add_box(slide, 1.5, 4.2,  10.3, 0.45, "texto",      "contact_name\n14pt, blanco, centrado")
    add_box(slide, 1.5, 4.75, 10.3, 0.4,  "texto",      "contact_email\n12pt, accent, centrado")
    add_box(slide, 1.5, 5.2,  10.3, 0.4,  "texto",      "contact_phone (opcional)\n12pt, muted")


def add_legend(prs: Presentation) -> None:
    slide = new_slide(prs)
    add_bg(slide, "background")
    add_slide_title(slide, "leyenda de colores")

    items = [
        ("logo",      "Logo / imagen de marca"),
        ("titulo",    "Titulos y encabezados"),
        ("texto",     "Texto, bullets, informacion"),
        ("decoracion","Barras decorativas y lineas"),
        ("imagen",    "Areas de foto / imagen"),
        ("tabla",     "Tablas"),
        ("destacado", "Numeros grandes (stat)"),
        ("cta",       "Call to action / skill tags"),
        ("fondo",     "Paneles de fondo"),
    ]
    for i, (key, desc) in enumerate(items):
        row = i % 5
        col = i // 5
        left = 0.5 + col * 6.5
        top  = 1.0 + row * 1.1
        add_box(slide, left, top, 5.8, 0.8, key, desc)

    # Nota de uso
    shape = slide.shapes.add_textbox(Inches(0.5), Inches(6.8), Inches(W - 1), Inches(0.6))
    tf = shape.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.font.size = Pt(9)
    run.font.italic = True
    run.font.color.rgb = RGBColor(80, 80, 80)
    run.text = (
        "Para ajustar posiciones: edita src/layout/renderers/<tipo>.py  |  "
        "Medidas en PowerPoint: click en shape -> Format Shape -> Size & Properties -> Position (cm / 2.54 = pulgadas)"
    )


def preview_content_one_col(prs: Presentation) -> None:
    slide = new_slide(prs)
    add_bg(slide, "background")
    add_slide_title(slide, "content_one_col")

    add_box(slide, 0,   0.4,  W,    1.0,  "fondo",   "header bar (primary)")
    add_box(slide, 0.5, 0.5,  12.3, 0.8,  "titulo",  "title\n24pt, texto blanco")
    add_box(slide, 0,   1.4,  W,    0.06, "decoracion", "linea accent")
    add_box(slide, 0.8, 1.55, 12.0, 0.45, "titulo",  "body_title (opcional)\n16pt, primary")
    add_box(slide, 0.8, 2.1,  12.0, 4.8,  "texto",   "bullets\n•  item 1\n•  item 2\n•  item 3\n(hasta 12 bullets en 1 columna)")


def preview_grafico(prs: Presentation) -> None:
    slide = new_slide(prs)
    add_bg(slide, "background")
    add_slide_title(slide, "grafico")

    add_box(slide, 0,   0.4,  W,    1.0,  "fondo",      "header bar (primary)")
    add_box(slide, 0.5, 0.5,  12.3, 0.8,  "titulo",     "title\n24pt, texto blanco")
    add_box(slide, 0,   1.4,  W,    0.06, "decoracion", "linea accent")
    add_box(slide, 0.5, 1.5,  12.33, 5.4, "tabla",      "chart_area\nchart_type: column / bar / line / pie\ncategories: [ene, feb, mar]\nseries: [{name, values}]")
    add_box(slide, 0.5, 7.0,  12.33, 0.4, "texto",      "source (opcional)\ncursiva, caption")


def main() -> None:
    prs = Presentation()
    prs.slide_width  = Inches(W)
    prs.slide_height = Inches(H)

    print("Generando preview de layouts...\n")

    add_legend(prs)
    preview_cover(prs)
    preview_section_divider(prs)
    preview_content_two_col(prs)
    preview_content_one_col(prs)
    preview_pricing_table(prs)
    preview_profile_card(prs)
    preview_stat_callout(prs)
    preview_closing(prs)
    preview_grafico(prs)

    out = ROOT / "output" / "PREVIEW_layouts.pptx"
    out.parent.mkdir(exist_ok=True)
    prs.save(str(out))

    print(f"Guardado en: {out}")
    print()
    print("Abre ese archivo en PowerPoint para ver donde va cada elemento.")
    print("Para ajustar una posicion:")
    print("  1. Mira el numero que necesitas en el slide del preview")
    print("  2. Abre src/layout/renderers/<tipo>.py")
    print("  3. Busca el elemento (ej: 'logo') y cambia left_in / top_in")
    print("  4. Corre: python generate.py --input examples/propuesta_comercial.yaml --output output/test.pptx")


if __name__ == "__main__":
    main()
