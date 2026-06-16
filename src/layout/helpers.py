from __future__ import annotations

import logging
from io import BytesIO
from pathlib import Path
from typing import TYPE_CHECKING

from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

if TYPE_CHECKING:
    from pptx.shapes.autoshape import Shape
    from pptx.slide import Slide
    from pptx.text.text import TextFrame

logger = logging.getLogger("softgic.layout.helpers")

# Integer value for MSO_AUTO_SHAPE_TYPE.RECTANGLE
_RECTANGLE = 1


def hex_to_rgb(hex_color: str) -> RGBColor:
    h = hex_color.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def set_slide_background(slide: "Slide", color_hex: str) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = hex_to_rgb(color_hex)


def add_colored_box(
    slide: "Slide",
    left_in: float,
    top_in: float,
    width_in: float,
    height_in: float,
    fill_hex: str,
    no_line: bool = True,
) -> "Shape":
    shape = slide.shapes.add_shape(
        _RECTANGLE,
        Inches(left_in),
        Inches(top_in),
        Inches(width_in),
        Inches(height_in),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_to_rgb(fill_hex)
    if no_line:
        shape.line.fill.background()
    return shape


def add_text_box(
    slide: "Slide",
    left_in: float,
    top_in: float,
    width_in: float,
    height_in: float,
    text: str,
    font_family: str,
    font_size: int,
    color_hex: str,
    bold: bool = False,
    align: PP_ALIGN = PP_ALIGN.LEFT,
    wrap: bool = True,
    italic: bool = False,
) -> None:
    txBox = slide.shapes.add_textbox(
        Inches(left_in), Inches(top_in), Inches(width_in), Inches(height_in)
    )
    tf = txBox.text_frame
    tf.word_wrap = wrap
    _apply_text(tf, text, font_family, font_size, color_hex, bold, align, italic)


def set_shape_text(
    shape: "Shape",
    text: str,
    font_family: str,
    font_size: int,
    color_hex: str,
    bold: bool = False,
    align: PP_ALIGN = PP_ALIGN.LEFT,
    italic: bool = False,
) -> None:
    tf = shape.text_frame
    tf.word_wrap = True
    _apply_text(tf, text, font_family, font_size, color_hex, bold, align, italic)


def _apply_text(
    tf: "TextFrame",
    text: str,
    font_family: str,
    font_size: int,
    color_hex: str,
    bold: bool,
    align: PP_ALIGN,
    italic: bool,
) -> None:
    tf.clear()
    para = tf.paragraphs[0]
    para.alignment = align
    run = para.add_run()
    run.text = text
    run.font.name = font_family
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = hex_to_rgb(color_hex)


def add_bullet_list(
    slide: "Slide",
    left_in: float,
    top_in: float,
    width_in: float,
    height_in: float,
    items: list[str],
    font_family: str,
    font_size: int,
    color_hex: str,
    bullet_char: str = "•",
) -> None:
    txBox = slide.shapes.add_textbox(
        Inches(left_in), Inches(top_in), Inches(width_in), Inches(height_in)
    )
    tf = txBox.text_frame
    tf.word_wrap = True

    for i, item in enumerate(items):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = PP_ALIGN.LEFT
        run = para.add_run()
        run.text = f"{bullet_char}  {item}"
        run.font.name = font_family
        run.font.size = Pt(font_size)
        run.font.color.rgb = hex_to_rgb(color_hex)


def add_image(
    slide: "Slide",
    image_path: str,
    left_in: float,
    top_in: float,
    width_in: float,
    height_in: float,
    fallback_path: str | None = None,
) -> None:
    path = Path(image_path)
    if not path.exists():
        if fallback_path and Path(fallback_path).exists():
            logger.warning("Image '%s' not found, using placeholder: '%s'", image_path, fallback_path)
            path = Path(fallback_path)
        else:
            logger.warning("Image '%s' not found and no fallback available — skipping", image_path)
            return

    slide.shapes.add_picture(
        str(path),
        Inches(left_in),
        Inches(top_in),
        Inches(width_in),
        Inches(height_in),
    )


def truncate_text(text: str, max_chars: int, field_name: str) -> str:
    if len(text) <= max_chars:
        return text
    truncated = text[: max_chars - 1] + "…"
    logger.warning(
        "Text truncated in field '%s': %d → %d chars",
        field_name,
        len(text),
        len(truncated),
    )
    return truncated
