from __future__ import annotations

from typing import Any

from pptx.enum.text import PP_ALIGN

from src.layout import layout_config as lc
from src.layout.base_layout import BaseLayoutRenderer
from src.layout.helpers import (
    add_colored_box,
    add_text_box,
    render_layout_extras,
    set_slide_background,
    truncate_text,
)
from src.models.slides import SectionDividerSlide
from src.models.theme import ThemeModel

_D = {
    "top_line":       {"left": 0.0,   "top": 0.0,  "width": 13.33, "height": 0.1 },
    "right_bar":      {"left": 12.53, "top": 0.0,  "width": 0.8,   "height": 7.5 },
    "section_number": {"left": 0.8,   "top": 0.8,  "width": 4.0,   "height": 2.0,  "font_size": 72},
    "section_title":  {"left": 0.8,   "top": 2.8,  "width": 11.3,  "height": 1.5,  "font_size": 36},
    "divider":        {"left": 0.8,   "top": 4.4,  "width": 3.5,   "height": 0.05},
    "tagline":        {"left": 0.8,   "top": 4.6,  "width": 11.3,  "height": 0.6,  "font_size": 14},
}


def _p(el: str) -> dict[str, Any] | None:
    override = lc.pos("section_divider", el)
    if override is None:
        return None
    return {**_D[el], **override}


class SectionDividerRenderer(BaseLayoutRenderer):
    def render(self, slide: object, model: object, theme: ThemeModel) -> None:
        assert isinstance(model, SectionDividerSlide)
        c = theme.colors
        f = theme.fonts
        lim = theme.limits

        set_slide_background(slide, c.background)  # type: ignore[arg-type]

        if (p := _p("top_line")) is not None:
            add_colored_box(slide, p["left"], p["top"], p["width"], p["height"], p.get("fill", c.accent))  # type: ignore[arg-type]

        if (p := _p("right_bar")) is not None:
            add_colored_box(slide, p["left"], p["top"], p["width"], p["height"], p.get("fill", c.accent))  # type: ignore[arg-type]

        if (p := _p("section_number")) is not None:
            add_text_box(slide, p["left"], p["top"], p["width"], p["height"],  # type: ignore[arg-type]
                         model.section_number, f.family, p.get("font_size", f.size_stat),
                         p.get("color", c.accent), bold=p.get("bold", True), align=PP_ALIGN.LEFT)

        section_title = truncate_text(model.section_title, lim.title_max_chars, "section_divider.section_title")
        if (p := _p("section_title")) is not None:
            add_text_box(slide, p["left"], p["top"], p["width"], p["height"],  # type: ignore[arg-type]
                         section_title, f.family, p.get("font_size", f.size_heading),
                         p.get("color", c.text_dark), bold=p.get("bold", f.bold_headings))

        if (p := _p("divider")) is not None:
            add_colored_box(slide, p["left"], p["top"], p["width"], p["height"], p.get("fill", c.accent))  # type: ignore[arg-type]

        if model.tagline and (p := _p("tagline")) is not None:
            add_text_box(slide, p["left"], p["top"], p["width"], p["height"],  # type: ignore[arg-type]
                         model.tagline, f.family, p.get("font_size", f.size_body),
                         p.get("color", c.text_muted), italic=True)

        render_layout_extras(slide, "section_divider", set(_D.keys()), c.accent, model=model)  # type: ignore[arg-type]
