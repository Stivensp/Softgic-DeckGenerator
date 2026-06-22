from __future__ import annotations

from typing import Any

from pptx.enum.text import PP_ALIGN

from src.layout import layout_config as lc
from src.layout.base_layout import BaseLayoutRenderer
from src.layout.helpers import (
    add_colored_box,
    add_image,
    add_text_box,
    render_layout_extras,
    set_slide_background,
    truncate_text,
)
from src.models.slides import CoverSlide
from src.models.theme import ThemeModel

_D = {
    "logo":      {"left": 0.5,  "top": 0.25,  "width": 2.2,   "height": 0.7 },
    "accent_bar":{"left": 0.0,  "top": 0.0,   "width": 13.33, "height": 0.12},
    "title":     {"left": 0.8,  "top": 2.3,   "width": 11.5,  "height": 1.5,  "font_size": 40},
    "subtitle":  {"left": 0.8,  "top": 3.85,  "width": 9.0,   "height": 0.65, "font_size": 20},
    "divider":   {"left": 0.8,  "top": 4.55,  "width": 2.5,   "height": 0.04},
    "client":    {"left": 0.8,  "top": 4.7,   "width": 6.0,   "height": 0.45, "font_size": 14},
    "meta":      {"left": 8.5,  "top": 6.6,   "width": 4.5,   "height": 0.5,  "font_size": 10},
}


def _p(el: str) -> dict[str, Any] | None:
    override = lc.pos("cover", el)
    if override is None:
        return None
    return {**_D[el], **override}


class CoverRenderer(BaseLayoutRenderer):
    def render(self, slide: object, model: object, theme: ThemeModel) -> None:
        assert isinstance(model, CoverSlide)
        c = theme.colors
        f = theme.fonts
        lim = theme.limits

        set_slide_background(slide, c.background)  # type: ignore[arg-type]

        if (p := _p("accent_bar")) is not None:
            add_colored_box(slide, p["left"], p["top"], p["width"], p["height"], p.get("fill", c.accent))  # type: ignore[arg-type]

        if (p := _p("logo")) is not None:
            add_image(slide, theme.assets.logo_dark, p["left"], p["top"], p["width"], p["height"])  # type: ignore[arg-type]

        title = truncate_text(model.title, lim.title_max_chars, "cover.title")
        if (p := _p("title")) is not None:
            add_text_box(slide, p["left"], p["top"], p["width"], p["height"], title,  # type: ignore[arg-type]
                         f.family, p.get("font_size", f.size_heading + 4), p.get("color", c.text_dark),
                         bold=p.get("bold", f.bold_headings), align=PP_ALIGN.LEFT)

        if model.subtitle and (p := _p("subtitle")) is not None:
            subtitle = truncate_text(model.subtitle, lim.subtitle_max_chars, "cover.subtitle")
            add_text_box(slide, p["left"], p["top"], p["width"], p["height"], subtitle,  # type: ignore[arg-type]
                         f.family, p.get("font_size", f.size_subheading - 4), p.get("color", c.accent))

        if (p := _p("divider")) is not None:
            add_colored_box(slide, p["left"], p["top"], p["width"], p["height"], p.get("fill", c.accent))  # type: ignore[arg-type]

        if model.client and (p := _p("client")) is not None:
            add_text_box(slide, p["left"], p["top"], p["width"], p["height"], model.client,  # type: ignore[arg-type]
                         f.family, p.get("font_size", f.size_body), p.get("color", c.text_muted))

        meta_parts = [x for x in [model.author, model.date] if x]
        if meta_parts and (p := _p("meta")) is not None:
            add_text_box(slide, p["left"], p["top"], p["width"], p["height"],  # type: ignore[arg-type]
                         "  |  ".join(meta_parts), f.family, p.get("font_size", f.size_caption),
                         p.get("color", c.text_muted), align=PP_ALIGN.RIGHT)

        render_layout_extras(slide, "cover", set(_D.keys()), c.accent, model=model)  # type: ignore[arg-type]
