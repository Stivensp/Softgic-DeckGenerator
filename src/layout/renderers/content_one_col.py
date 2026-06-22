from __future__ import annotations

from typing import Any

from src.layout import layout_config as lc
from src.layout.base_layout import BaseLayoutRenderer
from src.layout.helpers import (
    add_bullet_list,
    add_colored_box,
    add_text_box,
    render_layout_extras,
    set_slide_background,
    truncate_text,
)
from src.models.slides import ContentOneColSlide
from src.models.theme import ThemeModel

_D = {
    "header_bar":  {"left": 0.0, "top": 0.0,  "width": 13.33, "height": 1.2 },
    "title":       {"left": 0.5, "top": 0.15, "width": 12.33, "height": 0.9,  "font_size": 24},
    "accent_line": {"left": 0.0, "top": 1.2,  "width": 13.33, "height": 0.06},
    "body_title":  {"left": 0.8, "top": 1.45, "width": 12.0,  "height": 0.55, "font_size": 16},
    "bullets":     {"left": 0.8, "top": 2.1,  "width": 12.0,  "height": 5.0,  "font_size": 14},
}


def _p(el: str) -> dict[str, Any] | None:
    override = lc.pos("content_one_col", el)
    if override is None:
        return None
    return {**_D[el], **override}


class ContentOneColRenderer(BaseLayoutRenderer):
    def render(self, slide: object, model: object, theme: ThemeModel) -> None:
        assert isinstance(model, ContentOneColSlide)
        c = theme.colors
        f = theme.fonts
        lim = theme.limits

        set_slide_background(slide, c.background)  # type: ignore[arg-type]

        if (p := _p("header_bar")) is not None:
            add_colored_box(slide, p["left"], p["top"], p["width"], p["height"], p.get("fill", c.primary))  # type: ignore[arg-type]

        title = truncate_text(model.title, lim.title_max_chars, "content_one_col.title")
        if (p := _p("title")) is not None:
            add_text_box(slide, p["left"], p["top"], p["width"], p["height"], title,  # type: ignore[arg-type]
                         f.family, p.get("font_size", f.size_subheading), p.get("color", c.text_light),
                         bold=p.get("bold", f.bold_headings))

        al = _p("accent_line")
        if al is not None:
            add_colored_box(slide, al["left"], al["top"], al["width"], al["height"], al.get("fill", c.accent))  # type: ignore[arg-type]

        if model.body_title and (p := _p("body_title")) is not None:
            add_text_box(slide, p["left"], p["top"], p["width"], p["height"], model.body_title,  # type: ignore[arg-type]
                         f.family, p.get("font_size", f.size_body + 2), p.get("color", c.primary), bold=p.get("bold", True))

        if model.bullets and (p := _p("bullets")) is not None:
            top = p["top"] if model.body_title else (al["top"] + 0.3 if al is not None else p["top"])
            items = [truncate_text(b, lim.bullet_max_chars, f"bullets[{i}]") for i, b in enumerate(model.bullets)]
            add_bullet_list(slide, p["left"], top, p["width"], p["height"],  # type: ignore[arg-type]
                            items, f.family, p.get("font_size", f.size_body), p.get("color", c.text_dark))

        render_layout_extras(slide, "content_one_col", set(_D.keys()), c.accent, model=model)  # type: ignore[arg-type]
