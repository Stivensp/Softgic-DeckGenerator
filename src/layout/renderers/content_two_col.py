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
from src.models.slides import ContentTwoColSlide
from src.models.theme import ThemeModel

_D = {
    "header_bar":        {"left": 0.0,  "top": 0.0,  "width": 13.33, "height": 1.2 },
    "title":             {"left": 0.5,  "top": 0.15, "width": 12.33, "height": 0.9,  "font_size": 24},
    "accent_line":       {"left": 0.0,  "top": 1.2,  "width": 13.33, "height": 0.06},
    "left_col_title":    {"left": 0.5,  "top": 1.45, "width": 5.9,   "height": 0.55, "font_size": 15},
    "left_col_bullets":  {"left": 0.5,  "top": 2.05, "width": 5.9,   "height": 5.1,  "font_size": 14},
    "divider":           {"left": 6.55, "top": 1.35, "width": 0.04,  "height": 5.9 },
    "right_col_title":   {"left": 7.0,  "top": 1.45, "width": 5.83,  "height": 0.55, "font_size": 15},
    "right_col_bullets": {"left": 7.0,  "top": 2.05, "width": 5.83,  "height": 5.1,  "font_size": 14},
}


def _p(el: str) -> dict[str, Any] | None:
    override = lc.pos("content_two_col", el)
    if override is None:
        return None
    return {**_D[el], **override}


class ContentTwoColRenderer(BaseLayoutRenderer):
    def render(self, slide: object, model: object, theme: ThemeModel) -> None:
        assert isinstance(model, ContentTwoColSlide)
        c = theme.colors
        f = theme.fonts
        lim = theme.limits

        set_slide_background(slide, c.background)  # type: ignore[arg-type]

        if (p := _p("header_bar")) is not None:
            add_colored_box(slide, p["left"], p["top"], p["width"], p["height"], p.get("fill", c.primary))  # type: ignore[arg-type]

        title = truncate_text(model.title, lim.title_max_chars, "content_two_col.title")
        if (p := _p("title")) is not None:
            add_text_box(slide, p["left"], p["top"], p["width"], p["height"], title,  # type: ignore[arg-type]
                         f.family, p.get("font_size", f.size_subheading), p.get("color", c.text_light),
                         bold=p.get("bold", f.bold_headings))

        if (p := _p("accent_line")) is not None:
            add_colored_box(slide, p["left"], p["top"], p["width"], p["height"], p.get("fill", c.accent))  # type: ignore[arg-type]

        # Left column
        lct = _p("left_col_title")
        lcb = _p("left_col_bullets")
        if lcb is not None:
            if model.left.title and lct is not None:
                add_text_box(slide, lct["left"], lct["top"], lct["width"], lct["height"],  # type: ignore[arg-type]
                             model.left.title, f.family, lct.get("font_size", f.size_body + 1),
                             lct.get("color", c.accent), bold=lct.get("bold", True))
            bullets_top = lcb["top"] if model.left.title else (lct["top"] if lct is not None else lcb["top"])
            left_bullets = [truncate_text(b, lim.bullet_max_chars, f"left.bullets[{i}]") for i, b in enumerate(model.left.bullets)]
            add_bullet_list(slide, lcb["left"], bullets_top, lcb["width"], lcb["height"],  # type: ignore[arg-type]
                            left_bullets, f.family, lcb.get("font_size", f.size_body), lcb.get("color", c.text_dark))

        if (p := _p("divider")) is not None:
            add_colored_box(slide, p["left"], p["top"], p["width"], p["height"], p.get("fill", c.divider))  # type: ignore[arg-type]

        # Right column
        rct = _p("right_col_title")
        rcb = _p("right_col_bullets")
        if rcb is not None:
            if model.right.title and rct is not None:
                add_text_box(slide, rct["left"], rct["top"], rct["width"], rct["height"],  # type: ignore[arg-type]
                             model.right.title, f.family, rct.get("font_size", f.size_body + 1),
                             rct.get("color", c.accent), bold=rct.get("bold", True))
            bullets_top = rcb["top"] if model.right.title else (rct["top"] if rct is not None else rcb["top"])
            right_bullets = [truncate_text(b, lim.bullet_max_chars, f"right.bullets[{i}]") for i, b in enumerate(model.right.bullets)]
            add_bullet_list(slide, rcb["left"], bullets_top, rcb["width"], rcb["height"],  # type: ignore[arg-type]
                            right_bullets, f.family, rcb.get("font_size", f.size_body), rcb.get("color", c.text_dark))

        render_layout_extras(slide, "content_two_col", set(_D.keys()), c.accent, model=model)  # type: ignore[arg-type]
