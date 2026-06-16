from __future__ import annotations

from pptx.enum.text import PP_ALIGN

from src.layout import layout_config as lc
from src.layout.base_layout import BaseLayoutRenderer
from src.layout.helpers import add_bullet_list, add_colored_box, add_text_box, set_slide_background, truncate_text
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


def _p(el: str) -> dict:
    return {**_D[el], **lc.pos("content_two_col", el)}


class ContentTwoColRenderer(BaseLayoutRenderer):
    def render(self, slide: object, model: object, theme: ThemeModel) -> None:
        assert isinstance(model, ContentTwoColSlide)
        c = theme.colors
        f = theme.fonts
        lim = theme.limits

        set_slide_background(slide, c.background)  # type: ignore[arg-type]

        p = _p("header_bar")
        add_colored_box(slide, p["left"], p["top"], p["width"], p["height"], p.get("fill", c.primary))  # type: ignore[arg-type]

        p = _p("title")
        title = truncate_text(model.title, lim.title_max_chars, "content_two_col.title")
        add_text_box(slide, p["left"], p["top"], p["width"], p["height"], title,  # type: ignore[arg-type]
                     f.family, p.get("font_size", f.size_subheading), p.get("color", c.text_light),
                     bold=f.bold_headings)

        p = _p("accent_line")
        add_colored_box(slide, p["left"], p["top"], p["width"], p["height"], p.get("fill", c.accent))  # type: ignore[arg-type]

        # Left column
        lct = _p("left_col_title")
        lcb = _p("left_col_bullets")
        if model.left.title:
            add_text_box(slide, lct["left"], lct["top"], lct["width"], lct["height"],  # type: ignore[arg-type]
                         model.left.title, f.family, lct.get("font_size", f.size_body + 1),
                         lct.get("color", c.accent), bold=True)
        bullets_top = lcb["top"] if model.left.title else lct["top"]
        left_bullets = [truncate_text(b, lim.bullet_max_chars, f"left.bullets[{i}]") for i, b in enumerate(model.left.bullets)]
        add_bullet_list(slide, lcb["left"], bullets_top, lcb["width"], lcb["height"],  # type: ignore[arg-type]
                        left_bullets, f.family, lcb.get("font_size", f.size_body), lcb.get("color", c.text_dark))

        p = _p("divider")
        add_colored_box(slide, p["left"], p["top"], p["width"], p["height"], p.get("fill", c.divider))  # type: ignore[arg-type]

        # Right column
        rct = _p("right_col_title")
        rcb = _p("right_col_bullets")
        if model.right.title:
            add_text_box(slide, rct["left"], rct["top"], rct["width"], rct["height"],  # type: ignore[arg-type]
                         model.right.title, f.family, rct.get("font_size", f.size_body + 1),
                         rct.get("color", c.accent), bold=True)
        bullets_top = rcb["top"] if model.right.title else rct["top"]
        right_bullets = [truncate_text(b, lim.bullet_max_chars, f"right.bullets[{i}]") for i, b in enumerate(model.right.bullets)]
        add_bullet_list(slide, rcb["left"], bullets_top, rcb["width"], rcb["height"],  # type: ignore[arg-type]
                        right_bullets, f.family, rcb.get("font_size", f.size_body), rcb.get("color", c.text_dark))
