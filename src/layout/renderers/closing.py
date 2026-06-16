from __future__ import annotations

from pptx.enum.text import PP_ALIGN

from src.layout import layout_config as lc
from src.layout.base_layout import BaseLayoutRenderer
from src.layout.helpers import add_colored_box, add_image, add_text_box, set_slide_background, truncate_text
from src.models.slides import ClosingSlide
from src.models.theme import ThemeModel

_D = {
    "accent_bar_top":    {"left": 0.0, "top": 0.0,  "width": 13.33, "height": 0.12},
    "accent_bar_bottom": {"left": 0.0, "top": 7.38, "width": 13.33, "height": 0.12},
    "logo":              {"left": 0.5, "top": 0.25, "width": 2.0,   "height": 0.65},
    "headline":          {"left": 0.8, "top": 1.3,  "width": 11.5,  "height": 1.2,  "font_size": 36},
    "divider":           {"left": 2.5, "top": 2.7,  "width": 8.33,  "height": 0.05},
    "contact_label":     {"left": 0.8, "top": 2.9,  "width": 12.0,  "height": 0.4,  "font_size": 10},
    "contact_name":      {"left": 0.8, "top": 3.35, "width": 4.0,   "height": 0.5,  "font_size": 14},
    "contact_email":     {"left": 4.8, "top": 3.35, "width": 4.0,   "height": 0.5,  "font_size": 14},
    "contact_phone":     {"left": 8.6, "top": 3.35, "width": 4.0,   "height": 0.5,  "font_size": 14},
    "cta_box":           {"left": 3.0, "top": 4.35, "width": 7.33,  "height": 0.75},
    "cta_text":          {"left": 3.1, "top": 4.45, "width": 7.13,  "height": 0.55, "font_size": 14},
}


def _p(el: str) -> dict:
    return {**_D[el], **lc.pos("closing", el)}


class ClosingRenderer(BaseLayoutRenderer):
    def render(self, slide: object, model: object, theme: ThemeModel) -> None:
        assert isinstance(model, ClosingSlide)
        c = theme.colors
        f = theme.fonts
        lim = theme.limits

        set_slide_background(slide, c.primary)  # type: ignore[arg-type]

        for bar in ("accent_bar_top", "accent_bar_bottom"):
            p = _p(bar)
            add_colored_box(slide, p["left"], p["top"], p["width"], p["height"], p.get("fill", c.accent))  # type: ignore[arg-type]

        p = _p("logo")
        add_image(slide, theme.assets.logo_white, p["left"], p["top"], p["width"], p["height"])  # type: ignore[arg-type]

        p = _p("headline")
        headline = truncate_text(model.headline, lim.title_max_chars, "closing.headline")
        add_text_box(slide, p["left"], p["top"], p["width"], p["height"], headline,  # type: ignore[arg-type]
                     f.family, p.get("font_size", f.size_heading), p.get("color", c.text_light),
                     bold=f.bold_headings, align=PP_ALIGN.CENTER)

        p = _p("divider")
        add_colored_box(slide, p["left"], p["top"], p["width"], p["height"], p.get("fill", c.accent))  # type: ignore[arg-type]

        p = _p("contact_label")
        add_text_box(slide, p["left"], p["top"], p["width"], p["height"], "CONTACTO",  # type: ignore[arg-type]
                     f.family, p.get("font_size", f.size_caption), p.get("color", c.accent),
                     bold=True, align=PP_ALIGN.CENTER)

        p = _p("contact_name")
        add_text_box(slide, p["left"], p["top"], p["width"], p["height"], model.contact_name,  # type: ignore[arg-type]
                     f.family, p.get("font_size", f.size_body), p.get("color", c.text_light),
                     bold=True, align=PP_ALIGN.CENTER)

        p = _p("contact_email")
        add_text_box(slide, p["left"], p["top"], p["width"], p["height"], model.contact_email,  # type: ignore[arg-type]
                     f.family, p.get("font_size", f.size_body), p.get("color", c.accent),
                     align=PP_ALIGN.CENTER)

        if model.contact_phone:
            p = _p("contact_phone")
            add_text_box(slide, p["left"], p["top"], p["width"], p["height"], model.contact_phone,  # type: ignore[arg-type]
                         f.family, p.get("font_size", f.size_body), p.get("color", c.text_light),
                         align=PP_ALIGN.CENTER)

        if model.cta:
            p = _p("cta_box")
            add_colored_box(slide, p["left"], p["top"], p["width"], p["height"], p.get("fill", c.accent))  # type: ignore[arg-type]
            p = _p("cta_text")
            add_text_box(slide, p["left"], p["top"], p["width"], p["height"], model.cta,  # type: ignore[arg-type]
                         f.family, p.get("font_size", f.size_body), p.get("color", c.primary),
                         bold=True, align=PP_ALIGN.CENTER)
