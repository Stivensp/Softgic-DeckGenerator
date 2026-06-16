from __future__ import annotations

from pptx.enum.text import PP_ALIGN

from src.layout import layout_config as lc
from src.layout.base_layout import BaseLayoutRenderer
from src.layout.helpers import add_colored_box, add_text_box, set_slide_background, truncate_text
from src.models.slides import StatCalloutSlide
from src.models.theme import ThemeModel

_D = {
    "accent_bar_top":    {"left": 0.0,  "top": 0.0,  "width": 13.33, "height": 0.12},
    "accent_bar_bottom": {"left": 0.0,  "top": 7.38, "width": 13.33, "height": 0.12},
    "big_number":        {"left": 0.5,  "top": 0.8,  "width": 12.33, "height": 2.8 },
    "label":             {"left": 0.5,  "top": 3.6,  "width": 12.33, "height": 0.75},
    "divider":           {"left": 5.5,  "top": 4.5,  "width": 2.33,  "height": 0.05},
    "context":           {"left": 1.5,  "top": 4.7,  "width": 10.33, "height": 0.55},
    "source":            {"left": 8.0,  "top": 6.6,  "width": 4.83,  "height": 0.45},
}


def _p(el: str) -> dict:
    return {**_D[el], **lc.pos("stat_callout", el)}


class StatCalloutRenderer(BaseLayoutRenderer):
    def render(self, slide: object, model: object, theme: ThemeModel) -> None:
        assert isinstance(model, StatCalloutSlide)
        c = theme.colors
        f = theme.fonts
        lim = theme.limits

        set_slide_background(slide, c.primary)  # type: ignore[arg-type]

        for bar in ("accent_bar_top", "accent_bar_bottom"):
            p = _p(bar)
            add_colored_box(slide, p["left"], p["top"], p["width"], p["height"], c.accent)  # type: ignore[arg-type]

        p = _p("big_number")
        add_text_box(slide, p["left"], p["top"], p["width"], p["height"],  # type: ignore[arg-type]
                     model.big_number, f.family, f.size_stat, c.accent, bold=True, align=PP_ALIGN.CENTER)

        p = _p("label")
        label = truncate_text(model.label, lim.title_max_chars, "stat_callout.label")
        add_text_box(slide, p["left"], p["top"], p["width"], p["height"],  # type: ignore[arg-type]
                     label, f.family, f.size_subheading, c.text_light, bold=f.bold_headings, align=PP_ALIGN.CENTER)

        p = _p("divider")
        add_colored_box(slide, p["left"], p["top"], p["width"], p["height"], c.accent)  # type: ignore[arg-type]

        if model.context:
            p = _p("context")
            add_text_box(slide, p["left"], p["top"], p["width"], p["height"],  # type: ignore[arg-type]
                         model.context, f.family, f.size_body, c.text_muted, align=PP_ALIGN.CENTER)

        if model.source:
            p = _p("source")
            add_text_box(slide, p["left"], p["top"], p["width"], p["height"],  # type: ignore[arg-type]
                         f"Fuente: {model.source}", f.family, f.size_caption, c.text_muted,
                         align=PP_ALIGN.RIGHT, italic=True)
