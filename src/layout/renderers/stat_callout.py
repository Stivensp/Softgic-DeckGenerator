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
from src.models.slides import StatCalloutSlide
from src.models.theme import ThemeModel

_D = {
    "accent_bar_top":    {"left": 0.0,  "top": 0.0,  "width": 13.33, "height": 0.12},
    "accent_bar_bottom": {"left": 0.0,  "top": 7.38, "width": 13.33, "height": 0.12},
    "big_number":        {"left": 0.5,  "top": 0.8,  "width": 12.33, "height": 2.8,  "font_size": 72},
    "label":             {"left": 0.5,  "top": 3.6,  "width": 12.33, "height": 0.75, "font_size": 24},
    "divider":           {"left": 5.5,  "top": 4.5,  "width": 2.33,  "height": 0.05},
    "context":           {"left": 1.5,  "top": 4.7,  "width": 10.33, "height": 0.55, "font_size": 14},
    "source":            {"left": 8.0,  "top": 6.6,  "width": 4.83,  "height": 0.45, "font_size": 10},
}


def _p(el: str) -> dict[str, Any] | None:
    override = lc.pos("stat_callout", el)
    if override is None:
        return None
    return {**_D[el], **override}


class StatCalloutRenderer(BaseLayoutRenderer):
    def render(self, slide: object, model: object, theme: ThemeModel) -> None:
        assert isinstance(model, StatCalloutSlide)
        c = theme.colors
        f = theme.fonts
        lim = theme.limits

        set_slide_background(slide, c.background)  # type: ignore[arg-type]

        for bar in ("accent_bar_top", "accent_bar_bottom"):
            if (p := _p(bar)) is not None:
                add_colored_box(slide, p["left"], p["top"], p["width"], p["height"], p.get("fill", c.accent))  # type: ignore[arg-type]

        if (p := _p("big_number")) is not None:
            add_text_box(slide, p["left"], p["top"], p["width"], p["height"],  # type: ignore[arg-type]
                         model.big_number, f.family, p.get("font_size", f.size_stat),
                         p.get("color", c.accent), bold=p.get("bold", True), align=PP_ALIGN.CENTER)

        label = truncate_text(model.label, lim.title_max_chars, "stat_callout.label")
        if (p := _p("label")) is not None:
            add_text_box(slide, p["left"], p["top"], p["width"], p["height"],  # type: ignore[arg-type]
                         label, f.family, p.get("font_size", f.size_subheading),
                         p.get("color", c.text_dark), bold=p.get("bold", f.bold_headings), align=PP_ALIGN.CENTER)

        if (p := _p("divider")) is not None:
            add_colored_box(slide, p["left"], p["top"], p["width"], p["height"], p.get("fill", c.accent))  # type: ignore[arg-type]

        if model.context and (p := _p("context")) is not None:
            add_text_box(slide, p["left"], p["top"], p["width"], p["height"],  # type: ignore[arg-type]
                         model.context, f.family, p.get("font_size", f.size_body),
                         p.get("color", c.text_muted), align=PP_ALIGN.CENTER)

        if model.source and (p := _p("source")) is not None:
            add_text_box(slide, p["left"], p["top"], p["width"], p["height"],  # type: ignore[arg-type]
                         f"Fuente: {model.source}", f.family, p.get("font_size", f.size_caption),
                         p.get("color", c.text_muted), align=PP_ALIGN.RIGHT, italic=True)

        render_layout_extras(slide, "stat_callout", set(_D.keys()), c.accent, model=model)  # type: ignore[arg-type]
