from __future__ import annotations

from typing import Any

from pptx.enum.text import PP_ALIGN

from src.layout import layout_config as lc
from src.layout.base_layout import BaseLayoutRenderer
from src.layout.helpers import (
    add_bullet_list,
    add_colored_box,
    add_image,
    add_text_box,
    render_layout_extras,
    set_shape_text,
    set_slide_background,
    truncate_text,
)
from src.models.slides import ProfileCardSlide
from src.models.theme import ThemeModel

_D = {
    "left_panel":       {"left": 0.0, "top": 0.0,  "width": 4.6,  "height": 7.5 },
    "left_accent_bar":  {"left": 0.0, "top": 0.0,  "width": 4.6,  "height": 0.12},
    "photo":            {"left": 0.8, "top": 0.6,  "width": 3.0,  "height": 3.0 },
    "name":             {"left": 0.5, "top": 3.75, "width": 4.0,  "height": 0.65},
    "role":             {"left": 0.5, "top": 4.4,  "width": 4.0,  "height": 0.45},
    "seniority":        {"left": 0.5, "top": 4.85, "width": 4.0,  "height": 0.4 },
    "panel_divider":    {"left": 4.6, "top": 0.0,  "width": 0.06, "height": 7.5 },
    "skills_label":     {"left": 5.0, "top": 0.35, "width": 7.8,  "height": 0.45},
    "skills_area":      {"left": 5.0, "top": 0.85, "width": 7.8,  "height": 1.8 },
    "highlights_label": {"left": 5.0, "top": 2.95, "width": 7.8,  "height": 0.4 },
    "highlights_area":  {"left": 5.0, "top": 3.4,  "width": 7.8,  "height": 3.6 },
}


def _p(el: str) -> dict[str, Any] | None:
    override = lc.pos("profile_card", el)
    if override is None:
        return None
    return {**_D[el], **override}


class ProfileCardRenderer(BaseLayoutRenderer):
    def render(self, slide: object, model: object, theme: ThemeModel) -> None:
        assert isinstance(model, ProfileCardSlide)
        c = theme.colors
        f = theme.fonts
        lim = theme.limits

        set_slide_background(slide, c.background)  # type: ignore[arg-type]

        if (p := _p("left_panel")) is not None:
            add_colored_box(slide, p["left"], p["top"], p["width"], p["height"], c.primary)  # type: ignore[arg-type]

        if (p := _p("left_accent_bar")) is not None:
            add_colored_box(slide, p["left"], p["top"], p["width"], p["height"], c.accent)  # type: ignore[arg-type]

        if (p := _p("photo")) is not None:
            photo_src = model.photo or theme.assets.placeholder_profile
            add_image(slide, photo_src, p["left"], p["top"], p["width"], p["height"],  # type: ignore[arg-type]
                      fallback_path=theme.assets.placeholder_profile)

        name = truncate_text(model.name, lim.title_max_chars, "profile_card.name")
        if (p := _p("name")) is not None:
            add_text_box(slide, p["left"], p["top"], p["width"], p["height"], name,  # type: ignore[arg-type]
                         f.family, p.get("font_size", f.size_body + 3), p.get("color", c.text_light), bold=p.get("bold", True))

        if (p := _p("role")) is not None:
            add_text_box(slide, p["left"], p["top"], p["width"], p["height"], model.role,  # type: ignore[arg-type]
                         f.family, p.get("font_size", f.size_body), p.get("color", c.accent))

        if (p := _p("seniority")) is not None:
            meta = f"{model.seniority}  ·  {model.years_experience} anos de exp."
            add_text_box(slide, p["left"], p["top"], p["width"], p["height"], meta,  # type: ignore[arg-type]
                         f.family, p.get("font_size", f.size_caption + 1), p.get("color", c.text_muted))

        if (p := _p("panel_divider")) is not None:
            add_colored_box(slide, p["left"], p["top"], p["width"], p["height"], c.accent)  # type: ignore[arg-type]

        # Skills
        if (p := _p("skills_label")) is not None:
            add_text_box(slide, p["left"], p["top"], p["width"], p["height"], "HABILIDADES TECNICAS",  # type: ignore[arg-type]
                         f.family, p.get("font_size", f.size_caption), p.get("color", c.accent), bold=p.get("bold", True))

        pa = _p("skills_area")
        if pa is not None:
            skills = model.skills[: lim.skills_max]
            tag_w, tag_h = 1.8, 0.38
            gap_x, gap_y = 0.15, 0.12
            cur_x, cur_y = pa["left"], pa["top"]
            for skill in skills:
                if cur_x + tag_w > pa["left"] + pa["width"]:
                    cur_x = pa["left"]
                    cur_y += tag_h + gap_y
                box = add_colored_box(slide, cur_x, cur_y, tag_w, tag_h, c.secondary)  # type: ignore[arg-type]
                set_shape_text(box, skill, f.family, f.size_caption + 1, c.text_light, align=PP_ALIGN.CENTER)
                cur_x += tag_w + gap_x

        # Highlights: elemento independiente con su propia posicion (no depende
        # de la visibilidad de skills_area — antes ocultar skills_area tambien
        # ocultaba los logros, aunque conceptualmente no estan relacionados).
        if model.highlights:
            hl_bullets = [truncate_text(h, lim.bullet_max_chars, f"highlights[{i}]") for i, h in enumerate(model.highlights)]
            if (p := _p("highlights_label")) is not None:
                add_text_box(slide, p["left"], p["top"], p["width"], p["height"], "LOGROS DESTACADOS",  # type: ignore[arg-type]
                             f.family, p.get("font_size", f.size_caption), p.get("color", c.accent), bold=p.get("bold", True))
            if (p := _p("highlights_area")) is not None:
                add_bullet_list(slide, p["left"], p["top"], p["width"], p["height"],  # type: ignore[arg-type]
                                hl_bullets, f.family, p.get("font_size", f.size_body), p.get("color", c.text_dark))

        render_layout_extras(slide, "profile_card", set(_D.keys()), c.accent, model=model)  # type: ignore[arg-type]
