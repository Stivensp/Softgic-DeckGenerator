from __future__ import annotations

import logging

from pptx import Presentation

from src.exceptions import RenderError
from src.layout.registry import LayoutRegistry
from src.models.deck import DeckModel
from src.models.theme import ThemeModel

logger = logging.getLogger("softgic.renderer")


class DeckRenderer:
    def render(
        self,
        deck: DeckModel,
        theme: ThemeModel,
        registry: LayoutRegistry,
        presentation: Presentation,
    ) -> Presentation:
        total = len(deck.slides)
        logger.info("Rendering %d slide(s)...", total)

        for i, slide_model in enumerate(deck.slides):
            slide_type = slide_model.type  # type: ignore[union-attr]
            logger.debug("Rendering slide %d/%d: %s", i + 1, total, slide_type)

            layout_index = self._layout_index(slide_type, theme, len(presentation.slide_layouts))
            layout = presentation.slide_layouts[layout_index]

            try:
                slide = presentation.slides.add_slide(layout)
                renderer_class = registry.get(slide_type)
                renderer_class().render(slide, slide_model, theme)
            except Exception as exc:
                raise RenderError(
                    f"Failed to render slide {i + 1} (type='{slide_type}'): {exc}",
                    slide_index=i,
                ) from exc

            logger.info("Slide %d/%d rendered: %s", i + 1, total, slide_type)

        logger.info("All %d slide(s) rendered successfully", total)
        return presentation

    def _layout_index(self, slide_type: str, theme: ThemeModel, available: int) -> int:
        configured = getattr(theme.layout_indices, slide_type, 6)
        if configured >= available:
            logger.warning(
                "Layout index %d for '%s' exceeds available layouts (%d); using blank layout (last available)",
                configured,
                slide_type,
                available,
            )
            return available - 1
        return configured
