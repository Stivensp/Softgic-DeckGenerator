from __future__ import annotations

from abc import ABC, abstractmethod
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from pptx.slide import Slide

    from src.models.slides import (
        CoverSlide,
        SectionDividerSlide,
        ContentTwoColSlide,
        PricingTableSlide,
        ProfileCardSlide,
        StatCalloutSlide,
        ClosingSlide,
    )
    from src.models.theme import ThemeModel

    SlideModel = (
        CoverSlide
        | SectionDividerSlide
        | ContentTwoColSlide
        | PricingTableSlide
        | ProfileCardSlide
        | StatCalloutSlide
        | ClosingSlide
    )


class BaseLayoutRenderer(ABC):
    @abstractmethod
    def render(self, slide: "Slide", model: "SlideModel", theme: "ThemeModel") -> None:  # type: ignore[override]
        ...
