from __future__ import annotations

import logging
from typing import Type

from src.exceptions import UnregisteredLayoutError
from src.layout.base_layout import BaseLayoutRenderer

logger = logging.getLogger("softgic.layout.registry")


class LayoutRegistry:
    def __init__(self) -> None:
        self._registry: dict[str, Type[BaseLayoutRenderer]] = {}

    def register(self, slide_type: str, renderer_class: Type[BaseLayoutRenderer]) -> None:
        self._registry[slide_type] = renderer_class
        logger.debug("Registered layout renderer: '%s' → %s", slide_type, renderer_class.__name__)

    def get(self, slide_type: str) -> Type[BaseLayoutRenderer]:
        if slide_type not in self._registry:
            raise UnregisteredLayoutError(
                f"No renderer registered for slide type '{slide_type}'. "
                f"Available: {sorted(self._registry.keys())}"
            )
        return self._registry[slide_type]

    def available_types(self) -> list[str]:
        return sorted(self._registry.keys())
