from __future__ import annotations

import logging
from typing import Any

from pydantic import ValidationError

from src.exceptions import ParseError
from src.models.deck import DeckModel

logger = logging.getLogger("softgic.parser")


class DeckParser:
    def parse(self, data: dict[str, Any]) -> DeckModel:
        logger.debug("Parsing %d slide(s) into domain models", len(data.get("slides", [])))
        try:
            deck = DeckModel.model_validate(data)
        except ValidationError as exc:
            raise ParseError(f"Failed to parse input into domain model: {exc}") from exc

        logger.info("Parsed %d slide(s) successfully", len(deck.slides))
        return deck
