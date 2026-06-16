from __future__ import annotations

import logging
from typing import Any

from src.validator.asset_validator import AssetValidator
from src.validator.business_rules import BusinessRulesValidator
from src.validator.schema_rules import SchemaRulesValidator

logger = logging.getLogger("softgic.validator")


class DeckValidator:
    def __init__(self) -> None:
        self._schema = SchemaRulesValidator()
        self._business = BusinessRulesValidator()
        self._assets = AssetValidator()

    def validate(self, data: dict[str, Any], limits: dict[str, Any] | None = None) -> None:
        logger.debug("Validation pass 1: schema rules")
        self._schema.validate(data)
        logger.debug("Validation pass 1: OK")

        logger.debug("Validation pass 2: business rules")
        self._business.validate(data, limits or {})
        logger.debug("Validation pass 2: OK")

        logger.debug("Validation pass 3: asset checks")
        if limits:
            self._assets.validate(
                data,
                min_width=limits.get("profile_image_min_width"),
                min_height=limits.get("profile_image_min_height"),
            )
        else:
            self._assets.validate(data)
        logger.debug("Validation pass 3: OK")

        slide_count = len(data.get("slides", []))
        logger.info("Validation passed — %d slide(s) ready to render", slide_count)
