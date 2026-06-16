from __future__ import annotations

import logging
from pathlib import Path

from src.exceptions import AssetValidationError

logger = logging.getLogger("softgic.validator.assets")


class AssetValidator:
    def __init__(self, min_width: int = 200, min_height: int = 200) -> None:
        self._min_width = min_width
        self._min_height = min_height

    def validate(self, data: dict, min_width: int | None = None, min_height: int | None = None) -> None:  # type: ignore[type-arg]
        min_w = min_width if min_width is not None else self._min_width
        min_h = min_height if min_height is not None else self._min_height

        for idx, slide in enumerate(data.get("slides", [])):
            if slide.get("type") == "profile_card":
                photo = slide.get("photo")
                if photo:
                    self._validate_image(photo, idx, min_w, min_h)

    def _validate_image(self, photo_path: str, idx: int, min_w: int, min_h: int) -> None:
        path = Path(photo_path)
        if not path.exists():
            raise AssetValidationError(
                f"Profile photo not found: '{photo_path}'",
                field="photo",
                slide_index=idx,
            )

        try:
            from PIL import Image

            with Image.open(path) as img:
                w, h = img.size
        except Exception as exc:
            raise AssetValidationError(
                f"Profile photo is not a valid image: '{photo_path}' — {exc}",
                field="photo",
                slide_index=idx,
            ) from exc

        if w < min_w or h < min_h:
            logger.warning(
                "[slide %d] Profile photo resolution %dx%d is below recommended %dx%d: '%s'",
                idx + 1,
                w,
                h,
                min_w,
                min_h,
                photo_path,
            )
