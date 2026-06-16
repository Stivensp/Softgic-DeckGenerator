from __future__ import annotations

from pathlib import Path

import pytest
from PIL import Image

from src.exceptions import AssetValidationError
from src.validator.asset_validator import AssetValidator


def _make_image(path: Path, width: int = 400, height: int = 400) -> None:
    img = Image.new("RGB", (width, height), color=(128, 128, 128))
    img.save(str(path))


def test_valid_image_passes(tmp_path: Path) -> None:
    photo = tmp_path / "photo.jpg"
    _make_image(photo, 400, 400)
    AssetValidator(min_width=200, min_height=200).validate(
        {
            "slides": [
                {
                    "type": "profile_card",
                    "name": "A",
                    "role": "B",
                    "years_experience": 1,
                    "seniority": "Mid",
                    "skills": ["x"],
                    "photo": str(photo),
                }
            ]
        }
    )


def test_missing_image_raises_asset_error(tmp_path: Path) -> None:
    with pytest.raises(AssetValidationError, match="not found"):
        AssetValidator().validate(
            {
                "slides": [
                    {
                        "type": "profile_card",
                        "name": "A",
                        "role": "B",
                        "years_experience": 1,
                        "seniority": "Mid",
                        "skills": ["x"],
                        "photo": str(tmp_path / "ghost.jpg"),
                    }
                ]
            }
        )


def test_corrupt_image_raises_asset_error(tmp_path: Path) -> None:
    bad = tmp_path / "bad.jpg"
    bad.write_bytes(b"not an image")
    with pytest.raises(AssetValidationError, match="valid image"):
        AssetValidator().validate(
            {
                "slides": [
                    {
                        "type": "profile_card",
                        "name": "A",
                        "role": "B",
                        "years_experience": 1,
                        "seniority": "Mid",
                        "skills": ["x"],
                        "photo": str(bad),
                    }
                ]
            }
        )


def test_small_image_emits_warning_not_error(tmp_path: Path, caplog) -> None:
    import logging

    photo = tmp_path / "tiny.jpg"
    _make_image(photo, 100, 100)
    with caplog.at_level(logging.WARNING, logger="softgic"):
        AssetValidator(min_width=200, min_height=200).validate(
            {
                "slides": [
                    {
                        "type": "profile_card",
                        "name": "A",
                        "role": "B",
                        "years_experience": 1,
                        "seniority": "Mid",
                        "skills": ["x"],
                        "photo": str(photo),
                    }
                ]
            }
        )
    assert any("resolution" in r.message.lower() or "below" in r.message.lower() for r in caplog.records)


def test_no_photo_field_skips_validation() -> None:
    AssetValidator().validate(
        {
            "slides": [
                {
                    "type": "profile_card",
                    "name": "A",
                    "role": "B",
                    "years_experience": 1,
                    "seniority": "Mid",
                    "skills": ["x"],
                }
            ]
        }
    )


def test_non_profile_card_slides_skipped(tmp_path: Path) -> None:
    AssetValidator().validate(
        {"slides": [{"type": "cover", "title": "T"}]}
    )
