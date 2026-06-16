from __future__ import annotations

from pydantic import BaseModel, field_validator

from src.models.slides import SlideModel


class DeckMetadata(BaseModel):
    generated_at: str | None = None
    input_file: str | None = None
    theme_version: str | None = None


class DeckModel(BaseModel):
    slides: list[SlideModel]
    metadata: DeckMetadata | None = None

    @field_validator("slides")
    @classmethod
    def slides_not_empty(cls, v: list[SlideModel]) -> list[SlideModel]:
        if not v:
            raise ValueError("deck must contain at least one slide")
        return v
