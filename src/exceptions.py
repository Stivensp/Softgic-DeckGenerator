from __future__ import annotations


class SoftgicDeckError(Exception):
    """Base exception for all Softgic Deck Generator errors."""

    def __init__(self, message: str, field: str | None = None, slide_index: int | None = None) -> None:
        self.message = message
        self.field = field
        self.slide_index = slide_index
        super().__init__(self._format())

    def _format(self) -> str:
        parts = [self.message]
        if self.slide_index is not None:
            parts.insert(0, f"[slide {self.slide_index + 1}]")
        if self.field:
            parts.append(f"(field: '{self.field}')")
        return " ".join(parts)


class InputReadError(SoftgicDeckError):
    pass


class ValidationError(SoftgicDeckError):
    pass


class SchemaValidationError(ValidationError):
    pass


class BusinessRuleError(ValidationError):
    pass


class AssetValidationError(ValidationError):
    pass


class ParseError(SoftgicDeckError):
    pass


class ThemeLoadError(SoftgicDeckError):
    pass


class UnregisteredLayoutError(SoftgicDeckError):
    pass


class RenderError(SoftgicDeckError):
    pass


class ExportError(SoftgicDeckError):
    pass
