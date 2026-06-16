from __future__ import annotations

import pytest

from src.exceptions import UnregisteredLayoutError
from src.layout.base_layout import BaseLayoutRenderer
from src.layout.registry import LayoutRegistry
from src.models.theme import ThemeModel


class _DummyRenderer(BaseLayoutRenderer):
    def render(self, slide: object, model: object, theme: ThemeModel) -> None:
        pass


def test_register_and_get() -> None:
    reg = LayoutRegistry()
    reg.register("cover", _DummyRenderer)
    cls = reg.get("cover")
    assert cls is _DummyRenderer


def test_get_unregistered_raises_error() -> None:
    reg = LayoutRegistry()
    with pytest.raises(UnregisteredLayoutError, match="wizard"):
        reg.get("wizard")


def test_available_types_sorted() -> None:
    reg = LayoutRegistry()
    reg.register("section_divider", _DummyRenderer)
    reg.register("cover", _DummyRenderer)
    assert reg.available_types() == ["cover", "section_divider"]


def test_register_overwrites_existing() -> None:
    class _Alt(BaseLayoutRenderer):
        def render(self, slide: object, model: object, theme: ThemeModel) -> None:
            pass

    reg = LayoutRegistry()
    reg.register("cover", _DummyRenderer)
    reg.register("cover", _Alt)
    assert reg.get("cover") is _Alt


def test_error_message_includes_available_types() -> None:
    reg = LayoutRegistry()
    reg.register("cover", _DummyRenderer)
    with pytest.raises(UnregisteredLayoutError, match="cover"):
        reg.get("nonexistent")
