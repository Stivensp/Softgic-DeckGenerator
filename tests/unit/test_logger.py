from __future__ import annotations

import logging
from pathlib import Path

import pytest

from src.logger.deck_logger import DeckLogger


def test_init_creates_log_file(tmp_path: Path) -> None:
    dl = DeckLogger(log_dir=tmp_path)
    dl.init()
    log_files = list(tmp_path.glob("*_deck.log"))
    assert len(log_files) == 1


def test_init_adds_console_and_file_handlers(tmp_path: Path) -> None:
    dl = DeckLogger(log_dir=tmp_path)
    dl.init()
    root = logging.getLogger("softgic")
    assert len(root.handlers) == 2


def test_get_returns_child_logger(tmp_path: Path) -> None:
    dl = DeckLogger(log_dir=tmp_path)
    dl.init()
    child = dl.get("mymodule")
    assert child.name == "softgic.mymodule"


def test_init_twice_clears_old_handlers(tmp_path: Path) -> None:
    dl = DeckLogger(log_dir=tmp_path)
    dl.init()
    dl.init()
    root = logging.getLogger("softgic")
    assert len(root.handlers) == 2


def test_log_file_receives_messages(tmp_path: Path) -> None:
    dl = DeckLogger(log_dir=tmp_path)
    dl.init(level=logging.DEBUG)
    log = dl.get("test")
    log.info("hello from test")

    log_file = next(tmp_path.glob("*_deck.log"))
    content = log_file.read_text(encoding="utf-8")
    assert "hello from test" in content
