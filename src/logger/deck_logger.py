from __future__ import annotations

import logging
import sys
from datetime import datetime
from pathlib import Path


class DeckLogger:
    """Configures structured logging for the pipeline. Call init() once in generate.py."""

    _FORMAT = "[%(asctime)s] [%(levelname)-7s] [%(name)s] %(message)s"
    _DATE_FMT = "%Y-%m-%d %H:%M:%S"

    def __init__(self, log_dir: Path = Path("logs")) -> None:
        self._log_dir = log_dir
        self._root = logging.getLogger("softgic")

    def init(self, level: int = logging.INFO) -> None:
        self._root.setLevel(level)
        self._root.handlers.clear()

        formatter = logging.Formatter(self._FORMAT, datefmt=self._DATE_FMT)

        console = logging.StreamHandler(sys.stdout)
        console.setLevel(level)
        console.setFormatter(formatter)
        self._root.addHandler(console)

        self._log_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = self._log_dir / f"{timestamp}_deck.log"
        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(formatter)
        self._root.addHandler(file_handler)

        self._root.info("DeckLogger initialized — log file: %s", log_file)

    def get(self, name: str) -> logging.Logger:
        return logging.getLogger(f"softgic.{name}")
