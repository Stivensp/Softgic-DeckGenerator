from __future__ import annotations

import json
import logging
from pathlib import Path

import yaml

from src.exceptions import InputReadError
from src.input.formats import InputFormat

logger = logging.getLogger("softgic.file_reader")


class FileReader:
    def read(self, file_path: Path) -> dict:  # type: ignore[type-arg]
        logger.debug("Reading input file: %s", file_path)

        if not file_path.exists():
            raise InputReadError(f"Input file not found: {file_path}")

        fmt = self._detect_format(file_path)
        raw_text = file_path.read_text(encoding="utf-8").strip()

        if not raw_text:
            raise InputReadError(f"Input file is empty: {file_path}")

        data = self._parse(raw_text, fmt, file_path)

        if not isinstance(data, dict):
            raise InputReadError(
                f"Input file must be a YAML/JSON mapping (object), got {type(data).__name__}: {file_path}"
            )

        logger.info("Input file read — format: %s, path: %s", fmt.value.upper(), file_path)
        return data

    def _detect_format(self, file_path: Path) -> InputFormat:
        suffix = file_path.suffix.lower()
        if suffix in {".yaml", ".yml"}:
            return InputFormat.YAML
        if suffix == ".json":
            return InputFormat.JSON
        raise InputReadError(
            f"Unsupported file extension '{suffix}'. Supported: .yaml, .yml, .json"
        )

    def _parse(self, text: str, fmt: InputFormat, path: Path) -> object:
        try:
            if fmt == InputFormat.YAML:
                return yaml.safe_load(text)
            return json.loads(text)
        except (yaml.YAMLError, json.JSONDecodeError) as exc:
            raise InputReadError(f"Failed to parse {fmt.value.upper()} file '{path}': {exc}") from exc
