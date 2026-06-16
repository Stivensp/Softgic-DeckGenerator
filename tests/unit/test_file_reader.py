from __future__ import annotations

import json
from pathlib import Path

import pytest
import yaml

from src.exceptions import InputReadError
from src.input.file_reader import FileReader


@pytest.fixture
def reader() -> FileReader:
    return FileReader()


def test_read_yaml_returns_dict(tmp_path: Path, reader: FileReader) -> None:
    f = tmp_path / "deck.yaml"
    f.write_text("slides:\n  - type: cover\n    title: Hello\n", encoding="utf-8")
    result = reader.read(f)
    assert isinstance(result, dict)
    assert result["slides"][0]["title"] == "Hello"


def test_read_json_returns_dict(tmp_path: Path, reader: FileReader) -> None:
    f = tmp_path / "deck.json"
    f.write_text(json.dumps({"slides": [{"type": "cover", "title": "Hi"}]}), encoding="utf-8")
    result = reader.read(f)
    assert result["slides"][0]["title"] == "Hi"


def test_read_yml_extension(tmp_path: Path, reader: FileReader) -> None:
    f = tmp_path / "deck.yml"
    f.write_text("slides:\n  - type: cover\n    title: YML\n", encoding="utf-8")
    result = reader.read(f)
    assert result["slides"][0]["title"] == "YML"


def test_missing_file_raises_input_read_error(tmp_path: Path, reader: FileReader) -> None:
    with pytest.raises(InputReadError, match="not found"):
        reader.read(tmp_path / "nonexistent.yaml")


def test_empty_file_raises_input_read_error(tmp_path: Path, reader: FileReader) -> None:
    f = tmp_path / "empty.yaml"
    f.write_text("", encoding="utf-8")
    with pytest.raises(InputReadError, match="empty"):
        reader.read(f)


def test_invalid_yaml_raises_input_read_error(tmp_path: Path, reader: FileReader) -> None:
    f = tmp_path / "bad.yaml"
    f.write_text("key: [unclosed", encoding="utf-8")
    with pytest.raises(InputReadError, match="parse"):
        reader.read(f)


def test_invalid_json_raises_input_read_error(tmp_path: Path, reader: FileReader) -> None:
    f = tmp_path / "bad.json"
    f.write_text("{bad json", encoding="utf-8")
    with pytest.raises(InputReadError, match="parse"):
        reader.read(f)


def test_unsupported_extension_raises_error(tmp_path: Path, reader: FileReader) -> None:
    f = tmp_path / "deck.txt"
    f.write_text("slides: []", encoding="utf-8")
    with pytest.raises(InputReadError, match="Unsupported"):
        reader.read(f)


def test_yaml_list_at_root_raises_error(tmp_path: Path, reader: FileReader) -> None:
    f = tmp_path / "list.yaml"
    f.write_text("- item1\n- item2\n", encoding="utf-8")
    with pytest.raises(InputReadError, match="mapping"):
        reader.read(f)
