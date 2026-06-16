from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest
from pptx import Presentation

from src.exceptions import ExportError
from src.exporter.pptx_exporter import PptxExporter


@pytest.fixture
def presentation() -> Presentation:
    return Presentation()


def test_export_writes_valid_pptx(tmp_path: Path, presentation: Presentation) -> None:
    out = tmp_path / "out.pptx"
    PptxExporter().export(presentation, out)
    assert out.exists()
    assert out.stat().st_size > 0


def test_export_creates_parent_directories(tmp_path: Path, presentation: Presentation) -> None:
    out = tmp_path / "a" / "b" / "c" / "deck.pptx"
    PptxExporter().export(presentation, out)
    assert out.exists()


def test_export_output_is_openable_pptx(tmp_path: Path, presentation: Presentation) -> None:
    out = tmp_path / "deck.pptx"
    PptxExporter().export(presentation, out)
    reloaded = Presentation(str(out))
    assert reloaded is not None


def test_export_raises_export_error_on_save_failure(tmp_path: Path) -> None:
    out = tmp_path / "deck.pptx"
    broken_prs = MagicMock()
    broken_prs.save.side_effect = OSError("disk full")
    with pytest.raises(ExportError, match="disk full"):
        PptxExporter().export(broken_prs, out)


def test_export_cleans_up_temp_file_on_failure(tmp_path: Path) -> None:
    out = tmp_path / "deck.pptx"
    broken_prs = MagicMock()
    broken_prs.save.side_effect = OSError("fail")
    try:
        PptxExporter().export(broken_prs, out)
    except ExportError:
        pass
    # No leftover .pptx files except the output (which was never written)
    remaining = list(tmp_path.glob("*.pptx"))
    assert len(remaining) == 0
