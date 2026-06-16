from __future__ import annotations

import logging
import os
import tempfile
from pathlib import Path

from pptx import Presentation

from src.exceptions import ExportError

logger = logging.getLogger("softgic.exporter")


class PptxExporter:
    def export(self, presentation: Presentation, output_path: Path) -> None:
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        tmp_fd, tmp_path = tempfile.mkstemp(suffix=".pptx", dir=output_path.parent)
        try:
            os.close(tmp_fd)
            presentation.save(str(tmp_path))
            os.replace(tmp_path, output_path)
        except Exception as exc:
            _safe_remove(tmp_path)
            raise ExportError(f"Failed to write output file '{output_path}': {exc}") from exc

        size_kb = output_path.stat().st_size / 1024
        logger.info("Export complete: %s (%.1f KB)", output_path, size_kb)


def _safe_remove(path: str) -> None:
    try:
        os.unlink(path)
    except OSError:
        pass
