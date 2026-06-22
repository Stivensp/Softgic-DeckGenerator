from __future__ import annotations

from pathlib import Path

import pytest


@pytest.fixture(scope="session")
def template_master_path() -> Path:
    """config/template_master.pptx es un asset binario de diseno real (no
    generado en el test) — si no esta presente en este checkout, los tests
    que dependen de el se saltan con un aviso claro en vez de fallar CI.
    """
    p = Path(__file__).resolve().parent.parent.parent / "config" / "template_master.pptx"
    if not p.exists():
        pytest.skip("config/template_master.pptx no disponible en este entorno")
    return p
