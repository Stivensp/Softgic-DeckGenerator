"""
Parsea un .pptx con el motor OOXML generico (src/ooxml_engine/) y vuelca el
modelo interno completo como YAML legible.

Entregable demostrable de la Fase 1 (M1) del motor universal de PPTX: antes
de que exista un renderer/round-trip, esto ya prueba que el parser entiende
la estructura real de un archivo PowerPoint arbitrario — posicion, texto
(runs/parrafos/fuente/color), relleno solido, e imagenes, sin estar atado a
los 9 tipos de slide fijos del generador.

No modifica nada del sistema de generacion existente (src/models/,
src/layout/) — es un modulo independiente, ver tools/apply_visual_config.py
para el flujo del editor visual que SI sigue siendo el del sistema actual.

Uso:
    python tools/ooxml_inspect.py config/template_master.pptx
    python tools/ooxml_inspect.py config/template_master.pptx --slide 0
    python tools/ooxml_inspect.py config/template_master.pptx --format json
"""
from __future__ import annotations

import argparse
import dataclasses
import json
import sys
from enum import Enum
from pathlib import Path
from typing import Any

import yaml

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.ooxml_engine.model.document import OoxmlPresentation  # noqa: E402
from src.ooxml_engine.parser.presentation_parser import parse_pptx  # noqa: E402


def _to_plain(value: Any) -> Any:
    """Convierte el arbol de dataclasses a tipos planos (dict/list/str/...)
    para volcar a YAML/JSON. bytes (ImageElement.blob) nunca se vuelcan
    crudos — se reemplazan por un resumen legible.
    """
    if dataclasses.is_dataclass(value) and not isinstance(value, type):
        return {f.name: _to_plain(getattr(value, f.name)) for f in dataclasses.fields(value)}
    if isinstance(value, Enum):
        return value.value
    if isinstance(value, bytes):
        return f"<{len(value)} bytes>"
    if isinstance(value, (list, tuple)):
        return [_to_plain(v) for v in value]
    if isinstance(value, dict):
        return {k: _to_plain(v) for k, v in value.items()}
    return value


def _dump(doc: OoxmlPresentation, slide_filter: int | None, fmt: str) -> str:
    data = _to_plain(doc)
    if slide_filter is not None:
        slides = data["slides"]
        data["slides"] = [s for s in slides if s["index"] == slide_filter]
    if fmt == "json":
        return json.dumps(data, indent=2, ensure_ascii=False)
    return yaml.dump(data, sort_keys=False, allow_unicode=True, default_flow_style=False)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Inspecciona un .pptx con el motor OOXML generico y vuelca su modelo interno"
    )
    parser.add_argument("pptx_path", help="Ruta al archivo .pptx a inspeccionar")
    parser.add_argument("--slide", type=int, default=None, help="Mostrar solo este indice de slide (0-based)")
    parser.add_argument("--format", choices=["yaml", "json"], default="yaml")
    parser.add_argument("--output", "-o", default=None, help="Archivo de salida (default: stdout)")
    args = parser.parse_args()

    path = Path(args.pptx_path)
    if not path.exists():
        print(f"ERROR: No se encontro {path}", file=sys.stderr)
        return 1

    try:
        doc = parse_pptx(path)
    except Exception as exc:
        print(f"ERROR: No se pudo parsear {path}: {exc}", file=sys.stderr)
        return 1

    if args.slide is not None and not any(s.index == args.slide for s in doc.slides):
        print(f"ERROR: el archivo tiene {len(doc.slides)} slide(s); indice {args.slide} fuera de rango", file=sys.stderr)
        return 1

    output = _dump(doc, args.slide, args.format)

    if args.output:
        Path(args.output).write_text(output, encoding="utf-8")
        print(f"Volcado escrito en: {args.output}")
    else:
        print(output)

    return 0


if __name__ == "__main__":
    sys.exit(main())
