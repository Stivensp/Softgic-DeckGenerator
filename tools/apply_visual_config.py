"""
Lee assets/config_visual.pptx (editado en PowerPoint) y actualiza
config/layout.yaml con las nuevas posiciones.

Despues de correr esto, regenera tus decks para ver los cambios.
"""
from __future__ import annotations

import sys
from pathlib import Path

import yaml
from pptx import Presentation
from pptx.util import Emu

ROOT = Path(__file__).resolve().parent.parent
PPTX = ROOT / "assets" / "config_visual.pptx"
LAYOUT_YAML = ROOT / "config" / "layout.yaml"

EMU = 914400  # EMUs per inch

SLIDE_ORDER = [
    "cover", "section_divider", "content_two_col", "content_one_col",
    "pricing_table", "profile_card", "stat_callout", "closing", "grafico",
]


def emu_to_in(v: int) -> float:
    return round(v / EMU, 3)


def main() -> None:
    if not PPTX.exists():
        print(f"ERROR: No se encontro {PPTX}")
        print("Primero ejecuta: Herramientas de diseno -> Abrir editor visual")
        sys.exit(1)

    prs = Presentation(str(PPTX))

    if len(prs.slides) != len(SLIDE_ORDER):
        print(f"ERROR: Se esperaban {len(SLIDE_ORDER)} slides, el archivo tiene {len(prs.slides)}")
        sys.exit(1)

    # Leer posiciones actuales como base
    current: dict = {}
    if LAYOUT_YAML.exists():
        current = yaml.safe_load(LAYOUT_YAML.read_text(encoding="utf-8")) or {}

    updated = 0
    for slide_pptx, slide_type in zip(prs.slides, SLIDE_ORDER):
        if slide_type not in current:
            current[slide_type] = {}

        for shape in slide_pptx.shapes:
            name = shape.name
            if name.startswith("_"):  # skip internal shapes like _SLIDE_TYPE_
                continue
            if name not in current.get(slide_type, {}):
                continue  # only update known elements

            try:
                new_pos = {
                    "left":   emu_to_in(shape.left),
                    "top":    emu_to_in(shape.top),
                    "width":  emu_to_in(shape.width),
                    "height": emu_to_in(shape.height),
                }
                old_pos = current[slide_type][name]
                if new_pos != old_pos:
                    current[slide_type][name] = new_pos
                    updated += 1
                    print(f"  [{slide_type}] {name}: {old_pos} -> {new_pos}")
            except Exception as e:
                print(f"  AVISO: no se pudo leer posicion de '{name}' en '{slide_type}': {e}")

    # Escribir layout.yaml actualizado
    LAYOUT_YAML.write_text(
        "# config/layout.yaml — actualizado desde editor visual\n"
        "# Medidas en pulgadas. 1 pulgada = 2.54 cm\n\n"
        + yaml.dump(current, default_flow_style=False, allow_unicode=True, sort_keys=False),
        encoding="utf-8",
    )

    # Limpiar cache para que los renderers lean el nuevo archivo
    try:
        from src.layout import layout_config
        layout_config.reload()
    except Exception:
        pass

    if updated == 0:
        print("No hubo cambios de posicion.")
    else:
        print(f"\n{updated} posicion(es) actualizadas en config/layout.yaml")
        print("Regenera tus decks para ver los cambios.")


if __name__ == "__main__":
    main()
