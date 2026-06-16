"""
Lee assets/config_visual.pptx (editado en PowerPoint) y actualiza
config/layout.yaml con las nuevas posiciones, tamanoss de fuente y colores.

Que lee por tipo de elemento:
  - Rectangulos de color (barras, paneles): posicion + color de relleno
  - Cajas de texto:                         posicion + tamano de fuente + color del texto

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
PT_EMU = 12700  # EMUs per point

SLIDE_ORDER = [
    "cover", "section_divider", "content_two_col", "content_one_col",
    "pricing_table", "profile_card", "stat_callout", "closing", "grafico",
]


def _emu_to_in(v: int) -> float:
    return round(v / EMU, 3)


def _emu_to_pt(v) -> int | None:
    if v is None:
        return None
    return round(int(v) / PT_EMU)


def _rgb_to_hex(rgb) -> str:
    return f"{rgb.red:02X}{rgb.green:02X}{rgb.blue:02X}"


def _read_shape_props(shape) -> dict:
    """Lee posicion, color de relleno y estilo de fuente de un shape."""
    props: dict = {
        "left":   _emu_to_in(shape.left),
        "top":    _emu_to_in(shape.top),
        "width":  _emu_to_in(shape.width),
        "height": _emu_to_in(shape.height),
    }

    # shape_type 17 = TEXT_BOX → leer estilo de fuente del primer run
    # shape_type 1  = AUTO_SHAPE (rectangulo) → leer color de relleno
    try:
        if shape.shape_type == 17:
            # Es una caja de texto: leer font_size y color del primer run
            paras = shape.text_frame.paragraphs
            for para in paras:
                for run in para.runs:
                    if run.font.size:
                        props["font_size"] = _emu_to_pt(run.font.size)
                    try:
                        if run.font.color and run.font.color.type is not None:
                            props["color"] = _rgb_to_hex(run.font.color.rgb)
                    except Exception:
                        pass
                    break  # solo el primer run con datos
                if "font_size" in props or "color" in props:
                    break
        else:
            # Es un rectangulo: leer color de relleno
            fill = shape.fill
            if fill.type == 1:  # MSO_FILL.SOLID = 1
                props["fill"] = _rgb_to_hex(fill.fore_color.rgb)
    except Exception:
        pass

    return props


def _props_changed(old: dict, new: dict) -> list[str]:
    """Retorna lista de claves que cambiaron."""
    keys = set(old) | set(new)
    return [k for k in keys if old.get(k) != new.get(k)]


def main() -> None:
    if not PPTX.exists():
        print(f"ERROR: No se encontro {PPTX}")
        print("Primero ejecuta: Herramientas de diseno -> Abrir editor visual")
        sys.exit(1)

    prs = Presentation(str(PPTX))

    if len(prs.slides) != len(SLIDE_ORDER):
        print(f"ERROR: Se esperaban {len(SLIDE_ORDER)} slides, el archivo tiene {len(prs.slides)}")
        print("El archivo puede estar desactualizado. Regenera el editor visual.")
        sys.exit(1)

    # Cargar layout.yaml actual como base
    current: dict = {}
    if LAYOUT_YAML.exists():
        current = yaml.safe_load(LAYOUT_YAML.read_text(encoding="utf-8")) or {}

    updated = 0
    for slide_pptx, slide_type in zip(prs.slides, SLIDE_ORDER):
        if slide_type not in current:
            current[slide_type] = {}

        for shape in slide_pptx.shapes:
            name = shape.name
            if name.startswith("_"):
                continue  # saltar shapes internos como _LABEL_
            if name not in current.get(slide_type, {}):
                continue  # solo actualizar elementos conocidos en layout.yaml

            try:
                new_props = _read_shape_props(shape)
                old_props = current[slide_type][name]
                changed = _props_changed(old_props, new_props)

                if changed:
                    current[slide_type][name] = new_props
                    updated += 1
                    changes_str = ", ".join(
                        f"{k}: {old_props.get(k)} -> {new_props.get(k)}" for k in changed
                    )
                    print(f"  [{slide_type}] {name}: {changes_str}")
            except Exception as e:
                print(f"  AVISO: no se pudo leer '{name}' en '{slide_type}': {e}")

    # Escribir layout.yaml actualizado
    LAYOUT_YAML.write_text(
        "# config/layout.yaml — actualizado desde editor visual\n"
        "# Medidas en pulgadas (1\" = 2.54 cm). font_size en puntos. color/fill en hex RGB.\n\n"
        + yaml.dump(current, default_flow_style=False, allow_unicode=True, sort_keys=False),
        encoding="utf-8",
    )

    # Invalidar cache de layout_config
    try:
        from src.layout import layout_config
        layout_config.reload()
    except Exception:
        pass

    if updated == 0:
        print("Sin cambios detectados.")
    else:
        print(f"\n{updated} elemento(s) actualizado(s) en config/layout.yaml")
        print("Regenera tus decks para ver los cambios.")


if __name__ == "__main__":
    main()
