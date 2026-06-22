"""
Softgic Deck Generator — Editor web de CONTENIDO de deck.

Esto NO es un editor de diseno/posiciones (eso sigue siendo PowerPoint +
tools/open_visual_editor.py + tools/apply_visual_config.py, sin cambios).
Este frontend sirve para llenar y enviar los DATOS de cada slide (titulos,
bullets, filas de precios, etc.) usando como esquema/punto de partida la
estructura de config/generated/input_template.yaml, y generar el .pptx final
con un boton — sin tocar YAML a mano.

Uso:
    python frontend/server.py
    -> abre http://127.0.0.1:5000
"""
from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

import yaml
from flask import Flask, jsonify, render_template, request, send_file
from flask_cors import CORS

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.layout import layout_config  # noqa: E402
from tools.apply_visual_config import (  # noqa: E402
    FIELD_TO_ELEMENTS,
    GENERATED_DIR,
    LAYOUT_YAML,
    SLIDE_DEFAULTS,
    SLIDE_ORDER,
    TEMPLATE_YAML,
    is_field_visible,
)

THEME_YAML = ROOT / "config" / "theme.yaml"
EXAMPLES_DIR = ROOT / "examples"
OUTPUT_DIR = ROOT / "output"

app = Flask(__name__, static_folder="static", template_folder="templates")
CORS(app)


def _load_layout_yaml() -> dict[str, Any]:
    """Lectura de solo-lectura de config/generated/layout.yaml — para saber
    que elementos canonicos estan visibles/ocultos en el editor visual actual
    y filtrar el esquema del formulario en consecuencia. Este servidor nunca
    escribe en layout.yaml; eso sigue siendo trabajo exclusivo de
    tools/apply_visual_config.py (PowerPoint -> layout.yaml).
    """
    if not LAYOUT_YAML.exists():
        return {}
    try:
        loaded = yaml.safe_load(LAYOUT_YAML.read_text(encoding="utf-8"))
        return loaded if isinstance(loaded, dict) else {}
    except Exception:
        return {}


# ───────────────────── esquema de campos por tipo de slide ────────────────────
# Espejo de los modelos pydantic en src/models/slides.py — describe que campos
# mostrar en el formulario y de que tipo es cada uno, para que app.js pueda
# construir el formulario dinamicamente sin hardcodear nada por tipo.

def _field(name: str, label: str, type_: str, **kwargs: Any) -> dict[str, Any]:
    return {"name": name, "label": label, "type": type_, **kwargs}


FIELD_SCHEMAS: dict[str, list[dict[str, Any]]] = {
    "cover": [
        _field("title", "Titulo", "text", required=True, max_chars_key="title_max_chars"),
        _field("subtitle", "Subtitulo", "text", max_chars_key="subtitle_max_chars"),
        _field("client", "Cliente", "text"),
        _field("date", "Fecha", "text"),
        _field("author", "Autor", "text"),
    ],
    "section_divider": [
        _field("section_number", "Numero de seccion", "text", required=True),
        _field(
            "section_title", "Titulo de seccion", "text",
            required=True, max_chars_key="title_max_chars",
        ),
        _field("tagline", "Tagline", "text"),
    ],
    "content_two_col": [
        _field("title", "Titulo", "text", required=True, max_chars_key="title_max_chars"),
        _field("left", "Columna izquierda", "bullet_column", required=True),
        _field("right", "Columna derecha", "bullet_column", required=True),
    ],
    "content_one_col": [
        _field("title", "Titulo", "text", required=True, max_chars_key="title_max_chars"),
        _field("body_title", "Encabezado de contenido", "text"),
        _field("bullets", "Bullets", "list", required=True),
    ],
    "pricing_table": [
        _field("title", "Titulo", "text", required=True, max_chars_key="title_max_chars"),
        _field("currency", "Moneda", "text", required=True),
        _field("rows", "Filas", "pricing_rows", required=True),
        _field("totals", "Total", "number", required=True),
        _field("notes", "Notas", "text"),
    ],
    "profile_card": [
        _field("name", "Nombre", "text", required=True, max_chars_key="title_max_chars"),
        _field("role", "Rol", "text", required=True),
        _field("years_experience", "Anos de experiencia", "number", required=True),
        _field("seniority", "Seniority", "text", required=True),
        _field("skills", "Skills", "list", required=True),
        _field("highlights", "Logros destacados", "list"),
        _field("photo", "Ruta de foto (opcional)", "text"),
    ],
    "stat_callout": [
        _field("big_number", "Numero grande", "text", required=True),
        _field("label", "Descripcion", "text", required=True, max_chars_key="title_max_chars"),
        _field("context", "Contexto", "text"),
        _field("source", "Fuente", "text"),
    ],
    "closing": [
        _field("headline", "Titular", "text", required=True, max_chars_key="title_max_chars"),
        _field("contact_name", "Nombre de contacto", "text", required=True),
        _field("contact_email", "Email de contacto", "text", required=True),
        _field("contact_phone", "Telefono", "text"),
        _field("cta", "Llamado a la accion", "text"),
    ],
    "grafico": [
        _field("title", "Titulo", "text", required=True, max_chars_key="title_max_chars"),
        _field("chart_title", "Titulo del grafico", "text"),
        _field("chart_type", "Tipo de grafico", "select", options=["column", "bar", "line", "pie"]),
        _field("categories", "Categorias", "list", required=True),
        _field("series", "Series", "chart_series", required=True),
        _field("source", "Fuente", "text"),
    ],
}


# ──────────────────────────────────── paginas ─────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/slide-types")
def api_slide_types():
    return jsonify(SLIDE_ORDER)


@app.route("/api/schema")
def api_schema():
    """Esquema de campos por slide. Cada campo lleva un flag 'hidden' segun si
    su elemento esta visible en config/generated/layout.yaml en este momento.
    Un campo oculto/reemplazado en el editor visual no tiene ningun efecto en
    el deck generado — el frontend no le muestra un input. Los campos se
    devuelven TODOS (no se eliminan del esquema) para que el frontend pueda
    seguir distinguiendolos de campos dinamicos genuinos (formas nuevas sin
    correspondencia en el modelo) — ver extraFieldNames() en app.js.
    """
    layout_data = _load_layout_yaml()
    result: dict[str, list[dict[str, Any]]] = {}
    for slide_type, fields in FIELD_SCHEMAS.items():
        section = layout_data.get(slide_type, {})
        if not isinstance(section, dict):
            section = {}
        result[slide_type] = [
            {**f, "hidden": not is_field_visible(slide_type, f["name"], section)} for f in fields
        ]
    return jsonify(result)


@app.route("/api/limits")
def api_limits():
    from src.theme.theme_loader import ThemeLoader

    try:
        theme = ThemeLoader().load(THEME_YAML)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500
    return jsonify(theme.limits.model_dump())


@app.route("/api/canvas-layout/<slide_type>")
def api_canvas_layout(slide_type: str):
    """Posicion/estilo resuelto de cada elemento de un tipo de slide, para la
    vista previa en vivo del canvas del editor web — tanto elementos canonicos
    (_D del renderer, reutilizando la MISMA logica que _p() ya usa: {**_D[el],
    **layout_config.pos(...)}) como elementos 'extra' agregados a mano en el
    editor visual y NO presentes en _D (ej. un usuario que reemplazo por
    completo el diseño canonico de 'cover' con sus propias formas nombradas
    'Text 1', 'Image 0', etc. — ver render_layout_extras() en helpers.py, que
    hace exactamente este mismo canonico+extra al generar el deck real).
    Sin esto, un deck cuyo diseño visual reemplazo todos los elementos
    canonicos (caso real, no hipotetico — ver config/generated/layout.yaml)
    mostraria un canvas vacio aunque el deck generado si tenga contenido.

    'fields' por elemento viene de invertir FIELD_TO_ELEMENTS para elementos
    canonicos (le dice al frontend que valor en vivo del formulario mostrar
    ahi). Los extras no tienen ese mapeo estatico — su contenido es el texto
    de muestra capturado del editor visual (mismo valor que usa el deck
    generado salvo que el usuario tambien haya nombrado la forma igual a un
    campo real del modelo, caso dinamico que el formulario ya soporta por
    separado via los campos 'dinamicos' del esquema).
    """
    defaults = SLIDE_DEFAULTS.get(slide_type)
    if defaults is None:
        return jsonify({"error": f"tipo de slide desconocido: {slide_type}"}), 404

    element_to_fields: dict[str, list[str]] = {}
    for field, elements in FIELD_TO_ELEMENTS.get(slide_type, {}).items():
        for el in elements:
            element_to_fields.setdefault(el, []).append(field)

    result: dict[str, dict] = {}
    for el, base in defaults.items():
        override = layout_config.pos(slide_type, el)
        if override is None:
            continue
        result[el] = {**base, **override, "fields": element_to_fields.get(el, [])}

    section = _load_layout_yaml().get(slide_type, {})
    if isinstance(section, dict):
        for name, props in section.items():
            if name.startswith("_") or name in defaults or name == "background_color":
                continue
            if not isinstance(props, dict) or props.get("hidden"):
                continue
            try:
                left, top = float(props["left"]), float(props["top"])
                width, height = float(props["width"]), float(props["height"])
            except (KeyError, TypeError, ValueError):
                continue
            result[name] = {**props, "left": left, "top": top, "width": width, "height": height, "fields": []}

    return jsonify(result)


# ───────────────────────────────── deck data ──────────────────────────────────

def _resolve_deck_file(filename: str) -> Path | None:
    """Resuelve un nombre de archivo para LECTURA, restringido a GENERATED_DIR
    o EXAMPLES_DIR. Usa solo el nombre base (Path.name) y verifica que el
    resultado final siga dentro del directorio esperado — sin esto, un nombre
    como 'C:/Windows/win.ini' o '../../etc/passwd' como query param 'file'
    permitia leer archivos arbitrarios fuera de examples/ (Path("x") / ruta
    absoluta IGNORA "x" y devuelve la ruta absoluta tal cual en Python).
    """
    if not filename:
        return None
    if filename.startswith("generated/"):
        base = GENERATED_DIR
        rel_name = Path(filename.split("/", 1)[1]).name
    else:
        base = EXAMPLES_DIR
        rel_name = Path(filename).name
    if not rel_name:
        return None
    candidate = (base / rel_name).resolve()
    try:
        candidate.relative_to(base.resolve())
    except ValueError:
        return None
    return candidate


@app.route("/api/deck-files")
def api_deck_files():
    files = []
    if TEMPLATE_YAML.exists():
        files.append(f"generated/{TEMPLATE_YAML.name}")
    if EXAMPLES_DIR.exists():
        files.extend(sorted(p.name for p in EXAMPLES_DIR.glob("*.yaml")))
    return jsonify(files)


@app.route("/api/deck-data")
def api_get_deck_data():
    filename = request.args.get("file", "")
    path = _resolve_deck_file(filename)
    if not path or not path.exists():
        return jsonify({"error": "archivo no encontrado"}), 404
    try:
        parsed = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    except yaml.YAMLError as exc:
        return jsonify({"error": f"YAML invalido: {exc}"}), 400
    slides = parsed.get("slides", []) if isinstance(parsed, dict) else []
    return jsonify({"filename": filename, "slides": slides})


@app.route("/api/deck-data", methods=["POST"])
def api_save_deck_data():
    body = request.get_json(force=True, silent=True) or {}
    raw_filename = body.get("filename", "").strip()
    slides = body.get("slides", [])
    if not raw_filename:
        return jsonify({"ok": False, "error": "falta el nombre del archivo"}), 400

    # Solo el nombre base — descarta cualquier componente de ruta (absoluta o
    # con '..') que intente escribir fuera de examples/ (path traversal).
    filename = Path(raw_filename).name
    if not filename:
        return jsonify({"ok": False, "error": "nombre de archivo invalido"}), 400
    if not filename.endswith((".yaml", ".yml")):
        filename += ".yaml"
    if not isinstance(slides, list) or not slides:
        return jsonify({"ok": False, "error": "el deck debe tener al menos 1 slide"}), 400
    if filename == TEMPLATE_YAML.name:
        return jsonify({
            "ok": False,
            "error": (
                f"'{filename}' es el nombre de la plantilla auto-generada — se sobreescribe "
                "cada vez que corres 'Aplicar cambios del editor visual'. Elige otro nombre "
                "para tus datos reales (ej. 'propuesta_cliente.yaml')."
            ),
        }), 400

    content = yaml.dump(
        {"slides": slides}, default_flow_style=False, allow_unicode=True, sort_keys=False
    )
    EXAMPLES_DIR.mkdir(parents=True, exist_ok=True)
    (EXAMPLES_DIR / filename).write_text(content, encoding="utf-8")
    return jsonify({"ok": True, "filename": filename})


# ──────────────────────────────────── render ───────────────────────────────────

def _build_registry():  # type: ignore[no-untyped-def]
    from src.layout.registry import LayoutRegistry
    from src.layout.renderers.closing import ClosingRenderer
    from src.layout.renderers.content_one_col import ContentOneColRenderer
    from src.layout.renderers.content_two_col import ContentTwoColRenderer
    from src.layout.renderers.cover import CoverRenderer
    from src.layout.renderers.grafico import GraficoRenderer
    from src.layout.renderers.pricing_table import PricingTableRenderer
    from src.layout.renderers.profile_card import ProfileCardRenderer
    from src.layout.renderers.section_divider import SectionDividerRenderer
    from src.layout.renderers.stat_callout import StatCalloutRenderer

    registry = LayoutRegistry()
    registry.register("cover", CoverRenderer)
    registry.register("section_divider", SectionDividerRenderer)
    registry.register("content_two_col", ContentTwoColRenderer)
    registry.register("content_one_col", ContentOneColRenderer)
    registry.register("pricing_table", PricingTableRenderer)
    registry.register("profile_card", ProfileCardRenderer)
    registry.register("stat_callout", StatCalloutRenderer)
    registry.register("closing", ClosingRenderer)
    registry.register("grafico", GraficoRenderer)
    return registry


@app.route("/api/render", methods=["POST"])
def api_render():
    body = request.get_json(force=True, silent=True) or {}
    slides = body.get("slides", [])
    output_name = (body.get("output_name") or "preview").strip()
    if not output_name.endswith(".pptx"):
        output_name += ".pptx"
    output_name = Path(output_name).name  # evitar path traversal

    raw = {"slides": slides}
    try:
        from pptx import Presentation

        from src.exporter.pptx_exporter import PptxExporter
        from src.parser.deck_parser import DeckParser
        from src.renderer.deck_renderer import DeckRenderer
        from src.theme.theme_loader import ThemeLoader
        from src.validator.deck_validator import DeckValidator

        theme = ThemeLoader().load(THEME_YAML)
        DeckValidator().validate(raw, dict(theme.limits.model_dump()))
        deck = DeckParser().parse(raw)

        template_path = ROOT / theme.assets.template
        if template_path.exists():
            presentation = Presentation(str(template_path))
        else:
            from pptx.util import Inches

            presentation = Presentation()
            presentation.slide_width = Inches(13.33)
            presentation.slide_height = Inches(7.5)

        registry = _build_registry()
        presentation = DeckRenderer().render(deck, theme, registry, presentation)

        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        out_path = OUTPUT_DIR / output_name
        PptxExporter().export(presentation, out_path)

        return jsonify({"ok": True, "download_url": f"/api/download/{out_path.name}"})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.route("/api/download/<filename>")
def api_download(filename: str):
    path = OUTPUT_DIR / Path(filename).name
    if not path.exists():
        return jsonify({"error": "archivo no encontrado"}), 404
    return send_file(path, as_attachment=True)


if __name__ == "__main__":
    print("Softgic Deck Generator — editor de contenido web")
    print("Abriendo en http://127.0.0.1:5000")
    app.run(debug=True, port=5000)
