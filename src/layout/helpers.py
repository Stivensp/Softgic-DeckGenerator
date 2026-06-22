from __future__ import annotations

import logging
from pathlib import Path
from typing import TYPE_CHECKING, Any

from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

if TYPE_CHECKING:
    from pptx.shapes.autoshape import Shape
    from pptx.slide import Slide
    from pptx.text.text import TextFrame

logger = logging.getLogger("softgic.layout.helpers")

# Raiz del proyecto: src/layout/helpers.py → src/layout → src → project root
_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent

# Superficie total del slide en pulgadas cuadradas (13.33" × 7.5")
_SLIDE_AREA = 13.33 * 7.5

# MSO_AUTO_SHAPE_TYPE integer values (confirmado: OVAL=9, RECTANGLE=1)
_RECTANGLE = 1
_OVAL      = 9

# Palabras clave en el nombre de un shape que indican que es un ovalo,
# usadas como fallback cuando 'kind' no esta almacenado en layout.yaml.
_OVAL_NAME_HINTS = ("oval", "ellipse", "circle", "circulo", "elipse")


def hex_to_rgb(hex_color: str) -> RGBColor:
    """Convierte hex string (con o sin '#') a RGBColor. Devuelve negro si el formato es invalido."""
    try:
        h = str(hex_color).strip().lstrip("#")
        if len(h) != 6:
            raise ValueError(f"longitud invalida: {len(h)} chars")
        return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except (ValueError, IndexError, AttributeError) as exc:
        logger.warning("Color hex invalido '%s' (%s) — se usara negro como fallback", hex_color, exc)
        return RGBColor(0, 0, 0)


def set_slide_background(slide: "Slide", color_hex: str) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = hex_to_rgb(color_hex)


def set_slide_background_gradient(slide: "Slide", angle: int, stops: list[dict[str, Any]]) -> None:
    """Aplica fondo degradado lineal al slide via XML directo.

    angle: grados (0=izq-der, 90=arriba-abajo). stops: [{pos:0-100, color:"RRGGBB"}, ...]
    """
    import lxml.etree as _etree
    from pptx.oxml.ns import qn

    # Crear p:bg / p:bgPr si no existe
    slide.background.fill.solid()

    bg_pr = slide._element.find('.//' + qn('p:bgPr'))
    if bg_pr is None:
        return

    # Limpiar fill existente
    for tag in (qn('a:solidFill'), qn('a:gradFill'), qn('a:noFill'), qn('a:pattFill'), qn('a:blipFill')):
        el = bg_pr.find(tag)
        if el is not None:
            bg_pr.remove(el)

    grad_fill = _etree.Element(qn('a:gradFill'))
    gs_lst = _etree.SubElement(grad_fill, qn('a:gsLst'))
    for stop in stops:
        pos_pct = int(stop.get("pos", 0))
        color   = str(stop.get("color", "000000")).upper().lstrip("#")
        gs = _etree.SubElement(gs_lst, qn('a:gs'))
        gs.set("pos", str(pos_pct * 1000))   # 0-100 → 0-100000
        srgb = _etree.SubElement(gs, qn('a:srgbClr'))
        srgb.set("val", color)

    lin = _etree.SubElement(grad_fill, qn('a:lin'))
    lin.set("ang", str(angle * 60000))        # grados → 1/60000 de grado
    lin.set("scaled", "0")

    bg_pr.insert(0, grad_fill)               # fill debe ir primero en p:bgPr


def add_colored_box(
    slide: "Slide",
    left_in: float,
    top_in: float,
    width_in: float,
    height_in: float,
    fill_hex: str,
    no_line: bool = True,
) -> "Shape":
    shape = slide.shapes.add_shape(
        _RECTANGLE,
        Inches(left_in),
        Inches(top_in),
        Inches(width_in),
        Inches(height_in),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_to_rgb(fill_hex)
    if no_line:
        shape.line.fill.background()
    return shape


def add_text_box(
    slide: "Slide",
    left_in: float,
    top_in: float,
    width_in: float,
    height_in: float,
    text: str,
    font_family: str,
    font_size: int,
    color_hex: str,
    bold: bool = False,
    align: PP_ALIGN = PP_ALIGN.LEFT,
    wrap: bool = True,
    italic: bool = False,
    margins: "dict[str, float] | None" = None,
) -> "Shape":
    txBox = slide.shapes.add_textbox(
        Inches(left_in), Inches(top_in), Inches(width_in), Inches(height_in)
    )
    tf = txBox.text_frame
    tf.word_wrap = wrap
    _apply_margins(tf, margins)
    _apply_text(tf, text, font_family, font_size, color_hex, bold, align, italic)
    return txBox


def _apply_margins(tf: "TextFrame", margins: "dict[str, float] | None") -> None:
    """Aplica los margenes internos (insets) capturados del shape original.
    Sin esto, python-pptx usa sus defaults (~0.1in L/R, 0.05in T/B), que
    casi siempre son mas anchos que los del diseño original — el texto
    envuelve distinto y ya no encaja en la misma caja/proporcion."""
    if not margins:
        return
    try:
        if "margin_left" in margins:
            tf.margin_left = Inches(margins["margin_left"])
        if "margin_right" in margins:
            tf.margin_right = Inches(margins["margin_right"])
        if "margin_top" in margins:
            tf.margin_top = Inches(margins["margin_top"])
        if "margin_bottom" in margins:
            tf.margin_bottom = Inches(margins["margin_bottom"])
    except Exception:
        pass


def set_shape_text(
    shape: "Shape",
    text: str,
    font_family: str,
    font_size: int,
    color_hex: str,
    bold: bool = False,
    align: PP_ALIGN = PP_ALIGN.LEFT,
    italic: bool = False,
) -> None:
    tf = shape.text_frame
    tf.word_wrap = True
    _apply_text(tf, text, font_family, font_size, color_hex, bold, align, italic)


def _write_text_into_shape(
    shape: "Shape",
    props: dict,
    font_family: str,
    name: str = "",
    model: object = None,
    extra_data: "dict[str, object] | None" = None,
) -> None:
    """Escribe el texto (o texto multicolor via 'runs') guardado en props
    directamente en el text_frame de un shape de relleno ya creado — para
    formas que en PowerPoint tienen color de fondo Y texto a la vez (ej. un
    boton/CTA solido con una etiqueta). No-op si props no tiene texto.

    Resuelve dato dinamico igual que el camino kind:"text" — sin esto, una
    forma de color nombrada igual a un campo del modelo (ej. 'cta') nunca
    podia tomar el valor real del YAML de entrada, solo el texto de muestra.
    """
    if not props.get("text") and not props.get("runs"):
        return
    tf = shape.text_frame
    tf.word_wrap = True
    margins = {
        k: props[k] for k in ("margin_left", "margin_right", "margin_top", "margin_bottom")
        if k in props
    }
    _apply_margins(tf, margins)
    tf.clear()
    para = tf.paragraphs[0]

    dynamic_value = _resolve_dynamic_value(model, name, extra_data) if name else None
    if dynamic_value is not None:
        run = para.add_run()
        run.text = str(dynamic_value)
        run.font.name = font_family
        run.font.size = Pt(int(props.get("font_size", 12)))
        if props.get("bold") is not None:
            run.font.bold = bool(props["bold"])
        clr = str(props.get("color", "000000")).lstrip("#")
        run.font.color.rgb = hex_to_rgb(clr)
        return

    runs_data = props.get("runs")
    if runs_data:
        default_sz = int(props.get("font_size", 12))
        for run_data in runs_data:
            r = para.add_run()
            r.text = str(run_data.get("text", ""))
            r.font.name = font_family
            r.font.size = Pt(int(run_data.get("font_size", default_sz)))
            if run_data.get("bold") is not None:
                r.font.bold = bool(run_data["bold"])
            clr = str(run_data.get("color", "000000")).lstrip("#")
            r.font.color.rgb = hex_to_rgb(clr)
    else:
        run = para.add_run()
        run.text = str(props.get("text", ""))
        run.font.name = font_family
        run.font.size = Pt(int(props.get("font_size", 12)))
        if props.get("bold") is not None:
            run.font.bold = bool(props["bold"])
        clr = str(props.get("color", "000000")).lstrip("#")
        run.font.color.rgb = hex_to_rgb(clr)


def _apply_text(
    tf: "TextFrame",
    text: str,
    font_family: str,
    font_size: int,
    color_hex: str,
    bold: bool,
    align: PP_ALIGN,
    italic: bool,
) -> None:
    tf.clear()
    para = tf.paragraphs[0]
    para.alignment = align
    run = para.add_run()
    run.text = text
    run.font.name = font_family
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = hex_to_rgb(color_hex)


def add_bullet_list(
    slide: "Slide",
    left_in: float,
    top_in: float,
    width_in: float,
    height_in: float,
    items: list[str],
    font_family: str,
    font_size: int,
    color_hex: str,
    bullet_char: str = "•",
) -> None:
    txBox = slide.shapes.add_textbox(
        Inches(left_in), Inches(top_in), Inches(width_in), Inches(height_in)
    )
    tf = txBox.text_frame
    tf.word_wrap = True

    for i, item in enumerate(items):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = PP_ALIGN.LEFT
        run = para.add_run()
        run.text = f"{bullet_char}  {item}"
        run.font.name = font_family
        run.font.size = Pt(font_size)
        run.font.color.rgb = hex_to_rgb(color_hex)


def add_image(
    slide: "Slide",
    image_path: str,
    left_in: float,
    top_in: float,
    width_in: float,
    height_in: float,
    fallback_path: str | None = None,
) -> None:
    path = Path(image_path)
    if not path.exists():
        if fallback_path and Path(fallback_path).exists():
            logger.warning("Image '%s' not found, using placeholder: '%s'", image_path, fallback_path)
            path = Path(fallback_path)
        else:
            logger.warning("Image '%s' not found and no fallback available — skipping", image_path)
            return

    slide.shapes.add_picture(
        str(path),
        Inches(left_in),
        Inches(top_in),
        Inches(width_in),
        Inches(height_in),
    )


def add_oval(
    slide: "Slide",
    left_in: float,
    top_in: float,
    width_in: float,
    height_in: float,
    fill_hex: str,
) -> "Shape":
    shape = slide.shapes.add_shape(
        _OVAL,
        Inches(left_in),
        Inches(top_in),
        Inches(width_in),
        Inches(height_in),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_to_rgb(fill_hex)
    shape.line.fill.background()
    return shape


def _patch_shape_prst(shape: "Shape", prst: str) -> None:
    """Cambia el preset geometry de un shape directamente en su XML.

    Permite reproducir cualquier forma de PowerPoint (snip corners, estrellas de Office
    2019+, etc.) aunque python-pptx no la tenga en su enum. PowerPoint usa los valores
    de ajuste (avLst) por defecto del prst indicado cuando avLst esta vacio.
    """
    try:
        from pptx.oxml.ns import qn
        prstGeom = shape._element.spPr.find(qn('a:prstGeom'))
        if prstGeom is not None:
            prstGeom.set('prst', prst)
            avLst = prstGeom.find(qn('a:avLst'))
            if avLst is not None:
                avLst.clear()
    except Exception as exc:
        logger.warning("No se pudo aplicar prst '%s': %s — shape queda como rectangulo", prst, exc)


def add_auto_shape(
    slide: "Slide",
    shape_id: "int | str",
    left_in: float,
    top_in: float,
    width_in: float,
    height_in: float,
    fill_hex: str,
) -> "Shape":
    """Crea cualquier auto shape de PowerPoint.

    shape_id puede ser:
      - int  → MSO_AUTO_SHAPE_TYPE ID (p.ej. 32 para estrella de 5 puntas)
      - str  → prst name del XML (p.ej. 'snip2DiagRect', 'star5') para formas
               cuyo ID no esta en el enum de python-pptx

    Fallback a rectangulo si el tipo no puede crearse.
    """
    try:
        if isinstance(shape_id, int):
            shape = slide.shapes.add_shape(
                shape_id,
                Inches(left_in), Inches(top_in), Inches(width_in), Inches(height_in),
            )
        else:
            # Crear rectangulo y parchear prst para el tipo deseado
            shape = slide.shapes.add_shape(
                _RECTANGLE,
                Inches(left_in), Inches(top_in), Inches(width_in), Inches(height_in),
            )
            _patch_shape_prst(shape, str(shape_id))
    except Exception:
        logger.warning("No se pudo crear shape '%s' — usando rectangulo como fallback", shape_id)
        shape = slide.shapes.add_shape(
            _RECTANGLE,
            Inches(left_in), Inches(top_in), Inches(width_in), Inches(height_in),
        )
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_to_rgb(fill_hex)
    shape.line.fill.background()
    return shape


# Nombres que nunca se resuelven como campo dinamico aunque coincidan con el
# nombre de una forma personalizada — "type" es el discriminador del modelo
# pydantic y nombrar una forma asi corromperia la entrada del deck.
_RESERVED_DYNAMIC_NAMES = frozenset({"type"})


def _resolve_dynamic_value(model: object, name: str, extra_data: "dict[str, object] | None") -> object | None:
    """Resuelve el valor real de un campo dinamico por nombre de forma.

    Busca primero un atributo DECLARADO del modelo pydantic (ej. 'cta' en
    ClosingSlide, que no es un elemento posicional pero si un campo real) —
    esto cubre el caso de nombrar una forma igual a un campo del modelo que
    no esta en _D. Si no es un atributo declarado, cae a model_extra (campos
    no declarados, ej. un nombre de forma totalmente nuevo como 'promo_code').
    Solo se usan valores escalares (str/int/float); listas/objetos anidados
    no tienen una representacion de texto sensata y se ignoran.
    """
    if name in _RESERVED_DYNAMIC_NAMES:
        return None
    if model is not None and hasattr(model, name):
        value = getattr(model, name)
        if isinstance(value, (str, int, float)) and not isinstance(value, bool):
            return value
        return None
    if extra_data:
        return extra_data.get(name)
    return None


def render_layout_extras(
    slide: "Slide",
    slide_type: str,
    known_keys: "set[str]",
    default_fill: str,
    font_family: str = "Calibri",
    model: object = None,
) -> None:
    """Dibuja en el slide los elementos de layout.yaml que NO estan en _D del renderer.

    Permite que formas personalizadas (ovaIos, cajas decorativas) agregadas en el
    editor visual aparezcan en los decks generados.

    model: la slide pydantic original. Si una forma de texto o imagen personalizada
    tiene el mismo nombre que un campo del modelo (declarado o extra), su contenido
    viene del YAML de entrada en vez del texto/imagen de muestra capturado en el
    editor visual — asi el dato es dinamico por deck. Ver _resolve_dynamic_value().
    """
    extra_data = getattr(model, "model_extra", None) if model is not None else None
    try:
        from src.layout import layout_config as lc
        entries = lc._load().get(slide_type, {})
    except Exception as exc:
        logger.warning("render_layout_extras: no se pudo cargar layout.yaml (%s) — se omiten extras", exc)
        return

    if not isinstance(entries, dict):
        return

    # Fondo del slide: leido de Format Background en el editor visual
    bg_override = entries.get("_background")
    if bg_override:
        try:
            if isinstance(bg_override, str) and len(bg_override) == 6:
                # Formato legacy (string hex directo)
                set_slide_background(slide, bg_override)
            elif isinstance(bg_override, dict):
                bg_type = bg_override.get("type", "solid")
                if bg_type == "solid":
                    color = bg_override.get("color", "")
                    if color:
                        set_slide_background(slide, color)
                elif bg_type == "gradient":
                    set_slide_background_gradient(
                        slide,
                        int(bg_override.get("angle", 90)),
                        bg_override.get("stops", []),
                    )
        except Exception as exc:
            logger.warning("No se pudo aplicar fondo del slide: %s", exc)

    for name, props in entries.items():
        if name.startswith("_"):
            continue  # claves internas (_background, etc.)
        if name in known_keys:
            continue  # ya lo dibuja el renderer principal
        if name == "background_color":
            continue  # elemento legacy del editor visual anterior

        if not isinstance(props, dict) or props.get("hidden"):
            continue  # no existe, fue eliminado, o valor invalido

        # Validar que tiene posicion completa con valores numericos
        try:
            left   = float(props["left"])
            top    = float(props["top"])
            width  = float(props["width"])
            height = float(props["height"])
        except (KeyError, TypeError, ValueError):
            logger.warning("Elemento extra '%s' en '%s' tiene posicion incompleta/invalida — omitido", name, slide_type)
            continue

        if width <= 0 or height <= 0:
            logger.warning("Elemento extra '%s' en '%s' tiene dimensiones invalidas (%.3f x %.3f) — omitido",
                           name, slide_type, width, height)
            continue

        fill = props.get("fill") or default_fill
        kind = props.get("kind", "box")
        rotation = props.get("rotation")

        # Fallback por nombre: si 'kind' no esta almacenado (YAML de version antigua)
        # y el nombre del shape sugiere ovalo, tratarlo como tal.
        # Solo aplica cuando kind == "box" para no sobreescribir "shape" o "oval" explicitos.
        if kind == "box" and any(hint in name.lower() for hint in _OVAL_NAME_HINTS):
            kind = "oval"

        try:
            if kind == "image":
                # Imagen dinamica: si el nombre de la forma coincide con un campo
                # del modelo (ej. 'client_logo': 'assets/clientes/acme.png' en el
                # YAML de entrada), usar esa ruta en vez de la imagen de muestra
                # capturada en el editor visual. Si no existe en disco, degrada
                # a la imagen estatica sin romper el render.
                src = ""
                dynamic_src = _resolve_dynamic_value(model, name, extra_data)
                if dynamic_src:
                    candidate = Path(str(dynamic_src))
                    candidate = candidate if candidate.is_absolute() else _PROJECT_ROOT / candidate
                    if candidate.exists():
                        src = str(dynamic_src)
                    else:
                        logger.warning(
                            "Imagen dinamica '%s' para '%s' no existe — usando imagen de muestra",
                            dynamic_src, name,
                        )
                if not src:
                    src = props.get("src", "")
                if not src:
                    logger.warning("Imagen '%s' en '%s' no tiene campo 'src' — omitida", name, slide_type)
                    continue
                img_path = Path(src) if Path(src).is_absolute() else _PROJECT_ROOT / src
                if not img_path.exists():
                    logger.warning("Imagen no encontrada: '%s' — omitida", src)
                    continue
                pic = slide.shapes.add_picture(
                    str(img_path), Inches(left), Inches(top), Inches(width), Inches(height)
                )
                if rotation:
                    pic.rotation = rotation
                # Imagenes que cubren >55% del slide van al fondo (detras de otros shapes)
                if width * height > _SLIDE_AREA * 0.55:
                    sp_tree = slide.shapes._spTree
                    sp_tree.remove(pic._element)
                    sp_tree.insert(2, pic._element)
            elif kind == "text":
                # Margenes internos capturados del shape original (ver
                # apply_visual_config.py::_read_text_margins) — sin esto el
                # texto envuelve distinto y no encaja en la misma caja.
                margins = {
                    k: props[k] for k in ("margin_left", "margin_right", "margin_top", "margin_bottom")
                    if k in props
                }

                dynamic_value = _resolve_dynamic_value(model, name, extra_data)
                if dynamic_value is not None:
                    # Dato dinamico del YAML de entrada (resuelto por nombre de forma)
                    add_text_box(
                        slide, left, top, width, height,
                        str(dynamic_value),
                        font_family,
                        int(props.get("font_size", 12)),
                        str(props.get("color", "000000")),
                        bold=bool(props.get("bold", False)),
                        margins=margins,
                    )
                    continue
                runs_data = props.get("runs")
                if runs_data:
                    # Texto multicolor: renderizar cada run con su propio color
                    txBox = slide.shapes.add_textbox(
                        Inches(left), Inches(top), Inches(width), Inches(height)
                    )
                    tf = txBox.text_frame
                    tf.word_wrap = True
                    _apply_margins(tf, margins)
                    tf.clear()
                    default_sz = int(props.get("font_size", 12))
                    para = tf.paragraphs[0]
                    for run_data in runs_data:
                        r = para.add_run()
                        r.text = str(run_data.get("text", ""))
                        r.font.name = font_family
                        r.font.size = Pt(int(run_data.get("font_size", default_sz)))
                        if run_data.get("bold") is not None:
                            r.font.bold = bool(run_data["bold"])
                        clr = str(run_data.get("color", "000000")).lstrip("#")
                        r.font.color.rgb = hex_to_rgb(clr)
                else:
                    add_text_box(
                        slide, left, top, width, height,
                        str(props.get("text", "")),
                        font_family,
                        int(props.get("font_size", 12)),
                        str(props.get("color", "000000")),
                        bold=bool(props.get("bold", False)),
                        margins=margins,
                    )
            elif kind == "oval":
                shape_obj = add_oval(slide, left, top, width, height, fill)
                _write_text_into_shape(shape_obj, props, font_family, name, model, extra_data)
            elif kind == "shape":
                shape_type_id = props.get("shape_type_id")
                shape_prst    = props.get("shape_prst")
                if isinstance(shape_type_id, int):
                    shape_obj = add_auto_shape(slide, shape_type_id, left, top, width, height, fill)
                elif shape_prst:
                    shape_obj = add_auto_shape(slide, str(shape_prst), left, top, width, height, fill)
                else:
                    shape_obj = add_colored_box(slide, left, top, width, height, fill)
                _write_text_into_shape(shape_obj, props, font_family, name, model, extra_data)
            else:
                shape_obj = add_colored_box(slide, left, top, width, height, fill)
                _write_text_into_shape(shape_obj, props, font_family, name, model, extra_data)
        except Exception as exc:
            logger.warning("No se pudo renderizar elemento extra '%s' en '%s': %s", name, slide_type, exc)


def truncate_text(text: str, max_chars: int, field_name: str) -> str:
    if len(text) <= max_chars:
        return text
    truncated = text[: max_chars - 1] + "…"
    logger.warning(
        "Text truncated in field '%s': %d → %d chars",
        field_name,
        len(text),
        len(truncated),
    )
    return truncated
