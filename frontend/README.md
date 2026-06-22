# Editor de contenido web

Esto **no** es un editor de diseño/posiciones — eso sigue siendo PowerPoint vía
`tools/open_visual_editor.py` + `tools/apply_visual_config.py`, sin cambios.

Este frontend sirve para llenar y enviar el **contenido** de un deck (títulos,
bullets, filas de precios, datos de contacto, etc.) con un formulario por
campos, en vez de editar YAML a mano. Parte de la estructura real de cada
slide type (espejo de `src/models/slides.py`) y, por defecto, carga el último
`config/generated/input_template.yaml` generado por "Aplicar cambios del
editor visual".

## Cómo correrlo

```bash
pip install -r requirements.txt   # instala flask/flask-cors si no los tienes
python frontend/server.py
```

Abre `http://127.0.0.1:5000`. También se puede lanzar desde el menú principal:
`python menu.py` → `[3] Herramientas de diseño` → `[6] Abrir editor de contenido web`.

## Cómo funciona

- Al abrir, carga automáticamente `config/generated/input_template.yaml` (o el
  primer archivo de `examples/` si no existe) como punto de partida — un slide
  por tipo, con valores de muestra.
- Cada slide es una tarjeta con un selector de tipo y los campos reales de ese
  tipo (los requeridos llevan `*`). El formulario se genera dinámicamente desde
  `/api/schema`, que es un espejo de los modelos pydantic — si agregas un campo
  nuevo a `src/models/slides.py`, agrégalo también a `FIELD_SCHEMAS` en
  `frontend/server.py` para que aparezca en el formulario.
- Puedes reordenar slides (↑/↓), eliminarlos, cambiar su tipo, o agregar
  cuantos quieras del mismo tipo (ej. varios `section_divider`).
- Los campos con límite de caracteres (`title_max_chars`, etc.) muestran un
  contador en vivo.
- "Guardar datos" escribe el YAML resultante en `examples/<nombre>.yaml` —
  reutilizable después con `python generate.py --input examples/<nombre>.yaml`.
- "Generar deck" corre el pipeline completo (validación → parseo → render →
  export) en memoria y da un link de descarga del `.pptx` en `output/`.

## Arquitectura

- `server.py` — Flask. El esquema de formularios (`FIELD_SCHEMAS`) vive aquí,
  junto con los mismos módulos que usa el CLI (`src/validator`, `src/parser`,
  `src/renderer`, `src/exporter`) — cero lógica de render duplicada.
- `templates/index.html` + `static/app.js` + `static/style.css` — sin
  frameworks, JS plano con `fetch()`. Cada campo se direcciona con un "path"
  tipo `left.bullets.2` o `rows.0.unit_price` para soportar campos anidados
  (columnas de bullets, filas de tabla de precios, series de gráfico).

## Notas

- Servidor de desarrollo (`debug=True`), pensado para uso local.
- No toca `config/generated/layout.yaml` ni `config_visual.pptx` — el flujo de
  diseño visual (PowerPoint) es completamente independiente de este editor.
