# Estructura del Proyecto — Softgic Deck Generator

Guía completa de cada archivo y carpeta: qué hace, por qué existe y cómo encaja en el flujo general.

---

## ¿Qué hace este proyecto?

Convierte un archivo YAML (o JSON) en un archivo PowerPoint `.pptx` con la identidad visual de Softgic. No necesita tener PowerPoint instalado. No tiene interfaz gráfica. Se ejecuta desde la terminal con un solo comando:

```
python generate.py --input mi_deck.yaml --output output/propuesta.pptx
```

O desde el menú interactivo haciendo doble clic en `menu.bat`.

El resultado es un archivo `.pptx` listo para abrir, compartir o presentar.

---

## Árbol de carpetas y archivos

```
SoffgitDeck/
│
├── generate.py                        ← Punto de entrada (CLI)
├── menu.py                            ← Menú interactivo (todas las opciones en un solo lugar)
├── menu.bat                           ← Doble clic para abrir el menú en Windows
├── CODEOWNERS                         ← Quién es responsable de cada parte del proyecto (GitHub)
├── requirements.txt                   ← Dependencias Python
├── pyproject.toml                     ← Metadatos del paquete y configuración de herramientas
│
├── config/
│   ├── theme.yaml                     ← Colores, fuentes y límites de contenido
│   ├── layout.yaml                    ← Posiciones y estilos visuales de cada elemento
│   └── config_visual.pptx             ← Editor visual de diseño (generado por open_visual_editor.py)
│
├── images/                            ← Logos y placeholders de la marca
│   ├── logo_white.png                 ← Logo Softgic para fondos oscuros
│   ├── logo_dark.png                  ← Logo Softgic para fondos claros
│   └── placeholder_profile.png        ← Avatar genérico para perfiles sin foto
│
├── assets/
│   └── template.pptx                  ← Plantilla PowerPoint base (13.33" × 7.5")
│
├── examples/
│   ├── propuesta_comercial.yaml       ← Ejemplo: propuesta con todos los tipos de slide
│   ├── talent_profile.yaml            ← Ejemplo: presentación de equipo técnico
│   └── costeo.yaml                    ← Ejemplo: estimación de costos con gráficos y tablas
│
├── output/                            ← Aquí se guardan los .pptx generados (se crea automáticamente)
│
├── logs/                              ← Un .log por cada ejecución (se crea automáticamente)
│
├── src/                               ← Todo el código fuente
│   ├── exceptions.py                  ← Jerarquía de errores personalizados
│   │
│   ├── input/                         ← Capa 1: lectura del archivo
│   │   ├── file_reader.py
│   │   └── formats.py
│   │
│   ├── validator/                     ← Capa 2: validación en 3 pasadas
│   │   ├── deck_validator.py          ← Orquestador de las 3 pasadas
│   │   ├── schema_rules.py            ← Pasada 1: campos y tipos
│   │   ├── business_rules.py          ← Pasada 2: reglas de negocio
│   │   └── asset_validator.py         ← Pasada 3: imágenes y archivos
│   │
│   ├── parser/                        ← Capa 3: dict → modelos Pydantic
│   │   └── deck_parser.py
│   │
│   ├── models/                        ← Definición de las estructuras de datos
│   │   ├── slides.py                  ← 9 modelos de slide + tipo unión
│   │   ├── deck.py                    ← DeckModel (contenedor) + DeckMetadata
│   │   └── theme.py                   ← ThemeModel (valores visuales tipados)
│   │
│   ├── theme/                         ← Capa 4: carga del theme.yaml
│   │   └── theme_loader.py
│   │
│   ├── layout/                        ← Capa 5: renderizadores de slides
│   │   ├── base_layout.py             ← Clase base abstracta (ABC)
│   │   ├── layout_config.py           ← Lector de layout.yaml (puente entre config y renderers)
│   │   ├── registry.py                ← Registro tipo→renderizador
│   │   ├── helpers.py                 ← Utilidades de dibujo (cajas, texto, imágenes)
│   │   └── renderers/
│   │       ├── cover.py               ← Portada
│   │       ├── section_divider.py     ← Divisor de sección
│   │       ├── content_two_col.py     ← Contenido en dos columnas
│   │       ├── content_one_col.py     ← Contenido en una sola columna
│   │       ├── pricing_table.py       ← Tabla de precios
│   │       ├── profile_card.py        ← Tarjeta de perfil
│   │       ├── stat_callout.py        ← Número destacado
│   │       ├── closing.py             ← Slide de cierre
│   │       └── grafico.py             ← Gráfico nativo (barras, líneas, torta)
│   │
│   ├── renderer/                      ← Capa 6: orquestador de slides
│   │   └── deck_renderer.py
│   │
│   ├── exporter/                      ← Capa 7: escritura del .pptx al disco
│   │   └── pptx_exporter.py
│   │
│   └── logger/                        ← Logging estructurado (consola + archivo)
│       └── deck_logger.py
│
├── tests/                             ← Suite de tests automatizados
│   ├── conftest.py                    ← Fixtures compartidos por todos los tests
│   ├── unit/                          ← Tests por módulo aislado
│   │   ├── test_file_reader.py
│   │   ├── test_theme_loader.py
│   │   ├── test_validator.py
│   │   ├── test_schema_branches.py
│   │   ├── test_parser.py
│   │   ├── test_registry.py
│   │   ├── test_renderer.py
│   │   ├── test_exporter.py
│   │   ├── test_logger.py
│   │   ├── test_asset_validator.py
│   │   └── test_renderers/
│   │       ├── test_cover.py
│   │       └── test_all_renderers.py
│   └── integration/
│       ├── test_full_pipeline.py      ← Pipeline completo de inicio a fin
│       └── test_output_consistency.py ← Mismo input = mismo output
│
└── tools/
    ├── open_visual_editor.py          ← Genera config_visual.pptx para editar el diseño visualmente
    ├── apply_visual_config.py         ← Lee config_visual.pptx y actualiza layout.yaml
    ├── inspect_template.py            ← Muestra los layouts disponibles en template.pptx
    ├── preview_layout.py              ← Vista previa rápida de posiciones
    ├── create_template.py             ← Script de bootstrap original de assets
    └── create_test_assets.py          ← Genera template, logos y avatar con Pillow
```

---

## El flujo completo de principio a fin

Cuando ejecutas `python generate.py --input deck.yaml`, el sistema pasa por **7 capas en secuencia**. Si cualquier capa falla, el proceso se detiene y muestra un error claro. Nunca se escribe un archivo incompleto.

```
YAML/JSON
    ↓
[1] FileReader      — Lee el archivo y devuelve un dict Python puro
    ↓
[2] Validator       — Valida el dict en 3 pasadas (esquema → reglas → assets)
    ↓
[3] Parser          — Convierte el dict a modelos Pydantic tipados
    ↓
[4] ThemeLoader     — Carga config/theme.yaml como objeto inmutable
    ↓
[5] DeckRenderer    — Por cada slide: llama al renderizador correcto
    ↓
[6] (renderers)     — Dibujan el slide usando helpers.py y layout.yaml
    ↓
[7] PptxExporter    — Escribe el .pptx al disco de forma atómica
```

---

## Archivos raíz

### `generate.py` — El punto de entrada

El único archivo que ejecutas directamente desde la terminal. Hace tres cosas:

1. Procesa argumentos de línea de comandos: `--input`, `--output`, `--theme`, `--version`
2. Inicializa el logger para que todo el sistema pueda escribir desde el primer momento
3. Ejecuta el pipeline de 7 pasos en orden, con manejo de errores limpio

La función `_build_registry()` registra los 9 renderizadores disponibles. Cada vez que agregas un nuevo tipo de slide, solo agregas una línea aquí:

```python
registry.register("cover",            CoverRenderer)
registry.register("section_divider",  SectionDividerRenderer)
registry.register("content_two_col",  ContentTwoColRenderer)
registry.register("content_one_col",  ContentOneColRenderer)
registry.register("pricing_table",    PricingTableRenderer)
registry.register("profile_card",     ProfileCardRenderer)
registry.register("stat_callout",     StatCalloutRenderer)
registry.register("closing",          ClosingRenderer)
registry.register("grafico",          GraficoRenderer)
```

### `menu.py` y `menu.bat` — El menú interactivo

`menu.bat` es el punto de entrada para uso cotidiano (doble clic en Windows). Llama a `menu.py`, que presenta un menú numerado con todas las operaciones:

- Generar un deck desde un ejemplo
- Generar un deck propio
- Ejecutar los tests
- Herramientas de diseño (editor visual, aplicar cambios)
- Regenerar assets

No forma parte del pipeline — solo llama a los mismos scripts que usarías desde la terminal.

### `CODEOWNERS` — Responsabilidades del proyecto (GitHub)

Define qué equipo es responsable de revisar cambios en cada carpeta. GitHub lo usa para asignar revisores automáticamente en pull requests.

```
config/theme.yaml y layout.yaml  → platform-team + design-team
assets/                          → design-team
examples/                        → platform-team + comercial-team
src/ y generate.py               → platform-team
```

---

## `config/`

### `config/theme.yaml` — Los valores visuales globales

**Controla el 100% de las decisiones visuales de alto nivel.** Ningún color ni tamaño de fuente está escrito directamente en el código Python — todo viene de aquí.

Se divide en 6 secciones:

**`colors`** — 8 colores con nombre semántico:
- `primary` (`#0A1628`): azul marino oscuro. Fondos de portada, encabezados.
- `secondary` (`#1E3A5F`): azul medio. Fondo de slides de sección.
- `accent` (`#00AEEF`): azul Softgic. Barras decorativas, totales, etiquetas.
- `background` (`#FFFFFF`): blanco. Fondos de slides de contenido.
- `text_dark` / `text_light` / `text_muted`: colores de texto según el fondo.
- `divider` (`#E5E7EB`): gris claro para líneas y filas alternas de tabla.

**`fonts`** — Familia y tamaños:
- `family`: fuente del sistema (`Calibri`).
- `size_heading` (36pt), `size_subheading` (24pt), `size_body` (14pt), `size_caption` (10pt), `size_stat` (72pt).

**`assets`** — Rutas a los archivos de imagen que usa el sistema.

**`limits`** — Límites de contenido que el validador y los renderizadores respetan:
- `title_max_chars` (80): si un título supera esto, se trunca con `…`.
- `bullets_per_column_max` (6): máximo de viñetas por columna.
- `skills_max` (8): máximo de habilidades en una tarjeta de perfil.
- `pricing_rows_max` (10): máximo de filas en la tabla de precios.

**`layout_indices`** — Índice del layout de PowerPoint a usar (el `6` es el layout en blanco estándar).

---

### `config/layout.yaml` — Las posiciones y estilos de cada elemento

**Controla dónde está cada caja y cómo se ve en cada tipo de slide.** Se organiza por tipo de slide y luego por elemento:

```yaml
cover:
  title:  { left: 0.8, top: 2.3, width: 11.5, height: 1.5 }
  logo:   { left: 0.5, top: 0.25, width: 2.2,  height: 0.7 }

content_two_col:
  header_bar:       { left: 0.0, top: 0.0, width: 13.33, height: 1.2 }
  left_col_bullets: { left: 0.5, top: 2.05, width: 5.9,  height: 5.1 }
```

Todas las medidas están en **pulgadas**. El slide mide 13.33" × 7.5".

Además de posición y tamaño, puede guardar overrides de estilo por elemento. Estos los escribe `apply_visual_config.py` después de que editas el editor visual:

```yaml
cover:
  title:      { left: 0.8, top: 2.3, width: 11.5, height: 1.5, font_size: 44, color: "FFFFFF" }
  accent_bar: { left: 0.0, top: 0.0, width: 13.33, height: 0.12, fill: "00AEEF" }
```

Los renderizadores aplican estos overrides con `p.get("font_size", default)`, `p.get("color", default)`, `p.get("fill", default)`. Si no hay override en `layout.yaml`, usan el valor por defecto del código.

---

## `config/` (completo)

### `config/theme.yaml` — Los valores visuales globales
*(Ver sección anterior)*

### `config/layout.yaml` — Las posiciones y estilos de cada elemento
*(Ver sección anterior)*

### `config/config_visual.pptx` — El editor visual
Archivo generado por `tools/open_visual_editor.py`. Contiene un slide por cada tipo de slide con las cajas y textos reales del diseño. **Se abre y edita en PowerPoint** para mover, redimensionar o cambiar el estilo de los elementos. Después, `tools/apply_visual_config.py` lee los cambios y los escribe en `layout.yaml`. Ver la sección `tools/` para el flujo completo.

Va en `config/` porque es una herramienta de configuración del diseño, no un asset que se usa en la generación de decks.

---

## `images/`

Logos y placeholders de la marca Softgic. Estos archivos son los que los renderizadores insertan dentro de los slides. Sus rutas se configuran en `config/theme.yaml` bajo la sección `assets:`.

### `images/logo_white.png` y `images/logo_dark.png`
Logo Softgic con fondo transparente. `logo_white` se usa sobre fondos oscuros (portada, cierre); `logo_dark` sobre fondos claros.

### `images/placeholder_profile.png`
Avatar 500×500px que se usa cuando un slide `profile_card` no especifica una foto.

---

## `assets/`

### `assets/template.pptx`
Presentación PowerPoint vacía (13.33" × 7.5"). El sistema la abre y le agrega slides desde cero. La plantilla solo aporta las dimensiones y los layouts de PowerPoint disponibles. Si se borra, se puede regenerar ejecutando `python tools/create_test_assets.py`.

---

## `examples/`

Archivos YAML de ejemplo listos para generar.

### `examples/propuesta_comercial.yaml`
8 slides que demuestran los tipos principales:
`cover` → `section_divider` → `content_two_col` → `section_divider` → `stat_callout` → `profile_card` → `pricing_table` → `closing`

### `examples/talent_profile.yaml`
8 slides enfocados en presentación de equipo técnico, con múltiples `profile_card`.

### `examples/costeo.yaml`
11 slides de estimación de costos en tres fases. Demuestra los tipos más recientes:
- `content_one_col`: listado de alcance en una columna completa
- `grafico`: gráfico de columnas con costo por fase (5 categorías, 1 serie)
- Tres `pricing_table` (una por fase), usando el campo `unit` en las filas (horas, servidores, TB/año, sprints)
- `stat_callout`: inversión total final

---

## `src/` — El código fuente

### `src/exceptions.py` — Jerarquía de errores

9 excepciones personalizadas con herencia. Garantizan que los errores sean siempre claros y ubicados.

```
SoftgicDeckError          ← Base de todo
├── InputReadError         ← No se pudo leer el archivo de entrada
├── ValidationError        ← Error durante la validación
│   ├── SchemaValidationError   ← Campo faltante, tipo incorrecto
│   ├── BusinessRuleError       ← Regla de negocio violada
│   └── AssetValidationError    ← Foto no existe o es inválida
├── ParseError             ← El dict no pudo convertirse a modelos Pydantic
├── ThemeLoadError         ← Problema con theme.yaml
├── UnregisteredLayoutError ← Tipo de slide sin renderizador registrado
├── RenderError            ← Falló el renderizado de un slide
└── ExportError            ← No se pudo escribir el archivo .pptx
```

Cada error puede llevar el número de slide (`slide_index`) y el campo (`field`) que causó el problema.

---

### `src/input/` — Capa 1: Lectura del archivo

#### `src/input/formats.py`
Enum con dos valores: `YAML` y `JSON`.

#### `src/input/file_reader.py`
`FileReader().read(path)`:
1. Verifica que el archivo exista
2. Detecta el formato por extensión
3. Lee el texto en UTF-8
4. Parsea con `yaml.safe_load()` o `json.loads()`
5. Verifica que el resultado sea un `dict`
6. Devuelve el diccionario Python puro

---

### `src/validator/` — Capa 2: Validación en 3 pasadas

La validación ocurre **antes** de convertir los datos a modelos tipados. Esto garantiza mensajes de error claros en lugar de errores internos difíciles de interpretar.

#### `src/validator/deck_validator.py`
Orquestador. Ejecuta las 3 pasadas en orden y se detiene en el primer error.

#### `src/validator/schema_rules.py` — Pasada 1
Verifica **estructura y tipos**:
- Que exista la clave `slides` como lista no vacía
- Que cada slide tenga un `type` reconocido (de los 9 disponibles)
- Que los campos obligatorios de cada tipo estén presentes y no sean `null`
- Que los campos tengan el tipo correcto (`str`, `int`, `float`, `list`, etc.)
- Para `pricing_table`: que cada fila tenga `description`, `quantity`, `unit_price`, `total`

#### `src/validator/business_rules.py` — Pasada 2
Verifica **reglas de negocio**:
- Títulos no superan `title_max_chars` (80 chars)
- Columnas no superan `bullets_per_column_max` (6) viñetas
- Tabla de precios no supera `pricing_rows_max` (10) filas
- `quantity` en precios ≥ 0
- Perfil no supera `skills_max` (8) habilidades
- `contact_email` tiene formato válido (regex)
- Para `pricing_table`: la suma de todos los `row.total` debe coincidir con el campo `totals` declarado (tolerancia 0.01)
- Para `grafico`: cada serie tiene el mismo número de valores que categorías hay

#### `src/validator/asset_validator.py` — Pasada 3
Solo actúa si hay slides `profile_card` con foto:
- Verifica que el archivo exista
- Abre la imagen con Pillow para confirmar que es válida
- Si la resolución es menor a la mínima, registra una advertencia (no bloquea la generación)

---

### `src/models/` — Definición de estructuras de datos

Los modelos son la "forma" tipada que tienen los datos después de pasar la validación. Usan Pydantic v2.

#### `src/models/slides.py`
Define las **9 clases de slide** y el tipo unión `SlideModel`:

| Clase | Campos obligatorios | Campos opcionales |
|---|---|---|
| `CoverSlide` | `type`, `title` | `subtitle`, `client`, `date`, `author` |
| `SectionDividerSlide` | `type`, `section_number`, `section_title` | `tagline` |
| `ContentTwoColSlide` | `type`, `title`, `left`, `right` | — |
| `ContentOneColSlide` | `type`, `title`, `bullets` | `body_title` |
| `PricingTableSlide` | `type`, `title`, `currency`, `rows`, `totals` | `notes` |
| `ProfileCardSlide` | `type`, `name`, `role`, `years_experience`, `seniority`, `skills` | `highlights`, `photo` |
| `StatCalloutSlide` | `type`, `big_number`, `label` | `context`, `source` |
| `ClosingSlide` | `type`, `headline`, `contact_name`, `contact_email` | `contact_phone`, `cta` |
| `GraficoSlide` | `type`, `title`, `categories`, `series` | `chart_title`, `chart_type`, `source` |

`PricingRow` tiene los campos `description`, `quantity`, `unit_price`, `total` y el campo opcional `unit` para indicar la unidad de medida (horas, días, licencias, etc.).

`SlideModel` es un **tipo unión discriminado**: Pydantic mira el campo `type` y sabe exactamente qué clase usar, sin ambigüedad.

#### `src/models/deck.py`
`DeckModel` es el contenedor: lista de slides más metadatos opcionales. Impide que se cree un deck vacío.

#### `src/models/theme.py`
`ThemeModel` representa `config/theme.yaml` como un objeto **inmutable** (Pydantic `frozen=True`). Una vez cargado, ningún código puede modificar accidentalmente un color o fuente.

---

### `src/theme/` — Capa 4: Carga del tema

#### `src/theme/theme_loader.py`
`ThemeLoader().load(path)`:
1. Lee `config/theme.yaml`
2. Valida con `ThemeModel.model_validate()` (tipos correctos, campos presentes)
3. Devuelve el `ThemeModel` inmutable

---

### `src/layout/` — Los renderizadores

#### `src/layout/layout_config.py` — El puente entre `layout.yaml` y los renderers

Lee `config/layout.yaml` una sola vez (caché en memoria) y provee los datos a los renderizadores:

```python
lc.pos("cover", "title")
# → { left: 0.8, top: 2.3, width: 11.5, height: 1.5 }
# Si hay overrides en layout.yaml: también incluye font_size, color o fill
```

Si `layout.yaml` no existe o el elemento no está definido, devuelve `{}` y los renderers usan sus valores por defecto.

#### `src/layout/base_layout.py`
Clase abstracta con un solo método:
```python
def render(self, slide, model, theme) -> None: ...
```
Todos los renderizadores lo implementan, lo que permite llamarlos sin saber de qué tipo son.

#### `src/layout/registry.py`
Diccionario `str → clase de renderizador`. Si se pide un tipo no registrado, lanza `UnregisteredLayoutError` con la lista de tipos disponibles.

#### `src/layout/helpers.py`
Utilidades de bajo nivel que todos los renderizadores comparten:

| Función | Qué hace |
|---|---|
| `hex_to_rgb("#00AEEF")` | Convierte hex a `RGBColor` de python-pptx |
| `set_slide_background(slide, color_hex)` | Pinta el fondo del slide |
| `add_colored_box(slide, left, top, width, height, color_hex)` | Rectángulo de color |
| `add_text_box(slide, left, top, width, height, text, ...)` | Caja de texto con fuente, tamaño, color |
| `add_bullet_list(slide, left, top, width, height, items, ...)` | Lista de viñetas con `•` |
| `add_image(slide, path, left, top, width, height, fallback)` | Imagen; usa fallback si no existe |
| `truncate_text(text, max_chars, field_name)` | Recorta si supera el límite y registra WARNING |

Todas las medidas van en **pulgadas** porque python-pptx trabaja en EMU y `Inches()` hace la conversión.

---

#### `src/layout/renderers/` — Los 9 renderizadores

Cada renderizador recibe `(slide, model, theme)` y construye el slide usando `helpers.py`. Ningún valor visual está hardcodeado: todo viene de `theme` o de `layout.yaml`.

**Patrón común en todos los renderers:**

Cada renderer tiene un diccionario `_D` con los valores por defecto de cada elemento, y una función `_p()` que los fusiona con los overrides de `layout.yaml` (los overrides tienen prioridad):

```python
_D = {
    "title": {"left": 0.8, "top": 2.3, "width": 11.5, "height": 1.5, "font_size": 40},
}

def _p(el: str) -> dict:
    return {**_D[el], **lc.pos("cover", el)}   # layout.yaml sobreescribe _D

# En el render:
p = _p("title")
add_text_box(slide, p["left"], p["top"], ...,
             p.get("font_size", default),
             p.get("color", default))
```

Esto significa que puedes cambiar posición, tamaño, color o fuente de cualquier elemento editando `layout.yaml` — sin tocar el código Python.

---

**`cover.py` — Portada**
- Fondo `colors.primary` (azul marino oscuro)
- Barra delgada `colors.accent` arriba
- Logo Softgic en esquina superior izquierda
- Título grande, subtítulo, línea divisora, nombre del cliente, autor + fecha

**`section_divider.py` — Divisor de sección**
- Fondo `colors.secondary` (azul medio)
- Barra vertical derecha + línea superior en `colors.accent`
- Número de sección en 72pt, título de sección, tagline opcional en cursiva

**`content_two_col.py` — Contenido en dos columnas**
- Barra de encabezado `colors.primary` con título del slide
- Columna izquierda y derecha: cada una con título opcional y lista de bullets
- Línea divisora vertical central `colors.accent`

**`content_one_col.py` — Contenido en una sola columna**
- Igual estructura que `content_two_col` pero el cuerpo ocupa todo el ancho del slide
- Campo `body_title` opcional para un subtítulo antes de los bullets
- Útil para listados largos, alcances de proyecto, requisitos técnicos

**`pricing_table.py` — Tabla de precios**
- Tabla real de PowerPoint con 5 columnas: Descripción, Cant., Unidad, Precio Unit., Total
- La columna "Unidad" muestra el campo opcional `unit` de cada fila (horas, licencias, sprints, etc.)
- Filas de datos con fondo alternado
- Fila de totales con fondo `colors.accent` y texto blanco
- Valida que la suma de `row.total` coincida con el campo `totals` declarado

**`profile_card.py` — Tarjeta de perfil**
- Panel izquierdo (~4"): foto o avatar, nombre, rol, seniority, años de experiencia
- Panel derecho (~8"): etiquetas de skills como cajas `colors.accent`, lista de highlights

**`stat_callout.py` — Número destacado**
- Fondo `colors.primary`
- Número gigante centrado (72pt) en `colors.accent`
- Label debajo, contexto y fuente opcionales

**`closing.py` — Slide de cierre**
- Fondo `colors.primary` con dos barras `colors.accent` arriba y abajo
- Headline grande centrado
- Bloque de contacto: nombre, email, teléfono opcional
- CTA (call to action) con fondo `colors.accent`

**`grafico.py` — Gráfico nativo**
- Usa python-pptx para insertar un gráfico real (no una imagen estática)
- Tipos soportados vía campo `chart_type` en el YAML: `column` (barras verticales), `bar` (barras horizontales), `line` (líneas), `pie` (torta)
- Las categorías se definen en una lista; las series son pares `{name, values[]}`
- Los colores de series se asignan desde una paleta interna que rota si hay más de 5 series
- Campo `source` opcional para citar la fuente del dato

---

### `src/renderer/` — Capa 6: Orquestador

#### `src/renderer/deck_renderer.py`
Por cada slide en el deck:
1. Obtiene el índice de layout del theme
2. Agrega el slide a la presentación: `presentation.slides.add_slide(layout)`
3. Busca el renderizador: `registry.get(slide_type)` → ej. `GraficoRenderer`
4. Llama: `GraficoRenderer().render(slide, model, theme)`
5. Si falla → `RenderError` con el número de slide y el error original

Devuelve el objeto `Presentation` de python-pptx con todos los slides ya construidos.

---

### `src/exporter/` — Capa 7: Escritura al disco

#### `src/exporter/pptx_exporter.py`
Escribe el archivo final de forma **atómica**: primero guarda en un temporal, luego hace `os.replace(tmp, destino)`. Si algo falla durante la escritura, el archivo anterior no queda corrupto.

---

### `src/logger/` — Logging estructurado

#### `src/logger/deck_logger.py`
- **Consola**: nivel INFO — muestra el progreso al usuario
- **Archivo**: nivel DEBUG — guarda detalles para diagnóstico
- **Nombre del archivo**: `logs/YYYYMMDD_HHMMSS_deck.log` — uno por ejecución, nunca se sobreescriben
- **Formato**: `[2026-06-12 10:33:59] [INFO] [softgic.renderer] Rendering 8 slide(s)...`

---

## `tools/` — Herramientas de utilidad

Scripts auxiliares que no forman parte del pipeline de generación.

### El flujo del editor visual (cómo cambiar el diseño sin tocar código)

```
[1] python tools/open_visual_editor.py
    → Genera config/config_visual.pptx
    → Contiene 9 slides con el diseño real: colores, fuentes y textos de muestra

[2] Abres config_visual.pptx en PowerPoint
    → Mueves cajas, cambias tamaños, ajustas colores o fuentes a gusto
    → Guardas el archivo

[3] python tools/apply_visual_config.py
    → Lee config_visual.pptx shape por shape
    → Extrae posición (left/top/width/height), font_size, color de texto y fill de relleno
    → Actualiza config/layout.yaml con los nuevos valores

[4] La próxima vez que generes un deck
    → Los renderers leen layout.yaml y aplican tus cambios automáticamente
```

---

#### `tools/open_visual_editor.py`
Genera `config/config_visual.pptx`. Cada slide muestra el tipo correspondiente con:
- Rectángulos con los colores reales del tema
- Cajas de texto con textos de muestra, la fuente correcta y los tamaños reales
- Todas las formas nombradas internamente para que `apply_visual_config.py` sepa qué elemento es cada una

#### `tools/apply_visual_config.py`
Lee `config/config_visual.pptx` slide por slide y por cada shape:
- Si es una caja de texto (tipo 17): extrae `left`, `top`, `width`, `height`, `font_size` y `color`
- Si es una forma con relleno (tipo 1): extrae `left`, `top`, `width`, `height` y `fill`

Luego escribe estos valores en `config/layout.yaml`, actualizando solo los elementos que encontró.

#### `tools/inspect_template.py`
Muestra los layouts disponibles en `assets/template.pptx` con su índice y nombre. Útil si cambias la plantilla base y necesitas saber qué número tiene el layout en blanco.

#### `tools/preview_layout.py`
Vista previa rápida de las posiciones definidas en `layout.yaml` para un tipo de slide específico.

#### `tools/create_test_assets.py`
Genera los assets de arranque del proyecto:
- `template.pptx`: presentación base 13.33" × 7.5" con 11 layouts
- `logo_white.png` y `logo_dark.png`: logos placeholder generados con Pillow
- `placeholder_profile.png`: silueta de persona sobre fondo `#1E3A5F`

Se ejecuta **una sola vez** antes de generar decks por primera vez, o cuando se quieren regenerar los assets.

---

## `tests/` — Tests automatizados

### `tests/conftest.py`
Archivo especial de pytest. Define fixtures disponibles en todos los tests sin importarlos:
- `minimal_theme`: un `ThemeModel` completo con valores válidos para tests
- `blank_presentation` y `blank_slide`: objetos de python-pptx listos para usar
- 9 fixtures de slide (uno por tipo), ya construidos con datos válidos
- `full_deck`: un `DeckModel` con los 9 slides para tests de integración
- `minimal_raw_deck` y `full_raw_deck`: diccionarios Python para tests del validador

### `tests/unit/`
Un archivo de test por módulo. Cada test prueba una sola unidad de forma aislada.

| Archivo | Qué prueba |
|---|---|
| `test_file_reader.py` | Lectura de YAML/JSON, extensiones inválidas, archivo vacío |
| `test_theme_loader.py` | Carga de theme.yaml, YAML malformado, campos faltantes |
| `test_validator.py` | Las 3 pasadas de validación: casos válidos e inválidos |
| `test_schema_branches.py` | Ramas específicas del validador de esquema |
| `test_parser.py` | Conversión de dict a DeckModel |
| `test_registry.py` | Registro y obtención de renderizadores |
| `test_renderer.py` | DeckRenderer: presentación con slides, error en renderer |
| `test_exporter.py` | Escritura del .pptx, creación de carpetas, fallo atómico |
| `test_logger.py` | Inicialización del logger, handlers de consola y archivo |
| `test_asset_validator.py` | Foto no encontrada, imagen inválida, resolución baja |
| `test_renderers/test_cover.py` | Renderizador de portada: happy path, campos opcionales, truncado |
| `test_renderers/test_all_renderers.py` | Los demás renderizadores con los mismos patrones |

### `tests/integration/`
**`test_full_pipeline.py`**: ejecuta el pipeline completo con archivos reales y verifica que el `.pptx` resultante sea válido con el número correcto de slides.

**`test_output_consistency.py`**: ejecuta el mismo input dos veces y verifica que los archivos sean idénticos byte a byte. Garantiza que el sistema sea determinista.

---

## Reglas de diseño que guían todo el sistema

**N01 — El template manda**: ningún renderizador inventa layouts. Todo viene de `template.pptx` y los índices en `theme.yaml`.

**N02 — Cero hardcoding visual**: si buscas un color hexadecimal o un tamaño de fuente en cualquier archivo `.py`, no lo encontrarás. Todo pasa por `ThemeModel` o `layout.yaml`.

**N03 — Full validation before render**: el sistema valida completamente antes de tocar python-pptx. Si hay un error, falla rápido con un mensaje claro.

**N04 — Same input = same output**: el sistema es determinista. `test_output_consistency.py` lo garantiza.

**N05 — Every execution must produce logs**: `DeckLogger` escribe un archivo en `logs/` en cada ejecución, sin excepción.

---

## Cómo agregar un nuevo tipo de slide

Se tocan **5 lugares**, sin modificar ningún código existente:

1. **`src/models/slides.py`** — Agregar la clase del modelo y añadirla al `Union` de `SlideModel`
2. **`src/validator/schema_rules.py`** — Agregar el tipo a `REGISTERED_TYPES`, `REQUIRED_FIELDS` y `FIELD_TYPES`
3. **`src/layout/renderers/mi_slide.py`** — Crear el renderizador con su `_D`, su `_p()` y la lógica de render
4. **`generate.py`** — Agregar `registry.register("mi_slide", MiSlideRenderer)`
5. **`config/layout.yaml`** — Agregar la sección con las posiciones por defecto del nuevo tipo
