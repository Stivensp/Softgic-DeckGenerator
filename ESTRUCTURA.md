# Estructura del Proyecto — Softgic Deck Generator

Guía completa de cada archivo y carpeta: qué hace, por qué existe y cómo encaja en el flujo general.

---

## ¿Qué hace este proyecto?

Convierte un archivo YAML (o JSON) en un archivo PowerPoint `.pptx` con la identidad visual de Softgic. No necesita tener PowerPoint instalado. No tiene interfaz gráfica. Se ejecuta desde la terminal con un solo comando:

```
python generate.py --input mi_deck.yaml --output output/propuesta.pptx
```

El resultado es un archivo `.pptx` listo para abrir, compartir o presentar.

---

## Árbol de carpetas y archivos

```
SoffgitDeck/
│
├── generate.py                        ← Punto de entrada (CLI)
├── requirements.txt                   ← Dependencias Python
├── pyproject.toml                     ← Metadatos del paquete y configuración de herramientas
│
├── config/
│   └── theme.yaml                     ← Todos los valores visuales (colores, fuentes, límites)
│
├── assets/
│   ├── template.pptx                  ← Plantilla PowerPoint base (13.33" × 7.5")
│   ├── logo_white.png                 ← Logo Softgic para fondos oscuros
│   ├── logo_dark.png                  ← Logo Softgic para fondos claros
│   └── placeholder_profile.png        ← Avatar genérico para perfiles sin foto
│
├── examples/
│   ├── propuesta_comercial.yaml       ← Ejemplo: propuesta con todos los tipos de slide
│   └── talent_profile.yaml           ← Ejemplo: presentación de equipo técnico
│
├── output/                            ← Aquí se guardan los .pptx generados
│   ├── propuesta_comercial.pptx
│   └── talent_profile.pptx
│
├── logs/                              ← Un .log por cada ejecución (se crea automáticamente)
│   └── 20260612_103359_deck.log
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
│   │   ├── slides.py                  ← 7 modelos de slide + tipo unión
│   │   ├── deck.py                    ← DeckModel (contenedor) + DeckMetadata
│   │   └── theme.py                   ← ThemeModel (valores visuales tipados)
│   │
│   ├── theme/                         ← Capa 4: carga del theme.yaml
│   │   └── theme_loader.py
│   │
│   ├── layout/                        ← Capa 5: renderizadores de slides
│   │   ├── base_layout.py             ← Clase base abstracta (ABC)
│   │   ├── registry.py                ← Registro tipo→renderizador
│   │   ├── helpers.py                 ← Utilidades de dibujo (cajas, texto, imágenes)
│   │   └── renderers/
│   │       ├── cover.py
│   │       ├── section_divider.py
│   │       ├── content_two_col.py
│   │       ├── pricing_table.py
│   │       ├── profile_card.py
│   │       ├── stat_callout.py
│   │       └── closing.py
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
├── tests/                             ← Suite de tests automatizados (121 tests, 99% cobertura)
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
    ├── create_template.py             ← Script original de bootstrap de assets
    └── create_test_assets.py          ← Script mejorado con logo y avatar más visuales
```

---

## El flujo completo de principio a fin

Cuando ejecutas `python generate.py --input deck.yaml`, el sistema pasa por **7 capas en secuencia**. Si cualquier capa falla, el proceso se detiene y muestra un error claro. Nunca se escribe un archivo incompleto.

```
YAML/JSON  →  FileReader  →  Validator  →  Parser  →  ThemeLoader
                                                            ↓
                                                       DeckRenderer
                                                            ↓
                                                       PptxExporter  →  .pptx
```

Cada capa se explica a detalle en las secciones siguientes.

---

## Archivos raíz

### `generate.py` — El punto de entrada

Este es el único archivo que ejecutas directamente. Hace tres cosas:

1. **Procesa los argumentos de la línea de comandos** (`--input`, `--output`, `--theme`, `--version`) usando `argparse`.
2. **Inicializa el logger** para que todos los módulos puedan escribir en consola y en archivo desde el primer momento.
3. **Ejecuta el pipeline de 7 pasos** en orden, dentro de un bloque `try/except` que captura cualquier error y termina con código de salida `1` si algo falla, o `0` si todo fue exitoso.

La función `_build_registry()` registra los 7 renderizadores disponibles. Cada vez que agregas un nuevo tipo de slide, solo agregas una línea aquí.

```python
# Así se registran los renderizadores:
registry.register("cover",            CoverRenderer)
registry.register("section_divider",  SectionDividerRenderer)
registry.register("content_two_col",  ContentTwoColRenderer)
registry.register("pricing_table",    PricingTableRenderer)
registry.register("profile_card",     ProfileCardRenderer)
registry.register("stat_callout",     StatCalloutRenderer)
registry.register("closing",          ClosingRenderer)
```

### `requirements.txt` — Dependencias

Lista las librerías externas que necesita el proyecto:

| Librería | Para qué sirve |
|---|---|
| `python-pptx` | Crear y manipular archivos `.pptx` |
| `Pillow` | Abrir y validar imágenes (fotos de perfil, logos) |
| `pydantic` | Definir y validar modelos de datos tipados |
| `pyyaml` | Leer archivos `.yaml` |
| `pytest` / `pytest-cov` | Tests y reporte de cobertura |
| `black` / `ruff` | Formateo y linting de código |
| `mypy` | Verificación estática de tipos |

### `pyproject.toml` — Metadatos del paquete

Archivo estándar de Python moderno. Contiene:
- El nombre y versión del paquete (`softgic-deck-generator v1.0.0`)
- Que requiere Python 3.11 o superior
- Las mismas dependencias que `requirements.txt`, pero en formato estándar PEP 517
- Configuración de cada herramienta: `black` (largo de línea 100), `ruff` (linting), `mypy` (chequeo de tipos estricto), `pytest` (carpeta `tests/`, output verbose)

---

## `config/`

### `config/theme.yaml` — El cerebro visual

**Este archivo controla el 100% de las decisiones visuales**. Ningún color, tamaño de fuente ni límite está escrito directamente en el código Python. Todo viene de aquí.

Se divide en 6 secciones:

**`colors`** — 8 colores con nombre semántico:
- `primary` (`#0A1628`): azul marino oscuro. Fondos de portada, encabezados de tabla.
- `secondary` (`#1E3A5F`): azul medio. Fondo de slides de sección.
- `accent` (`#00AEEF`): azul Softgic. Barras decorativas, etiquetas, totales.
- `background` (`#FFFFFF`): blanco. Fondos de slides de contenido.
- `text_dark` / `text_light` / `text_muted`: colores de texto según el fondo.
- `divider` (`#E5E7EB`): gris claro para líneas y filas alternas de tabla.

**`fonts`** — Familia y tamaños:
- `family`: fuente del sistema (`Calibri`). Se puede cambiar a `Montserrat` si está instalada.
- `size_heading` (36pt), `size_subheading` (24pt), `size_body` (14pt), `size_caption` (10pt), `size_stat` (72pt — el número grande del stat callout).

**`assets`** — Rutas a los archivos de imagen que usa el sistema:
- `template`: el `.pptx` base
- `logo_white`, `logo_dark`: logos para fondos oscuros/claros
- `placeholder_profile`: avatar cuando no hay foto

**`limits`** — Límites de contenido que el validador y los renderizadores respetan:
- `title_max_chars` (80): si un título supera esto, se trunca con `…` y se registra una advertencia.
- `bullets_per_column_max` (6): máximo de viñetas por columna.
- `skills_max` (8): máximo de habilidades en una tarjeta de perfil.
- `pricing_rows_max` (10): máximo de filas en la tabla de precios.
- `profile_image_min_width/height` (200px): resolución mínima recomendada para fotos.

**`layout_indices`** — Índice del layout de PowerPoint a usar por cada tipo de slide. El índice `6` es el layout "en blanco" estándar que existe en cualquier plantilla de PowerPoint. Si en el futuro Softgic entrega una plantilla corporativa con layouts numerados, solo hay que actualizar estos números aquí.

---

## `assets/`

Archivos estáticos que el sistema usa durante el renderizado.

### `assets/template.pptx`
La presentación PowerPoint vacía que sirve como base. Tiene dimensiones 13.33" × 7.5" (formato 16:9 estándar). Cuando el sistema genera un deck, abre este archivo y le agrega slides uno a uno. Los slides se construyen completamente desde cero en Python — la plantilla solo aporta las dimensiones y los 11 layouts de PowerPoint disponibles.

Se genera ejecutando `python tools/create_test_assets.py`.

### `assets/logo_white.png` y `assets/logo_dark.png`
Imágenes PNG con fondo transparente. La versión `white` se usa sobre fondos oscuros (portada, sección), la versión `dark` sobre fondos claros. Actualmente son placeholders generados con Pillow: un cuadrado azul (`#00AEEF`) con un recorte interior y el wordmark "SOFTGIC".

### `assets/placeholder_profile.png`
Avatar 500×500px con fondo azul medio (`#1E3A5F`), silueta de persona en gris claro y borde circular en azul acento. Se usa cuando un slide `profile_card` no especifica una foto.

---

## `examples/`

Archivos YAML de ejemplo listos para generar.

### `examples/propuesta_comercial.yaml`
8 slides que demuestran todos los tipos disponibles en un solo deck:
`cover` → `section_divider` → `content_two_col` → `section_divider` → `stat_callout` → `profile_card` → `pricing_table` → `closing`

### `examples/talent_profile.yaml`
8 slides enfocados en presentación de equipo técnico, con múltiples `profile_card`.

---

## `src/` — El código fuente

### `src/exceptions.py` — Jerarquía de errores

Define 9 excepciones personalizadas con herencia. Esto permite que el código capture errores a distintos niveles de especificidad y que los mensajes de error al usuario sean siempre claros y ubicados.

```
SoftgicDeckError          ← Base de todo
├── InputReadError         ← No se pudo leer el archivo de entrada
├── ValidationError        ← Error durante la validación
│   ├── SchemaValidationError   ← Campo faltante, tipo incorrecto
│   ├── BusinessRuleError       ← Regla de negocio violada (ej: demasiados bullets)
│   └── AssetValidationError    ← Foto de perfil no existe o es inválida
├── ParseError             ← El dict no pudo convertirse a modelos Pydantic
├── ThemeLoadError         ← Problema con theme.yaml
├── UnregisteredLayoutError ← Se pidió un tipo de slide que no tiene renderizador
├── RenderError            ← Falló el renderizado de un slide
└── ExportError            ← No se pudo escribir el archivo .pptx
```

Cada error puede llevar el número de slide (`slide_index`) y el campo específico (`field`) que causó el problema, para que el mensaje de error sea lo más preciso posible.

---

### `src/input/` — Capa 1: Lectura del archivo

#### `src/input/formats.py`
Enum simple con dos valores: `YAML` y `JSON`. Lo usa `FileReader` para saber cómo parsear el archivo.

#### `src/input/file_reader.py`
La clase `FileReader` con su método `read(file_path)`:

1. Verifica que el archivo exista (si no → `InputReadError`)
2. Detecta el formato por extensión: `.yaml`/`.yml` → YAML, `.json` → JSON (cualquier otra → `InputReadError`)
3. Lee el texto del archivo en UTF-8
4. Verifica que no esté vacío
5. Parsea con `yaml.safe_load()` o `json.loads()`
6. Verifica que el resultado sea un `dict` (no una lista ni un valor simple)
7. Devuelve el `dict` en memoria

Produce: un diccionario Python puro, sin tipos específicos todavía.

---

### `src/validator/` — Capa 2: Validación en 3 pasadas

La validación ocurre **antes** de convertir los datos a modelos tipados. Esto garantiza mensajes de error claros y ubicados, en lugar de errores internos de Pydantic difíciles de interpretar.

#### `src/validator/deck_validator.py`
El `DeckValidator` es el orquestador. No contiene lógica propia: simplemente ejecuta las 3 pasadas en orden. Si cualquiera lanza una excepción, el proceso se detiene ahí.

```python
DeckValidator().validate(raw_dict, limits_dict)
# ↓ ejecuta en orden:
SchemaRulesValidator().validate(raw_dict)
BusinessRulesValidator().validate(raw_dict, limits)
AssetValidator().validate(raw_dict, min_width, min_height)
```

#### `src/validator/schema_rules.py` — Pasada 1
Verifica la **estructura y tipos** del diccionario:
- Que exista la clave `slides` y que sea una lista no vacía
- Que cada slide sea un `dict` con un campo `type` reconocido
- Que los campos obligatorios de cada tipo estén presentes y no sean `null`
- Que los campos tengan el tipo correcto (`title` debe ser `str`, `years_experience` debe ser `int`, `totals` debe ser `int` o `float`, etc.)
- Para `content_two_col`: que cada columna tenga `bullets` no vacío
- Para `pricing_table`: que cada fila tenga los 4 campos requeridos (`description`, `quantity`, `unit_price`, `total`)

#### `src/validator/business_rules.py` — Pasada 2
Verifica **reglas de negocio** que el esquema no puede detectar:
- Que los títulos no superen `title_max_chars` (80 chars por defecto)
- Que las columnas no tengan más de `bullets_per_column_max` (6) viñetas
- Que la tabla de precios no tenga más de `pricing_rows_max` (10) filas
- Que `quantity` en precios sea ≥ 0
- Que el perfil no tenga más de `skills_max` (8) habilidades
- Que `years_experience` sea ≥ 0
- Que `contact_email` tenga formato válido (via regex `^[^@\s]+@[^@\s]+\.[^@\s]+$`)

#### `src/validator/asset_validator.py` — Pasada 3
Solo actúa si hay slides de tipo `profile_card` con una foto:
- Verifica que el archivo de la foto exista
- Abre la imagen con Pillow para confirmar que es una imagen válida (no un PDF disfrazado, por ejemplo)
- Si las dimensiones son menores a `profile_image_min_width` × `profile_image_min_height`, registra una advertencia (pero NO bloquea — es una recomendación, no un error fatal)

---

### `src/models/` — Definición de estructuras de datos

Los modelos son la "forma" tipada que tienen los datos una vez que pasan la validación. Usan Pydantic v2.

#### `src/models/slides.py`
Define las 7 clases de slide y el tipo unión `SlideModel`:

| Clase | Campos obligatorios | Campos opcionales |
|---|---|---|
| `CoverSlide` | `type`, `title` | `subtitle`, `client`, `date`, `author` |
| `SectionDividerSlide` | `type`, `section_number`, `section_title` | `tagline` |
| `ContentTwoColSlide` | `type`, `title`, `left`, `right` | — |
| `PricingTableSlide` | `type`, `title`, `currency`, `rows`, `totals` | `notes` |
| `ProfileCardSlide` | `type`, `name`, `role`, `years_experience`, `seniority`, `skills` | `highlights`, `photo` |
| `StatCalloutSlide` | `type`, `big_number`, `label` | `context`, `source` |
| `ClosingSlide` | `type`, `headline`, `contact_name`, `contact_email` | `contact_phone`, `cta` |

El `SlideModel` al final es un **tipo unión discriminado**: Pydantic mira el campo `type` de cada slide y sabe exactamente qué clase usar sin ambigüedad.

```python
SlideModel = Annotated[
    Union[CoverSlide, SectionDividerSlide, ...],
    Field(discriminator="type"),
]
```

#### `src/models/deck.py`
`DeckModel` es el contenedor principal: una lista de `SlideModel` más metadatos opcionales. El validador de Pydantic integrado impide que se cree un deck vacío.

`DeckMetadata` registra cuándo se generó el deck, cuál fue el archivo de entrada y qué versión de theme se usó. Es puramente informativo.

#### `src/models/theme.py`
`ThemeModel` es el objeto tipado que representa `config/theme.yaml`. Tiene la propiedad `model_config = ConfigDict(frozen=True)` que lo hace **inmutable**: una vez cargado, ningún código puede modificar accidentalmente un color o fuente. Si algo lo intenta, Python lanza un error.

Sus submodelos son:
- `ColorPalette` — los 8 colores
- `FontConfig` — familia y tamaños
- `AssetPaths` — rutas a logos y template
- `Limits` — límites de contenido
- `SlideLayoutIndex` — índice de layout por tipo de slide

---

### `src/theme/` — Carga del tema

#### `src/theme/theme_loader.py`
`ThemeLoader` carga y valida `config/theme.yaml`:
1. Verifica que el archivo exista
2. Parsea el YAML con `yaml.safe_load()`
3. Valida con `ThemeModel.model_validate()` para que todos los campos sean del tipo correcto
4. Devuelve el `ThemeModel` inmutable

Si el YAML tiene un campo faltante o un tipo incorrecto (ej: `size_heading: "grande"` en vez de un número), lanza `ThemeLoadError` con el mensaje de Pydantic.

---

### `src/layout/` — Los renderizadores

Esta es la capa que convierte los modelos de datos en diapositivas de PowerPoint.

#### `src/layout/base_layout.py`
Define la clase abstracta `BaseLayoutRenderer` con un único método abstracto:

```python
def render(self, slide: Slide, model: SlideModel, theme: ThemeModel) -> None: ...
```

Todos los renderizadores deben implementar este método. Esto garantiza que el sistema puede llamar `.render()` en cualquier renderizador sin saber de qué tipo es — el principio de sustitución de Liskov.

#### `src/layout/registry.py`
`LayoutRegistry` es un diccionario que mapea `str → clase de renderizador`:

```python
registry.register("cover", CoverRenderer)
# Internamente: {"cover": CoverRenderer, ...}

renderer_class = registry.get("cover")  # → CoverRenderer
renderer_class().render(slide, model, theme)
```

Si se pide un tipo que no está registrado, lanza `UnregisteredLayoutError` con la lista de tipos disponibles. Esto hace posible agregar nuevos tipos de slide **sin modificar ningún código existente** — solo registras el nuevo renderizador.

#### `src/layout/helpers.py`
Utilidades de bajo nivel que todos los renderizadores comparten. Abstrae la API de python-pptx en funciones simples y legibles:

| Función | Qué hace |
|---|---|
| `hex_to_rgb("#00AEEF")` | Convierte hex a `RGBColor` de python-pptx |
| `set_slide_background(slide, color_hex)` | Pinta el fondo del slide con un color sólido |
| `add_colored_box(slide, left, top, width, height, color_hex)` | Agrega un rectángulo de color sin borde |
| `add_text_box(slide, left, top, width, height, text, ...)` | Agrega una caja de texto con fuente, tamaño, color y alineación |
| `add_bullet_list(slide, left, top, width, height, items, ...)` | Agrega una lista de viñetas con `•` como prefijo |
| `add_image(slide, path, left, top, width, height, fallback)` | Inserta una imagen; si no existe, usa el fallback |
| `truncate_text(text, max_chars, field_name)` | Recorta el texto si supera el límite y registra WARNING |

Todas las medidas van en **pulgadas** porque python-pptx trabaja internamente en EMU (English Metric Units) y la función `Inches()` hace la conversión.

#### `src/layout/renderers/` — Los 7 renderizadores

Cada archivo renderiza un tipo de slide específico. Reciben `(slide, model, theme)` y usan las funciones de `helpers.py` y los valores de `theme` para construir el slide. **Ningún valor visual está hardcodeado en estos archivos** — todo viene de `theme.colors`, `theme.fonts`, `theme.limits`.

**`cover.py` — Portada**
- Fondo: `colors.primary` (azul marino)
- Barra superior: rectángulo delgado `colors.accent`
- Logo: esquina superior izquierda, `assets.logo_white`
- Título: texto grande blanco, tamaño `fonts.size_heading + 4`
- Subtítulo: color `colors.accent`, tamaño reducido
- Línea divisora: rectángulo `colors.accent` de 0.04" de alto
- Cliente: texto muted en la parte media
- Autor + fecha: esquina inferior derecha, separados por `|`

**`section_divider.py` — Divisor de sección**
- Fondo: `colors.secondary` (azul medio)
- Barra vertical derecha: `colors.accent`
- Línea superior: `colors.accent`
- Número de sección: tamaño `fonts.size_stat` (72pt), color `colors.accent`
- Título: tamaño `fonts.size_heading`, color blanco
- Línea divisora horizontal
- Tagline: opcional, cursiva, color muted

**`content_two_col.py` — Contenido en dos columnas**
- Fondo: `colors.background` (blanco)
- Barra de encabezado: `colors.primary`
- Título del slide en la barra
- Columna izquierda: título opcional + lista de bullets
- Columna derecha: igual
- Línea divisora central en `colors.accent`

**`pricing_table.py` — Tabla de precios**
- Fondo: `colors.background`
- Encabezado: `colors.primary`
- Usa `slide.shapes.add_table()` de python-pptx para crear la tabla real
- Fila de encabezado: fondo `colors.primary`, texto blanco, negrita
- Filas de datos: fondo alternado `colors.background` / `colors.divider`
- Fila de totales: fondo `colors.accent`, texto blanco, negrita
- Nota al pie: opcional, cursiva, color muted

**`profile_card.py` — Tarjeta de perfil**
- Panel izquierdo (≈4"): fondo `colors.secondary`, foto o avatar, nombre/rol/seniority/experiencia
- Panel derecho (≈8"): etiquetas de skills como cajas `colors.accent`, lista de highlights

**`stat_callout.py` — Número destacado**
- Fondo: `colors.primary`
- Número gigante centrado: `fonts.size_stat` (72pt), color `colors.accent`
- Label debajo: `fonts.size_heading`, blanco
- Contexto y fuente: texto muted pequeño

**`closing.py` — Cierre**
- Fondo: `colors.primary`
- Headline grande centrado: blanco
- Bloque de contacto: nombre, email, teléfono (opcional)
- CTA (call to action): fondo `colors.accent`, texto blanco
- Logo: esquina inferior

---

### `src/renderer/` — Capa 6: Orquestador

#### `src/renderer/deck_renderer.py`
`DeckRenderer` itera sobre todos los slides del `DeckModel` y los renderiza uno a uno:

```
Para cada slide en deck.slides:
  1. Obtener el índice de layout del theme (layout_indices.cover = 6)
  2. Si el índice es inválido, usar el último disponible y registrar WARNING
  3. Agregar el slide a la presentación: presentation.slides.add_slide(layout)
  4. Buscar el renderizador: registry.get(slide_type) → CoverRenderer
  5. Instanciar y llamar: CoverRenderer().render(slide, model, theme)
  6. Si falla → RenderError con el número de slide y el error original
  7. Registrar: "Slide 1/8 rendered: cover"
```

Al terminar, devuelve el objeto `Presentation` de python-pptx con todos los slides ya agregados.

---

### `src/exporter/` — Capa 7: Escritura al disco

#### `src/exporter/pptx_exporter.py`
`PptxExporter` escribe el archivo final de forma **atómica**. Esto significa que si algo falla durante la escritura (disco lleno, error de permisos), el archivo de destino anterior no queda corrupto.

El mecanismo:
1. Crea la carpeta de destino si no existe (`mkdir -p`)
2. Crea un archivo temporal en la misma carpeta (`tempfile.mkstemp`)
3. Guarda el `.pptx` en el temporal con `presentation.save(tmp_path)`
4. Si el guardado fue exitoso: `os.replace(tmp_path, output_path)` — operación atómica en el sistema de archivos
5. Si algo falla: elimina el temporal y lanza `ExportError`
6. Registra el tamaño final en KB

---

### `src/logger/` — Logging estructurado

#### `src/logger/deck_logger.py`
`DeckLogger` configura Python's `logging` estándar para el proyecto entero:

- **Un solo logger raíz** llamado `"softgic"`. Todos los submódulos usan `logging.getLogger("softgic.nombre_modulo")`, lo que los convierte en hijos automáticos.
- **Consola**: nivel INFO — muestra el progreso de cada paso al usuario.
- **Archivo**: nivel DEBUG — guarda todo incluyendo detalles internos para diagnóstico.
- **Nombre del archivo**: `logs/YYYYMMDD_HHMMSS_deck.log` — uno por ejecución, nunca se sobreescriben.
- **Formato**: `[2026-06-12 10:33:59] [INFO   ] [softgic.renderer] Rendering 8 slide(s)...`

El logger se inicializa una vez en `generate.py` y desde ese momento todos los módulos pueden escribir solo llamando `logging.getLogger("softgic.nombre")`.

---

## `tests/` — Tests automatizados

### `tests/conftest.py` — Fixtures compartidos

`conftest.py` es un archivo especial de pytest. Todo lo que se define aquí está disponible automáticamente en todos los archivos de test sin necesidad de importarlo.

Define:
- **`minimal_theme`**: un `ThemeModel` completo con valores válidos para tests. Evita que cada test tenga que construir el theme desde cero.
- **`blank_presentation`** y **`blank_slide`**: objetos de python-pptx para tests de renderizadores.
- **7 fixtures de slide**: `cover_model`, `section_divider_model`, etc. — uno por cada tipo, ya construidos con datos válidos.
- **`full_deck`**: un `DeckModel` con los 7 slides, para tests de integración.
- **`minimal_raw_deck`** y **`full_raw_deck`**: diccionarios Python (no modelos), para tests del validador y parser.

### `tests/unit/` — Tests unitarios

Cada módulo tiene su propio archivo de test. Un test unitario prueba **una sola unidad** de código de forma aislada.

| Archivo | Qué prueba |
|---|---|
| `test_file_reader.py` | Lectura de YAML/JSON, extensiones inválidas, archivo vacío |
| `test_theme_loader.py` | Carga de theme.yaml, YAML malformado, campos faltantes |
| `test_validator.py` | Las 3 pasadas de validación: casos válidos e inválidos |
| `test_schema_branches.py` | Ramas específicas del validador de esquema (tipos erróneos, nulls) |
| `test_parser.py` | Conversión de dict a DeckModel, error cuando el dict es inválido |
| `test_registry.py` | Registro y obtención de renderizadores, error en tipo desconocido |
| `test_renderer.py` | DeckRenderer: presentación con 7 slides, error en renderer, índice fuera de rango |
| `test_exporter.py` | Escritura del .pptx, creación de carpetas, fallo atómico |
| `test_logger.py` | Inicialización del logger, handlers de consola y archivo |
| `test_asset_validator.py` | Foto no encontrada, imagen inválida, resolución baja |
| `test_renderers/test_cover.py` | Renderizador de portada: happy path, campos opcionales, truncado |
| `test_renderers/test_all_renderers.py` | Los otros 6 renderizadores con los mismos patrones |

### `tests/integration/` — Tests de integración

Los tests de integración ejecutan el **pipeline completo** de principio a fin con archivos reales.

**`test_full_pipeline.py`**: ejecuta `FileReader → Validator → Parser → DeckRenderer → PptxExporter` y verifica que el archivo `.pptx` resultante sea un archivo válido con el número correcto de slides.

**`test_output_consistency.py`**: ejecuta el mismo input dos veces y verifica que los archivos resultantes sean idénticos byte a byte. Esto garantiza que el sistema sea **determinista**: mismo input = mismo output, siempre.

---

## `tools/`

Scripts de utilidad para la configuración inicial. No forman parte del pipeline de generación.

### `tools/create_template.py`
El script original de bootstrap. Crea `assets/template.pptx` con las dimensiones correctas y las imágenes placeholder básicas usando Pillow.

### `tools/create_test_assets.py`
Versión mejorada. Genera:
- `template.pptx`: presentación base 13.33" × 7.5" con 11 layouts
- `logo_white.png`: cuadrado `#00AEEF` + texto "SOFTGIC" en blanco, fondo transparente
- `logo_dark.png`: mismo pero texto en `#0A1628`
- `placeholder_profile.png`: silueta de persona sobre fondo `#1E3A5F` con borde circular azul

Se ejecuta **una sola vez** antes de generar decks por primera vez, o cuando se quieren regenerar los assets.

---

## Reglas de diseño que guían todo el sistema

Estas reglas fueron establecidas al inicio del proyecto y están enforced por la arquitectura:

**N01 — El template manda**: ningún renderizador inventa layouts. Todo viene de `template.pptx` y los índices en `theme.yaml`.

**N02 — Cero hardcoding visual**: si buscas un color hexadecimal o un tamaño de fuente en cualquier archivo `.py`, no lo encontrarás. Todo va a través de `ThemeModel`.

**N03 — Full validation before render**: el sistema valida completamente antes de tocar python-pptx. Si hay un error, falla rápido con un mensaje claro.

**N04 — Same input = same output**: el sistema es determinista. La misma prueba pasa dos veces en `test_output_consistency.py`.

**N05 — Every execution must produce logs**: `DeckLogger` se inicializa en el primer paso de `generate.py` y escribe un archivo en `logs/` en cada ejecución, sin excepción.

---

## Cómo agregar un nuevo tipo de slide

Solo se tocan **4 archivos**, sin modificar ningún código existente:

1. **`src/models/slides.py`** — Agregar la clase `class MiSlide(BaseModel)` y añadirla al `Union` de `SlideModel`
2. **`src/layout/renderers/mi_slide.py`** — Crear el renderizador que extiende `BaseLayoutRenderer`
3. **`generate.py`** — Agregar `registry.register("mi_slide", MiSlideRenderer)`
4. **`config/theme.yaml`** — Agregar `mi_slide: 6` bajo `layout_indices`

El validador de esquema en `schema_rules.py` necesita que también agregues la entrada en `REGISTERED_TYPES`, `REQUIRED_FIELDS` y `FIELD_TYPES`.
