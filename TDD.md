# Technical Design Document
## Softgic Deck Generator — v1.0
**Status:** Pending Architectural Review  
**Date:** 2026-06-12  
**Classification:** Internal — Softgic Engineering  

---

## Table of Contents

1. [Executive Summary](#1-executive-summary)
2. [Architecture Overview](#2-architecture-overview)
3. [Folder Structure](#3-folder-structure)
4. [Layer & Module Design](#4-layer--module-design)
5. [Class Interfaces & Contracts](#5-class-interfaces--contracts)
6. [Pydantic Domain Models](#6-pydantic-domain-models)
7. [theme.yaml Specification](#7-themeyaml-specification)
8. [Input YAML Contract](#8-input-yaml-contract)
9. [Flow Diagrams](#9-flow-diagrams)
10. [Testing Strategy](#10-testing-strategy)
11. [Technical Roadmap](#11-technical-roadmap)
12. [Technical Risks](#12-technical-risks)
13. [Versioning Strategy](#13-versioning-strategy)
14. [Development Conventions](#14-development-conventions)

---

## 1. Executive Summary

The Softgic Deck Generator is a CLI Python application that transforms structured YAML/JSON input files into branded PowerPoint presentations (`.pptx`). It removes the manual burden of constructing decks by encoding Softgic's visual identity into a reusable, configuration-driven pipeline.

**Scope boundary (V1):**  
Input parsing → Validation → Layout rendering → PPTX export.  
No GUI, no AI, no external integrations, no multi-brand support.

**Guiding constraints:**  
- The corporate template (`.pptx`) and `theme.yaml` are the single source of truth for all visual decisions.  
- Identical input always produces identical output (deterministic).  
- Any rendering error must be caught before a partial file is written.  

---

## 2. Architecture Overview

### 2.1 Architectural Style

The system follows a **Layered Pipeline Architecture** with clear unidirectional data flow, enforced through domain models (Pydantic). Each layer communicates only with its adjacent layer via defined interfaces. No layer reaches across.

```
┌─────────────────────────────────────────────────────────────┐
│                        CLI ENTRY POINT                      │
│                    generate.py  (main)                      │
└────────────────────────┬────────────────────────────────────┘
                         │ raw file path
                         ▼
┌─────────────────────────────────────────────────────────────┐
│                      INPUT LAYER                            │
│              FileReader  (YAML / JSON → dict)               │
└────────────────────────┬────────────────────────────────────┘
                         │ raw dict
                         ▼
┌─────────────────────────────────────────────────────────────┐
│                      VALIDATOR                              │
│         Schema rules + business rules + asset checks        │
└────────────────────────┬────────────────────────────────────┘
                         │ validated dict (or raises)
                         ▼
┌─────────────────────────────────────────────────────────────┐
│                        PARSER                               │
│          dict → List[SlideModel]  (Pydantic models)         │
└────────────────────────┬────────────────────────────────────┘
                         │ List[SlideModel]
                         ▼
┌──────────────────────────────────────────────────────────────┐
│                      RENDERER                                │
│   ThemeLoader ──► LayoutRegistry ──► SlideRenderer per type  │
└────────────────────────┬─────────────────────────────────────┘
                         │ python-pptx Presentation object
                         ▼
┌─────────────────────────────────────────────────────────────┐
│                       EXPORTER                              │
│                  Writes output .pptx file                   │
└─────────────────────────────────────────────────────────────┘
                         │
                  [LOGGER cross-cuts all layers]
```

### 2.2 Design Principles Applied

| Principle | Application |
|-----------|------------|
| Single Responsibility | Each class owns exactly one concern |
| Open/Closed | New layouts are added as new classes, never by modifying existing ones |
| Liskov Substitution | All layout renderers implement `BaseLayoutRenderer` identically |
| Interface Segregation | Validator, Parser, Renderer each expose one public method |
| Dependency Inversion | Renderer depends on `LayoutRegistry` abstraction, not concrete renderers |
| Separation of Concerns | Visual config lives only in theme.yaml; business structure lives only in input YAML |
| Determinism (N04) | No random seeds, timestamps, or non-deterministic calls within rendering |

---

## 3. Folder Structure

```
softgic-deck-generator/
│
├── generate.py                    # CLI entry point
│
├── config/
│   └── theme.yaml                 # Visual identity — single source of truth
│
├── assets/
│   ├── template.pptx              # Corporate PowerPoint template
│   ├── logo_white.png
│   ├── logo_dark.png
│   └── placeholder_profile.png
│
├── examples/
│   ├── propuesta_comercial.yaml   # Sample input — commercial proposal
│   ├── propuesta_comercial.json   # Same example in JSON
│   └── talent_profile.yaml       # Sample input — talent acquisition
│
├── src/
│   ├── __init__.py
│   │
│   ├── input/
│   │   ├── __init__.py
│   │   ├── file_reader.py         # Reads YAML / JSON → raw dict
│   │   └── formats.py             # Enum: InputFormat (YAML, JSON)
│   │
│   ├── validator/
│   │   ├── __init__.py
│   │   ├── deck_validator.py      # Orchestrates all validation passes
│   │   ├── schema_rules.py        # Required fields, types per slide type
│   │   ├── business_rules.py      # Cross-field rules, length limits
│   │   └── asset_validator.py     # Verifies referenced images exist & are valid
│   │
│   ├── parser/
│   │   ├── __init__.py
│   │   └── deck_parser.py         # dict → List[SlideModel]
│   │
│   ├── models/
│   │   ├── __init__.py
│   │   ├── deck.py                # DeckModel (top-level)
│   │   ├── slides.py              # One Pydantic model per slide type
│   │   └── theme.py               # ThemeModel
│   │
│   ├── theme/
│   │   ├── __init__.py
│   │   └── theme_loader.py        # Loads & validates theme.yaml → ThemeModel
│   │
│   ├── layout/
│   │   ├── __init__.py
│   │   ├── base_layout.py         # Abstract BaseLayoutRenderer
│   │   ├── registry.py            # LayoutRegistry — Open/Closed extension point
│   │   └── renderers/
│   │       ├── __init__.py
│   │       ├── cover.py
│   │       ├── section_divider.py
│   │       ├── content_two_col.py
│   │       ├── pricing_table.py
│   │       ├── profile_card.py
│   │       ├── stat_callout.py
│   │       └── closing.py
│   │
│   ├── renderer/
│   │   ├── __init__.py
│   │   └── deck_renderer.py       # Orchestrates layout rendering per slide
│   │
│   ├── exporter/
│   │   ├── __init__.py
│   │   └── pptx_exporter.py       # Writes Presentation object → .pptx
│   │
│   └── logger/
│       ├── __init__.py
│       └── deck_logger.py         # Structured logging across all layers
│
├── tests/
│   ├── __init__.py
│   ├── conftest.py                # Shared fixtures
│   ├── unit/
│   │   ├── test_file_reader.py
│   │   ├── test_validator.py
│   │   ├── test_parser.py
│   │   ├── test_theme_loader.py
│   │   ├── test_registry.py
│   │   └── test_renderers/
│   │       ├── test_cover.py
│   │       ├── test_section_divider.py
│   │       ├── test_content_two_col.py
│   │       ├── test_pricing_table.py
│   │       ├── test_profile_card.py
│   │       ├── test_stat_callout.py
│   │       └── test_closing.py
│   └── integration/
│       ├── test_full_pipeline.py
│       └── test_output_consistency.py
│
├── pyproject.toml
├── requirements.txt
└── README.md
```

---

## 4. Layer & Module Design

### 4.1 CLI Entry Point — `generate.py`

**Responsibility:** Accept CLI arguments, wire all components, execute pipeline, handle top-level errors.

**Arguments:**
```
python generate.py --input deck.yaml --output output.pptx [--theme config/theme.yaml]
```

**Wiring sequence:**
1. Parse CLI args
2. Initialize DeckLogger
3. Invoke FileReader
4. Invoke DeckValidator
5. Invoke DeckParser
6. Invoke ThemeLoader
7. Invoke DeckRenderer (injecting LayoutRegistry + ThemeModel)
8. Invoke PptxExporter
9. Log success + exit 0 (or log error + exit 1)

---

### 4.2 Input Layer — `src/input/`

#### `FileReader`
- Reads a file path, detects format (by extension: `.yaml`/`.yml` → YAML, `.json` → JSON).
- Returns a plain `dict`.
- Raises `InputReadError` if file does not exist, is empty, or cannot be parsed.
- Zero business logic; zero model knowledge.

---

### 4.3 Validator — `src/validator/`

#### `DeckValidator` (orchestrator)
Runs three sequential passes. Raises `ValidationError` (with structured message) on first failure per pass — never silently continues.

**Pass 1 — SchemaRules:**
- Verifies top-level `slides` key exists and is a non-empty list.
- For each slide: verifies `type` field exists and is a registered type.
- Verifies all required fields per slide type are present.
- Verifies field value types match expected types.

**Pass 2 — BusinessRules:**
- Title length ≤ configured max (from theme.yaml limits section).
- Bullet list length ≤ configured max items per column.
- Pricing table rows ≤ configured max rows.
- Profile card skills ≤ configured max skills.
- Numeric fields (big_number, years_experience) are positive.
- Email field in Closing slide is a valid email format.

**Pass 3 — AssetValidator:**
- For each `photo` field in profile_card: verifies file path exists on disk.
- Verifies file is a valid image (Pillow can open it).
- Verifies image resolution meets minimum (configurable in theme.yaml).

---

### 4.4 Parser — `src/parser/`

#### `DeckParser`
- Receives a validated `dict`.
- Maps each slide dict to the corresponding Pydantic model using a type-dispatch map.
- Returns `DeckModel` containing `List[SlideModel]`.
- Never receives raw unvalidated dicts from outside.
- Raises `ParseError` if a model instantiation fails (defensive — should not happen after validation).

---

### 4.5 Models — `src/models/`

See Section 6 for full Pydantic model definitions.

---

### 4.6 Theme Loader — `src/theme/`

#### `ThemeLoader`
- Reads `theme.yaml` from disk.
- Validates it against `ThemeModel` (Pydantic).
- Returns an immutable `ThemeModel` instance.
- Raises `ThemeLoadError` if file is missing, malformed, or fails model validation.
- Called once per pipeline execution; result is passed by injection to Renderer.

---

### 4.7 Layout Registry — `src/layout/registry.py`

#### `LayoutRegistry`

The single Open/Closed extension point of the system. New layouts are registered here; no existing code changes.

```
LayoutRegistry
  ├── _registry: dict[str, Type[BaseLayoutRenderer]]
  ├── register(slide_type: str, renderer_class: Type[BaseLayoutRenderer]) → None
  ├── get(slide_type: str) → Type[BaseLayoutRenderer]      # raises if not found
  └── available_types() → List[str]
```

All seven V1 renderers are registered at application startup in `generate.py`.

#### `BaseLayoutRenderer` (Abstract)

```
BaseLayoutRenderer  (ABC)
  └── render(slide: pptx.Slide, model: SlideModel, theme: ThemeModel) → None
```

Every layout renderer must implement exactly this interface. The renderer receives a pre-added slide object (from the corporate template), the typed slide model, and the theme. It never creates slides — only populates them.

---

### 4.8 Layout Renderers — `src/layout/renderers/`

One class per slide type. Each:
- Extends `BaseLayoutRenderer`.
- Implements `render()`.
- Resolves placeholder indices or named shapes from the template.
- Applies text, fonts, colors, images from `ThemeModel`.
- Handles text overflow via truncation with ellipsis + warning log (never crashes).
- Has zero hardcoded visual values.

**Text overflow strategy:**  
If text exceeds the configured character limit for a shape, the renderer:
1. Truncates to limit and appends `"…"`.
2. Emits a `WARNING` log entry with field name and original length.
3. Continues rendering.

This is the only case where a warning is acceptable instead of an error.

---

### 4.9 Renderer — `src/renderer/`

#### `DeckRenderer`
- Receives: `DeckModel`, `ThemeModel`, `LayoutRegistry`, `Presentation` (loaded from template).
- Iterates over each `SlideModel` in order.
- Looks up the renderer class via `LayoutRegistry.get(slide.type)`.
- Adds a new slide to the Presentation from the correct template layout index.
- Calls `renderer_instance.render(slide_obj, model, theme)`.
- Returns the populated `Presentation` object.
- Never writes to disk.

---

### 4.10 Exporter — `src/exporter/`

#### `PptxExporter`
- Receives: `Presentation` object, output file path.
- Validates output directory exists (creates it if needed).
- Writes the file atomically: writes to a temp path first, then renames (prevents partial writes on failure).
- Logs output path and file size.
- Raises `ExportError` on failure.

---

### 4.11 Logger — `src/logger/`

#### `DeckLogger`
- Wraps Python's standard `logging` module.
- Structured log format: `[TIMESTAMP] [LEVEL] [MODULE] message`.
- Log levels: `DEBUG`, `INFO`, `WARNING`, `ERROR`.
- Outputs to both console (INFO+) and file (`logs/YYYYMMDD_HHMMSS_deck.log`).
- Mandatory log events:
  - Pipeline start (input file, output file, theme file, timestamp).
  - Validation pass/fail per pass with details.
  - Each slide rendered (type, position).
  - Text truncation warnings (field, original length, truncated length).
  - Export success (path, size in KB).
  - Any exception with full traceback.

---

## 5. Class Interfaces & Contracts

```
FileReader
  + read(file_path: Path) → dict
  ~ raises: InputReadError

DeckValidator
  + validate(raw: dict) → None
  ~ raises: ValidationError(message: str, field: str, slide_index: int)

DeckParser
  + parse(validated: dict) → DeckModel
  ~ raises: ParseError

ThemeLoader
  + load(theme_path: Path) → ThemeModel
  ~ raises: ThemeLoadError

LayoutRegistry
  + register(slide_type: str, cls: Type[BaseLayoutRenderer]) → None
  + get(slide_type: str) → Type[BaseLayoutRenderer]
  + available_types() → List[str]
  ~ raises: UnregisteredLayoutError

BaseLayoutRenderer  [ABC]
  + render(slide: pptx.slide.Slide, model: SlideModel, theme: ThemeModel) → None

DeckRenderer
  + render(deck: DeckModel, theme: ThemeModel, registry: LayoutRegistry,
           presentation: Presentation) → Presentation
  ~ raises: RenderError

PptxExporter
  + export(presentation: Presentation, output_path: Path) → None
  ~ raises: ExportError

DeckLogger
  + info(module: str, message: str) → None
  + warning(module: str, message: str) → None
  + error(module: str, message: str, exc: Exception | None) → None
  + debug(module: str, message: str) → None
```

**Custom Exception Hierarchy:**
```
SoftgicDeckError  (base)
  ├── InputReadError
  ├── ValidationError
  │     ├── SchemaValidationError
  │     ├── BusinessRuleError
  │     └── AssetValidationError
  ├── ParseError
  ├── ThemeLoadError
  ├── UnregisteredLayoutError
  ├── RenderError
  └── ExportError
```

All exceptions carry a human-readable `message`, an optional `field`, and optional `slide_index`.

---

## 6. Pydantic Domain Models

### 6.1 Top-Level

```python
class DeckModel(BaseModel):
    slides: List[SlideModel]          # minimum 1
    metadata: DeckMetadata | None = None

class DeckMetadata(BaseModel):
    generated_at: str | None = None   # ISO timestamp — injected by pipeline, not from input
    input_file: str | None = None
    theme_version: str | None = None
```

### 6.2 Slide Type Union

```python
SlideModel = Annotated[
    Union[
        CoverSlide,
        SectionDividerSlide,
        ContentTwoColSlide,
        PricingTableSlide,
        ProfileCardSlide,
        StatCalloutSlide,
        ClosingSlide,
    ],
    Field(discriminator="type")
]
```

### 6.3 Per-Slide Models

```python
class CoverSlide(BaseModel):
    type: Literal["cover"]
    title: str
    subtitle: str | None = None
    client: str | None = None
    date: str | None = None          # display string — not parsed as date
    author: str | None = None

class SectionDividerSlide(BaseModel):
    type: Literal["section_divider"]
    section_number: str              # e.g. "01"
    section_title: str
    tagline: str | None = None

class BulletColumn(BaseModel):
    title: str | None = None
    bullets: List[str]               # min 1, max driven by theme limits

class ContentTwoColSlide(BaseModel):
    type: Literal["content_two_col"]
    title: str
    left: BulletColumn
    right: BulletColumn

class PricingRow(BaseModel):
    description: str
    quantity: int
    unit_price: float
    total: float                     # renderer validates quantity * unit_price ≈ total

class PricingTableSlide(BaseModel):
    type: Literal["pricing_table"]
    title: str
    currency: str                    # e.g. "USD", "COP"
    rows: List[PricingRow]           # min 1
    totals: float
    notes: str | None = None

class ProfileCardSlide(BaseModel):
    type: Literal["profile_card"]
    name: str
    role: str
    years_experience: int            # >= 0
    seniority: str                   # e.g. "Senior", "Mid", "Junior"
    skills: List[str]                # max driven by theme limits
    highlights: List[str] | None = None
    photo: str | None = None         # file path — validated by AssetValidator

class StatCalloutSlide(BaseModel):
    type: Literal["stat_callout"]
    big_number: str                  # display string e.g. "98%", "+300"
    label: str
    context: str | None = None
    source: str | None = None

class ClosingSlide(BaseModel):
    type: Literal["closing"]
    headline: str
    contact_name: str
    contact_email: str               # validated as email format
    contact_phone: str | None = None
    cta: str | None = None           # call to action text
```

### 6.4 Theme Model

```python
class ColorPalette(BaseModel):
    primary: str                     # hex e.g. "#0A1628"
    secondary: str
    accent: str
    background: str
    text_dark: str
    text_light: str
    text_muted: str
    divider: str

class FontConfig(BaseModel):
    family: str                      # font name — must be installed on system
    size_heading: int                # pt
    size_subheading: int
    size_body: int
    size_caption: int
    size_stat: int
    bold_headings: bool = True

class AssetPaths(BaseModel):
    template: str                    # path to .pptx template
    logo_white: str
    logo_dark: str
    placeholder_profile: str

class Limits(BaseModel):
    title_max_chars: int = 80
    subtitle_max_chars: int = 120
    bullet_max_chars: int = 100
    bullets_per_column_max: int = 6
    skills_max: int = 8
    pricing_rows_max: int = 10
    profile_image_min_width: int = 200   # px
    profile_image_min_height: int = 200

class SlideLayoutIndex(BaseModel):
    cover: int
    section_divider: int
    content_two_col: int
    pricing_table: int
    profile_card: int
    stat_callout: int
    closing: int

class ThemeModel(BaseModel):
    version: str
    colors: ColorPalette
    fonts: FontConfig
    assets: AssetPaths
    limits: Limits
    layout_indices: SlideLayoutIndex

    model_config = ConfigDict(frozen=True)   # immutable after load
```

---

## 7. theme.yaml Specification

```yaml
# Softgic Deck Generator — Corporate Theme
# Version controlled. All visual decisions live here.
# Changing this file affects every generated deck.

version: "1.0.0"

colors:
  primary: "#0A1628"          # Softgic Navy — backgrounds, headers
  secondary: "#1E3A5F"        # Mid-blue — section backgrounds
  accent: "#00AEEF"           # Softgic Blue — highlights, icons
  background: "#FFFFFF"       # Slide background
  text_dark: "#0A1628"        # Body text on light backgrounds
  text_light: "#FFFFFF"       # Text on dark backgrounds
  text_muted: "#6B7280"       # Captions, secondary info
  divider: "#E5E7EB"          # Horizontal rules, table lines

fonts:
  family: "Montserrat"        # Must be installed; fallback handled in renderer warning
  size_heading: 36            # pt
  size_subheading: 24
  size_body: 14
  size_caption: 10
  size_stat: 72               # stat_callout big number
  bold_headings: true

assets:
  template: "assets/template.pptx"
  logo_white: "assets/logo_white.png"
  logo_dark: "assets/logo_dark.png"
  placeholder_profile: "assets/placeholder_profile.png"

limits:
  title_max_chars: 80
  subtitle_max_chars: 120
  bullet_max_chars: 100
  bullets_per_column_max: 6
  skills_max: 8
  pricing_rows_max: 10
  profile_image_min_width: 200    # px
  profile_image_min_height: 200

layout_indices:
  cover: 0
  section_divider: 1
  content_two_col: 2
  pricing_table: 3
  profile_card: 4
  stat_callout: 5
  closing: 6
```

**Rules:**
- `version` follows SemVer. Changing colors/fonts increments PATCH. Adding new keys increments MINOR. Removing keys increments MAJOR.
- All hex values must include the `#` prefix.
- `layout_indices` must match the actual slide layout order in `template.pptx`.
- `fonts.family` must name a font that is installed on the execution machine. Renderer warns if it falls back to a system default — it does not fail.

---

## 8. Input YAML Contract

### 8.1 Full Example — Commercial Proposal

```yaml
# propuesta_comercial.yaml
# Softgic Deck Generator — Input Contract v1

slides:

  - type: cover
    title: "Propuesta de Solución Digital"
    subtitle: "Transformación del área comercial"
    client: "Cliente XYZ S.A.S."
    date: "Junio 2026"
    author: "Equipo Comercial Softgic"

  - type: section_divider
    section_number: "01"
    section_title: "Contexto del Reto"
    tagline: "Entendemos el problema antes de proponer la solución."

  - type: content_two_col
    title: "Situación Actual vs. Situación Esperada"
    left:
      title: "Situación Actual"
      bullets:
        - "Proceso manual de cotización"
        - "Reprocesos por errores de datos"
        - "Tiempo promedio: 3 días por propuesta"
        - "Sin trazabilidad de versiones"
    right:
      title: "Situación Esperada"
      bullets:
        - "Generación automática desde plantilla"
        - "Reducción de errores a < 1%"
        - "Tiempo objetivo: 30 minutos"
        - "Historial completo de versiones"

  - type: pricing_table
    title: "Inversión del Proyecto"
    currency: "USD"
    rows:
      - description: "Análisis y Arquitectura"
        quantity: 1
        unit_price: 5000.00
        total: 5000.00
      - description: "Desarrollo Backend"
        quantity: 3
        unit_price: 8000.00
        total: 24000.00
      - description: "QA y Pruebas"
        quantity: 1
        unit_price: 3000.00
        total: 3000.00
    totals: 32000.00
    notes: "Precios en dólares. No incluye IVA."

  - type: stat_callout
    big_number: "98%"
    label: "Satisfacción de clientes Softgic"
    context: "Basado en encuestas post-entrega 2025"
    source: "Área de Customer Success"

  - type: profile_card
    name: "Laura Martínez"
    role: "Tech Lead — Backend"
    years_experience: 8
    seniority: "Senior"
    skills:
      - "Python"
      - "FastAPI"
      - "AWS"
      - "PostgreSQL"
      - "Docker"
    highlights:
      - "Lideró migración a microservicios en FinTech (2024)"
      - "Certificada AWS Solutions Architect"
    photo: "assets/profiles/laura_martinez.jpg"

  - type: closing
    headline: "¿Listo para transformar tu operación?"
    contact_name: "Carlos Gómez"
    contact_email: "carlos.gomez@softgic.com"
    contact_phone: "+57 300 123 4567"
    cta: "Agenda tu demo en los próximos 5 días"
```

### 8.2 Validation Rules Summary

| Field | Rule | Error Level |
|-------|------|-------------|
| `slides` | Required, non-empty list | ERROR |
| `type` | Required, must be registered | ERROR |
| Required fields per type | All present | ERROR |
| `title` length | ≤ `limits.title_max_chars` | WARNING (truncate) |
| `bullets` per column | ≤ `limits.bullets_per_column_max` | ERROR |
| `skills` count | ≤ `limits.skills_max` | ERROR |
| `pricing_table.rows` count | ≤ `limits.pricing_rows_max` | ERROR |
| `profile_card.photo` | File exists, valid image | ERROR |
| `profile_card.photo` resolution | ≥ min dimensions | WARNING |
| `closing.contact_email` | Valid email format | ERROR |
| `profile_card.years_experience` | Integer ≥ 0 | ERROR |

### 8.3 JSON Equivalent

The same contract applies to JSON. The `FileReader` normalizes both formats into the same `dict` before validation. No JSON-specific rules exist.

```json
{
  "slides": [
    {
      "type": "cover",
      "title": "Propuesta de Solución Digital",
      "subtitle": "Transformación del área comercial",
      "client": "Cliente XYZ S.A.S.",
      "date": "Junio 2026",
      "author": "Equipo Comercial Softgic"
    }
  ]
}
```

---

## 9. Flow Diagrams

### 9.1 Happy Path — Full Pipeline

```
User
  │
  ├─▶ python generate.py --input deck.yaml --output output.pptx
  │
  ▼
generate.py
  │
  ├─▶ DeckLogger.init()
  ├─▶ FileReader.read(deck.yaml) ──────────────────────────▶ raw_dict
  ├─▶ DeckValidator.validate(raw_dict) ───────────────────▶ (passes or raises)
  ├─▶ DeckParser.parse(raw_dict) ─────────────────────────▶ DeckModel
  ├─▶ ThemeLoader.load(theme.yaml) ───────────────────────▶ ThemeModel
  ├─▶ open template.pptx → Presentation
  ├─▶ DeckRenderer.render(DeckModel, ThemeModel, Registry, Pres) ▶ Presentation
  │     │
  │     ├─ for each SlideModel:
  │     │    ├─ LayoutRegistry.get(slide.type) → RendererClass
  │     │    ├─ prs.slides.add_slide(layout_index)
  │     │    └─ RendererClass.render(slide_obj, model, theme)
  │     │
  │     └─▶ populated Presentation
  │
  ├─▶ PptxExporter.export(Presentation, output.pptx)
  ├─▶ DeckLogger.info("Export complete: output.pptx — 842 KB")
  └─▶ exit 0
```

### 9.2 Validation Failure Path

```
FileReader.read() ──▶ raw_dict
  │
  ▼
DeckValidator.validate(raw_dict)
  │
  ├─ Pass 1 (Schema): FAIL — missing required field "title" in slide[2]
  │
  ▼
raise SchemaValidationError(
    message="Required field 'title' missing",
    field="title",
    slide_index=2
)
  │
  ▼
generate.py catches SoftgicDeckError
  │
  ├─▶ DeckLogger.error(message, exc)
  ├─▶ print("ERROR [slide 2]: Required field 'title' missing in slide type 'cover'")
  └─▶ exit 1
```

### 9.3 Layout Extension — Adding a New Layout

```
Developer adds new layout (no existing files touched):

1. Create src/layout/renderers/team_overview.py
   └── class TeamOverviewRenderer(BaseLayoutRenderer):
           def render(self, slide, model, theme): ...

2. Create src/models/slides.py addition:
   └── class TeamOverviewSlide(BaseModel):
           type: Literal["team_overview"]
           ...

3. Register in generate.py:
   └── registry.register("team_overview", TeamOverviewRenderer)

4. Add layout index to theme.yaml:
   └── layout_indices.team_overview: 7

5. Add tests in tests/unit/test_renderers/test_team_overview.py

No other files modified. Open/Closed satisfied.
```

---

## 10. Testing Strategy

### 10.1 Philosophy

- Tests verify behavior, not implementation.
- All tests are deterministic: fixed inputs, no timestamps, no randomness.
- No external services, no network, no real PowerPoint viewer.
- Integration tests produce a real `.pptx` and verify structural integrity, not visual output.

### 10.2 Test Levels

#### Unit Tests — `tests/unit/`

| Module | What's tested |
|--------|--------------|
| `FileReader` | YAML read OK, JSON read OK, missing file → InputReadError, invalid YAML → InputReadError, empty file → InputReadError |
| `DeckValidator` | Each required field missing, wrong type, email format, photo not found, photo below min resolution, excess bullets, excess skills, excess pricing rows |
| `DeckParser` | Each slide type produces correct Pydantic model, unknown type → ParseError, optional fields default correctly |
| `ThemeLoader` | Valid theme.yaml → ThemeModel, missing key → ThemeLoadError, wrong type → ThemeLoadError |
| `LayoutRegistry` | Register + get, get unregistered → UnregisteredLayoutError, available_types returns correct list |
| Each Renderer | Renders without exception, text truncation triggers warning, missing optional fields use defaults, photo placeholder used when photo=None |

**Fixture strategy:**
- `conftest.py` provides: minimal valid `DeckModel`, `ThemeModel`, a mock `pptx.slide.Slide`, and sample YAML strings.
- Renderer tests use a real in-memory `Presentation` object (no disk I/O).

#### Integration Tests — `tests/integration/`

| Test | What's verified |
|------|----------------|
| `test_full_pipeline` | Full pipeline with all 7 slide types produces a `.pptx` file with correct slide count |
| `test_output_consistency` | Same input run twice produces byte-identical output (determinism, N04) |
| `test_validation_blocks_render` | An invalid input never produces a `.pptx` file |
| `test_theme_change_propagates` | Changing a color in theme causes no crash; output is structurally valid |
| `test_text_truncation` | Input with overlong text produces a `.pptx` with truncated text and a WARNING log entry |

### 10.3 Coverage Target

| Layer | Target |
|-------|--------|
| Validator | 100% — all error branches |
| Parser | 100% — all slide types |
| Renderers | 90%+ — render path + truncation path |
| Exporter | 90%+ |
| Integration | All 7 slide types rendered |

### 10.4 Test Execution

```bash
pytest tests/ -v --tb=short
pytest tests/ --cov=src --cov-report=term-missing
```

### 10.5 Fixtures Design

```python
# conftest.py (sketch)

@pytest.fixture
def minimal_theme() -> ThemeModel:
    # Returns a fully populated ThemeModel with test values

@pytest.fixture
def cover_model() -> CoverSlide:
    return CoverSlide(type="cover", title="Test Title")

@pytest.fixture
def mock_slide(minimal_theme):
    # Returns a pptx.slide.Slide from an in-memory Presentation
    # built from a minimal test template

@pytest.fixture
def tmp_output_path(tmp_path) -> Path:
    return tmp_path / "output.pptx"
```

---

## 11. Technical Roadmap

### Phase 0 — Foundation (Week 1)
- [ ] Repository scaffolding (folder structure, `pyproject.toml`, `.gitignore`)
- [ ] Exception hierarchy
- [ ] `DeckLogger` implementation and unit tests
- [ ] `ThemeModel` + `theme.yaml` first draft
- [ ] `ThemeLoader` + unit tests

### Phase 1 — Data Pipeline (Week 2)
- [ ] `FileReader` + unit tests
- [ ] All Pydantic slide models
- [ ] `DeckParser` + unit tests
- [ ] `DeckValidator` (Schema + Business Rules) + unit tests
- [ ] `AssetValidator` + unit tests

### Phase 2 — Rendering Core (Week 3)
- [ ] `BaseLayoutRenderer` abstract class
- [ ] `LayoutRegistry` + unit tests
- [ ] `CoverRenderer` + unit tests
- [ ] `SectionDividerRenderer` + unit tests
- [ ] `ContentTwoColRenderer` + unit tests
- [ ] `ClosingRenderer` + unit tests

### Phase 3 — Rendering Complex Layouts (Week 4)
- [ ] `PricingTableRenderer` + unit tests
- [ ] `ProfileCardRenderer` (with image handling) + unit tests
- [ ] `StatCalloutRenderer` + unit tests
- [ ] `DeckRenderer` orchestrator + unit tests

### Phase 4 — Exporter & CLI (Week 5)
- [ ] `PptxExporter` + unit tests
- [ ] `generate.py` CLI wiring
- [ ] Integration tests (full pipeline)
- [ ] Determinism test (N04 compliance)
- [ ] End-to-end test with real template + real `theme.yaml`

### Phase 5 — Hardening (Week 6)
- [ ] Coverage audit (close gaps)
- [ ] Performance: benchmark against 10-minute acceptance criterion
- [ ] Documentation: README with usage examples
- [ ] Acceptance walkthrough with business stakeholders

---

## 12. Technical Risks

| # | Risk | Probability | Impact | Mitigation |
|---|------|-------------|--------|------------|
| R01 | Corporate `template.pptx` has unstable layout indices or unnamed placeholders | HIGH | HIGH | Define shape name conventions early; document expected placeholder names per layout; write an inspection utility to audit the template |
| R02 | Montserrat (or other corporate font) not installed on the execution machine | MEDIUM | MEDIUM | Renderer warns (does not fail); README documents font installation requirement; system font fallback is documented |
| R03 | Text overflow in complex layouts (e.g., pricing table with long descriptions) breaks visual output | HIGH | MEDIUM | All text fields have explicit character limits in `theme.yaml`; truncation + warning is the default behavior; limits are tunable without code changes |
| R04 | Profile photo images with unexpected formats or EXIF rotation | MEDIUM | LOW | Pillow normalizes rotation on load; unsupported formats → AssetValidationError before render |
| R05 | python-pptx doesn't fully support all template features (e.g., master slide effects, custom animations) | LOW | LOW | Scope explicitly excludes animations/transitions; any unsupported feature is documented as out of scope |
| R06 | Theme version drift — `theme.yaml` changes break existing input files | MEDIUM | MEDIUM | ThemeModel versioned; MINOR version bump required for additive changes; MAJOR for breaking changes; validated via CI |
| R07 | Output `.pptx` not byte-identical across OS (line endings, zip compression) | LOW | MEDIUM | Integration test for determinism run on CI using fixed OS; python-pptx uses zipfile which is OS-neutral |
| R08 | Corporate template file is replaced without updating `layout_indices` in `theme.yaml` | MEDIUM | HIGH | Template inspection utility validates `layout_indices` against actual template on startup (only in debug mode) |

---

## 13. Versioning Strategy

### 13.1 Application Version

Follows **Semantic Versioning (SemVer)**: `MAJOR.MINOR.PATCH`

| Change | Version bump |
|--------|-------------|
| Breaking input contract change | MAJOR |
| New layout type added | MINOR |
| Bug fix, internal refactor | PATCH |

Version is stored in `pyproject.toml` and exposed via `generate.py --version`.

### 13.2 Theme Version

`theme.yaml` carries its own `version` field (also SemVer). The `ThemeLoader` logs the loaded theme version. If `ThemeModel` introduces breaking schema changes, the version contract is documented in `CHANGELOG.md`.

### 13.3 Input Contract Version

The input YAML contract is considered stable at V1 (this document). Future field additions are additive (backward-compatible). Removing fields or changing types requires a MAJOR version bump in both the app and documentation.

### 13.4 Git Branching

```
main          ← stable, production-ready
  └── develop ← integration branch
        ├── feature/phase-1-data-pipeline
        ├── feature/phase-2-cover-renderer
        └── fix/profile-card-image-rotation
```

Tags on `main` follow the SemVer application version: `v1.0.0`, `v1.1.0`, etc.

---

## 14. Development Conventions

### 14.1 Language & Runtime

- Python 3.11+ (uses `match`/`case` for type dispatch, `tomllib` for config).
- Type hints are mandatory on all public interfaces. `mypy --strict` must pass.

### 14.2 Code Style

- **Formatter:** `black` (line length: 100).
- **Linter:** `ruff` with `E`, `F`, `I` rule sets.
- **Type checker:** `mypy --strict`.
- No `# type: ignore` without a documented justification comment.

### 14.3 Naming Conventions

| Artifact | Convention | Example |
|----------|-----------|---------|
| Modules | `snake_case` | `deck_validator.py` |
| Classes | `PascalCase` | `DeckValidator` |
| Methods | `snake_case` | `validate_schema()` |
| Pydantic models | `PascalCase` with `Slide` suffix | `CoverSlide` |
| Constants | `UPPER_SNAKE` | `DEFAULT_THEME_PATH` |
| Private methods | `_snake_case` | `_check_required_fields()` |
| Test files | `test_<module>.py` | `test_deck_validator.py` |
| Test functions | `test_<behavior>_<condition>` | `test_validate_raises_on_missing_title()` |

### 14.4 Dependency Rules (Clean Architecture Enforcement)

```
generate.py       → may import any layer
src/exporter      → may import: models, logger
src/renderer      → may import: models, layout, theme, logger
src/layout        → may import: models, theme, logger
src/parser        → may import: models, logger
src/validator     → may import: models, logger
src/input         → may import: logger
src/models        → may import nothing from src/
src/theme         → may import: models, logger
src/logger        → may import nothing from src/
```

No layer imports from a "higher" layer. Enforced via `import-linter` rules in `pyproject.toml`.

### 14.5 Configuration

```toml
# pyproject.toml

[tool.black]
line-length = 100

[tool.mypy]
strict = true
python_version = "3.11"

[tool.pytest.ini_options]
testpaths = ["tests"]
addopts = "-v --tb=short"

[importlinter]
root_packages = ["src"]
# contracts defined per layer dependency rules above
```

### 14.6 Prohibited Patterns

- **No hardcoded colors, fonts, sizes, or paths.** All visual values from `ThemeModel`. Enforced via code review.
- **No `dict` access after parsing.** Only typed Pydantic models flow through the system after `DeckParser`.
- **No silent failures.** Every error either raises a typed exception or emits a WARNING log. Never `except: pass`.
- **No global mutable state.** `ThemeModel` is frozen. `LayoutRegistry` is instantiated in `generate.py` and injected.
- **No direct calls to `print()` from within `src/`.** All output goes through `DeckLogger`.

---

## Appendix A — Dependencies

```
# requirements.txt
python-pptx>=0.6.23
Pillow>=10.0.0
pydantic>=2.5.0
pyyaml>=6.0.1
pytest>=7.4.0
pytest-cov>=4.1.0
black>=23.0.0
ruff>=0.1.0
mypy>=1.5.0
import-linter>=2.0
```

---

## Appendix B — Acceptance Criteria Traceability

| Acceptance Criterion | Covered By |
|---------------------|-----------|
| Deck generated in < 10 minutes | Phase 5 benchmark test; pipeline is local/offline |
| 100% Softgic branding compliance | All visual values from `theme.yaml`; N02 enforced |
| Errors are clear and controlled | Typed exception hierarchy + structured logger |
| Long text doesn't break layouts | Truncation + warning in all renderers; limits in theme |
| theme.yaml change impacts whole deck | ThemeModel injected to every renderer; zero hardcoding |
| New layout without touching core | LayoutRegistry + BaseLayoutRenderer; Open/Closed |

---

*Document prepared by: Engineering Team — Softgic Deck Generator*  
*Status: **Pending Architectural Review and Approval***  
*Next step: Upon approval, Phase 0 implementation begins.*
