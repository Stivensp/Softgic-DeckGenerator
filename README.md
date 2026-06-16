# Softgic Deck Generator

Generate branded Softgic PowerPoint decks from a YAML or JSON file.  
No PowerPoint required. No design knowledge required.

---

## Requirements

- Python 3.11+
- Fonts: `Calibri` (built-in on Windows). Replace with `Montserrat` in `config/theme.yaml` once installed.

## Setup

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Create the base template and placeholder assets (run once)
python tools/create_template.py
```

When the official Softgic corporate template is delivered:

1. Drop it into `assets/template.pptx`
2. Update `layout_indices` in `config/theme.yaml` to match its slide layout positions
3. Replace `assets/logo_white.png`, `assets/logo_dark.png`, and `assets/placeholder_profile.png` with the official assets

---

## Usage

```bash
python generate.py --input your_deck.yaml --output output/deck.pptx
python generate.py --input your_deck.json --output output/deck.pptx
python generate.py --version
```

| Flag | Default | Description |
|------|---------|-------------|
| `--input` / `-i` | required | YAML or JSON input file |
| `--output` / `-o` | `output.pptx` | Output `.pptx` path |
| `--theme` / `-t` | `config/theme.yaml` | Theme configuration file |

### Examples

```bash
python generate.py --input examples/propuesta_comercial.yaml --output output/propuesta.pptx
python generate.py --input examples/talent_profile.yaml     --output output/talent.pptx
```

---

## Input Format

The input file is a YAML (or JSON) document with a `slides` list.  
Each slide has a `type` field and type-specific fields.

### Slide Types

#### `cover`
```yaml
- type: cover
  title: "Propuesta Comercial"          # required
  subtitle: "Transformación digital"    # optional
  client: "Cliente XYZ"                 # optional
  date: "Junio 2026"                    # optional
  author: "Equipo Comercial"            # optional
```

#### `section_divider`
```yaml
- type: section_divider
  section_number: "01"                  # required
  section_title: "Contexto del Reto"    # required
  tagline: "Frase de impacto."          # optional
```

#### `content_two_col`
```yaml
- type: content_two_col
  title: "Situación Actual vs. Esperada"  # required
  left:
    title: "Hoy"                          # optional
    bullets:                              # required (1–6 items)
      - "Proceso manual"
      - "Sin trazabilidad"
  right:
    title: "Mañana"
    bullets:
      - "Proceso automatizado"
      - "Historial completo"
```

#### `pricing_table`
```yaml
- type: pricing_table
  title: "Inversión del Proyecto"       # required
  currency: "USD"                       # required
  rows:                                 # required (1–10 rows)
    - description: "Backend Development"
      quantity: 3
      unit_price: 8000.00
      total: 24000.00
  totals: 24000.00                      # required
  notes: "No incluye IVA."              # optional
```

#### `profile_card`
```yaml
- type: profile_card
  name: "Laura Martínez"               # required
  role: "Tech Lead — Backend"          # required
  years_experience: 8                  # required (>= 0)
  seniority: "Senior"                  # required
  skills:                              # required (1–8 items)
    - "Python"
    - "AWS"
  highlights:                          # optional
    - "Lideró migración a microservicios"
  photo: "assets/profiles/laura.jpg"   # optional (file must exist)
```

#### `stat_callout`
```yaml
- type: stat_callout
  big_number: "98%"                    # required
  label: "Satisfacción de clientes"    # required
  context: "Encuestas post-entrega"    # optional
  source: "Customer Success"           # optional
```

#### `closing`
```yaml
- type: closing
  headline: "¿Listo para empezar?"     # required
  contact_name: "Carlos Gómez"         # required
  contact_email: "carlos@softgic.com"  # required (valid email)
  contact_phone: "+57 300 123 4567"    # optional
  cta: "Agenda tu demo hoy"            # optional
```

---

## Branding & Theme

All visual values live in `config/theme.yaml`. **No values are hardcoded in Python.**

```yaml
colors:
  primary: "#0A1628"      # dark backgrounds, headers
  accent:  "#00AEEF"      # highlights, tags, CTAs
  ...

fonts:
  family: "Calibri"       # change to "Montserrat" when installed
  size_heading: 36
  ...

limits:
  title_max_chars: 80     # text exceeding this is truncated + warning logged
  bullets_per_column_max: 6
  ...
```

Changing any value in `theme.yaml` affects every deck generated afterwards.

---

## Validation

Validation runs before rendering in three passes:

1. **Schema** — required fields present, correct types
2. **Business rules** — lengths, counts, email format, positive numbers
3. **Asset checks** — photo files exist and are valid images

On any error the process stops immediately with a clear message and a non-zero exit code. No partial files are written.

---

## Extending with a New Layout

Adding a new slide type requires touching **4 files** — no existing code changes:

1. **Model** — add a new `class MySlide(BaseModel)` to [src/models/slides.py](src/models/slides.py) and add it to the `SlideModel` union
2. **Renderer** — create `src/layout/renderers/my_slide.py` extending `BaseLayoutRenderer`
3. **Registration** — call `registry.register("my_slide", MySlideRenderer)` in [generate.py](generate.py)
4. **Theme** — add `my_slide: <layout_index>` under `layout_indices` in `config/theme.yaml`

No existing files need modification. Open/Closed principle is enforced by design.

---

## Running Tests

```bash
# All tests
pytest tests/ -v

# With coverage
pytest tests/ --cov=src --cov-report=term-missing

# Unit tests only
pytest tests/unit/ -v

# Integration tests only
pytest tests/integration/ -v
```

Current status: **121 tests — 99% coverage**

---

## Project Structure

```
softgic-deck-generator/
├── generate.py               # CLI entry point
├── config/theme.yaml         # All visual values (colors, fonts, limits)
├── assets/                   # template.pptx + logo + placeholder images
├── examples/                 # Sample YAML input files
├── src/
│   ├── exceptions.py         # Typed exception hierarchy
│   ├── input/                # FileReader (YAML/JSON → dict)
│   ├── validator/            # 3-pass validation (schema, business, assets)
│   ├── parser/               # dict → Pydantic models
│   ├── models/               # ThemeModel, DeckModel, 7 SlideModels
│   ├── theme/                # ThemeLoader
│   ├── layout/               # LayoutRegistry, BaseLayoutRenderer, 7 renderers
│   ├── renderer/             # DeckRenderer (pipeline orchestrator)
│   ├── exporter/             # PptxExporter (atomic write)
│   └── logger/               # DeckLogger (console + file)
├── tests/
│   ├── unit/                 # Per-module unit tests
│   └── integration/          # Full pipeline + determinism tests
└── tools/create_template.py  # Bootstrap script for assets/template.pptx
```

---

## Constraints

This tool intentionally does **not** support:

- GUI or web interface
- AI content generation
- CRM / ERP integrations
- Multi-brand support
- Animations or transitions
- Real-time collaboration

These are by design, not limitations.
