"""
Softgic Deck Generator — CLI entry point.

Usage:
    python generate.py --input deck.yaml --output output.pptx
    python generate.py --input deck.json --output output.pptx --theme config/theme.yaml
    python generate.py --version
"""
from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

_VERSION = "1.0.0"
_DEFAULT_THEME = Path("config/theme.yaml")
_DEFAULT_TEMPLATE_DIR = Path("assets")


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


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Softgic Deck Generator — generate branded PowerPoint decks from YAML/JSON"
    )
    parser.add_argument("--input", "-i", required=False, help="Input YAML or JSON file")
    parser.add_argument("--output", "-o", required=False, default="output.pptx", help="Output .pptx file path")
    parser.add_argument("--theme", "-t", default=str(_DEFAULT_THEME), help="Path to theme.yaml")
    parser.add_argument("--version", action="store_true", help="Print version and exit")
    args = parser.parse_args()

    if args.version:
        print(f"Softgic Deck Generator v{_VERSION}")
        return 0

    if not args.input:
        parser.error("--input is required")

    # ── Initialize logger ──────────────────────────────────────────────────────
    from src.logger.deck_logger import DeckLogger

    deck_logger = DeckLogger()
    deck_logger.init()
    log = deck_logger.get("main")

    log.info("=" * 60)
    log.info("Softgic Deck Generator v%s", _VERSION)
    log.info("Input  : %s", args.input)
    log.info("Output : %s", args.output)
    log.info("Theme  : %s", args.theme)
    log.info("=" * 60)

    try:
        from src.exceptions import SoftgicDeckError

        # ── 1. Read input ──────────────────────────────────────────────────────
        from src.input.file_reader import FileReader

        raw = FileReader().read(Path(args.input))

        # ── 2. Load theme ──────────────────────────────────────────────────────
        from src.theme.theme_loader import ThemeLoader

        theme = ThemeLoader().load(Path(args.theme))

        # ── 3. Validate ────────────────────────────────────────────────────────
        from src.validator.deck_validator import DeckValidator

        DeckValidator().validate(raw, dict(theme.limits.model_dump()))

        # ── 4. Parse ───────────────────────────────────────────────────────────
        from src.parser.deck_parser import DeckParser

        deck = DeckParser().parse(raw)

        from src.models.deck import DeckMetadata

        deck.metadata = DeckMetadata(
            generated_at=datetime.now().isoformat(timespec="seconds"),
            input_file=args.input,
            theme_version=theme.version,
        )

        # ── 5. Load template ───────────────────────────────────────────────────
        from pptx import Presentation

        template_path = Path(theme.assets.template)
        if template_path.exists():
            log.info("Loading template: %s", template_path)
            presentation = Presentation(str(template_path))
        else:
            log.warning(
                "Template '%s' not found — creating blank presentation. "
                "Run tools/create_template.py to generate it.",
                template_path,
            )
            presentation = Presentation()

        # ── 6. Render ──────────────────────────────────────────────────────────
        from src.renderer.deck_renderer import DeckRenderer

        registry = _build_registry()
        presentation = DeckRenderer().render(deck, theme, registry, presentation)

        # ── 7. Export ──────────────────────────────────────────────────────────
        from src.exporter.pptx_exporter import PptxExporter

        PptxExporter().export(presentation, Path(args.output))

        log.info("Done. Deck saved to: %s", args.output)
        return 0

    except Exception as exc:
        import traceback

        log.error("Pipeline failed: %s", exc)
        log.debug(traceback.format_exc())
        print(f"\nERROR: {exc}\n", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
