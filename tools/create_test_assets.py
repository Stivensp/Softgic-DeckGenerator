"""
Creates polished placeholder assets for testing:
  images/logo_white.png          — logo sobre fondo oscuro
  images/logo_dark.png           — logo sobre fondo blanco
  images/placeholder_profile.png — avatar genérico profesional
  assets/template.pptx           — template con dimensiones 16:9

Run:  python tools/create_test_assets.py
"""
from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Inches

# ── Brand colors (must match config/theme.yaml) ───────────────────────────────
PRIMARY   = (10, 22, 40)       # #0A1628 navy
SECONDARY = (30, 58, 95)       # #1E3A5F mid-blue
ACCENT    = (0, 174, 239)      # #00AEEF Softgic blue
WHITE     = (255, 255, 255)
LIGHT_BG  = (245, 247, 250)    # near-white for light slides
GRAY_MID  = (107, 114, 128)    # #6B7280


def _draw_logo(draw: ImageDraw.Draw, x: int, y: int, w: int, h: int,
               text_color: tuple, accent: tuple) -> None:
    """Draws a simple 'S' glyph + SOFTGIC logotype."""
    # Accent square mark
    sq = int(h * 0.55)
    draw.rectangle([x, y + (h - sq) // 2, x + sq, y + (h - sq) // 2 + sq],
                   fill=accent)
    # White cut inside square
    inner = int(sq * 0.28)
    draw.rectangle([x + inner, y + (h - sq) // 2 + inner,
                    x + sq - inner, y + (h - sq) // 2 + sq - inner],
                   fill=text_color[:3] + (180,) if len(text_color) == 4 else (10, 22, 40))
    # Wordmark
    tx = x + sq + int(sq * 0.35)
    try:
        font = ImageFont.truetype("arial.ttf", int(h * 0.42))
    except (OSError, IOError):
        font = ImageFont.load_default()
    draw.text((tx, y + h // 2), "SOFTGIC", fill=text_color, font=font, anchor="lm")


def create_logo_white(path: Path) -> None:
    w, h = 480, 120
    img = Image.new("RGBA", (w, h), (0, 0, 0, 0))   # transparent background
    draw = ImageDraw.Draw(img)
    _draw_logo(draw, 20, 20, w - 40, h - 40, WHITE, ACCENT)
    img.save(str(path))
    print("  OK " + str(path))


def create_logo_dark(path: Path) -> None:
    w, h = 480, 120
    img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    _draw_logo(draw, 20, 20, w - 40, h - 40, PRIMARY, ACCENT)
    img.save(str(path))
    print("  OK " + str(path))


def create_placeholder_profile(path: Path) -> None:
    size = 500
    img = Image.new("RGB", (size, size), color=SECONDARY)
    draw = ImageDraw.Draw(img)

    # Circular face
    cx, cy = size // 2, size // 2
    face_r = int(size * 0.22)
    draw.ellipse(
        [cx - face_r, cy - int(size * 0.38),
         cx + face_r, cy - int(size * 0.38) + face_r * 2],
        fill=LIGHT_BG
    )
    # Body arc (shoulders)
    body_top = cy + int(size * 0.02)
    draw.ellipse(
        [cx - int(size * 0.38), body_top,
         cx + int(size * 0.38), body_top + int(size * 0.7)],
        fill=LIGHT_BG
    )
    # Accent ring border
    draw.ellipse([4, 4, size - 4, size - 4], outline=ACCENT, width=6)

    img.save(str(path))
    print("  OK " + str(path))


def create_template(path: Path) -> None:
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(path))
    print(f"  OK {path}  ({len(prs.slide_layouts)} layouts disponibles)")


if __name__ == "__main__":
    assets = Path("assets")
    images = Path("images")
    assets.mkdir(exist_ok=True)
    images.mkdir(exist_ok=True)

    print("\nGenerando assets de test...\n")
    create_template(assets / "template.pptx")
    create_logo_white(images / "logo_white.png")
    create_logo_dark(images / "logo_dark.png")
    create_placeholder_profile(images / "placeholder_profile.png")

    print("\nAssets listos. Ahora ejecuta:")
    print("  python generate.py --input examples/propuesta_comercial.yaml --output output/propuesta_comercial.pptx")
    print("  python generate.py --input examples/talent_profile.yaml      --output output/talent_profile.pptx")
