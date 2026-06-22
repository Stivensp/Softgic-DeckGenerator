"""
Creates a minimal assets/template.pptx with proper widescreen dimensions.
Run this once before using the generator if no corporate template is available:

    python tools/create_template.py

The resulting template.pptx contains a standard blank layout at every index
(0-10), so all layout_indices in theme.yaml will resolve correctly.
When the corporate Softgic template is provided, replace assets/template.pptx
with it and update layout_indices in config/theme.yaml accordingly.
"""
from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches


def create_template(output_path: Path = Path("assets/template.pptx")) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    prs = Presentation()

    # Standard 16:9 widescreen dimensions
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    # The default Presentation already contains slide layouts 0-10.
    # We don't add slides — the template only holds layouts.
    # Renderers will call prs.slides.add_slide(prs.slide_layouts[6]) for each slide.

    prs.save(str(output_path))
    print(f"Template created: {output_path}")
    print(f"Available layouts: {len(prs.slide_layouts)}")
    print("Layout 6 = blank (default for all slide types)")


def create_placeholder_images(images_dir: Path = Path("images")) -> None:
    """Creates minimal placeholder PNG files for logo and profile photo."""
    images_dir.mkdir(parents=True, exist_ok=True)

    try:
        from PIL import Image, ImageDraw
    except ImportError:
        print("Pillow not installed — skipping placeholder image creation.")
        return

    # Logo white (white text on transparent bg)
    _create_logo(images_dir / "logo_white.png", bg=(10, 22, 40), fg=(255, 255, 255))
    # Logo dark (dark text on white bg)
    _create_logo(images_dir / "logo_dark.png", bg=(255, 255, 255), fg=(10, 22, 40))
    # Placeholder profile photo
    _create_profile_placeholder(images_dir / "placeholder_profile.png")


def _create_logo(path: Path, bg: tuple, fg: tuple) -> None:
    from PIL import Image, ImageDraw, ImageFont

    img = Image.new("RGBA", (440, 140), color=(*bg, 255))
    draw = ImageDraw.Draw(img)
    draw.text((20, 40), "SOFTGIC", fill=(*fg, 255))
    img.save(str(path))
    print(f"Created: {path}")


def _create_profile_placeholder(path: Path) -> None:
    from PIL import Image, ImageDraw

    size = 400
    img = Image.new("RGB", (size, size), color=(200, 200, 200))
    draw = ImageDraw.Draw(img)
    # Simple person silhouette using circles and rectangles
    draw.ellipse([140, 60, 260, 180], fill=(160, 160, 160))
    draw.rectangle([80, 200, 320, 380], fill=(160, 160, 160))
    img.save(str(path))
    print(f"Created: {path}")


if __name__ == "__main__":
    create_template()
    create_placeholder_images(Path("images"))
    print("\nSetup complete. You can now run:")
    print("  python generate.py --input examples/propuesta_comercial.yaml --output output/deck.pptx")
