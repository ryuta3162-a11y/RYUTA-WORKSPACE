"""Generate Work-Space app icon — S:1 family (W:S + cyan arc, no subtitle)."""
from __future__ import annotations

import math
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter, ImageFont

ROOT = Path(__file__).resolve().parents[1]
PUBLIC = ROOT / "public"
ICONS = PUBLIC / "icons"

FONT_CANDIDATES = [
    Path(r"C:\Windows\Fonts\segoeuib.ttf"),
    Path(r"C:\Windows\Fonts\arialbd.ttf"),
    Path(r"C:\Windows\Fonts\calibrib.ttf"),
]


def load_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    for path in FONT_CANDIDATES:
        if path.exists():
            return ImageFont.truetype(str(path), size)
    return ImageFont.load_default()


def lerp(a: float, b: float, t: float) -> float:
    return a + (b - a) * t


def make_background(size: int) -> Image.Image:
    img = Image.new("RGBA", (size, size))
    px = img.load()
    cx, cy = size * 0.5, size * 0.46
    max_r = size * 0.78
    for y in range(size):
        for x in range(size):
            t = y / max(size - 1, 1)
            base = (
                int(lerp(10, 2, t)),
                int(lerp(20, 4, t)),
                int(lerp(38, 8, t)),
            )
            dx = x - cx
            dy = y - cy
            dist = math.sqrt(dx * dx + dy * dy) / max_r
            glow = max(0.0, 1.0 - dist) ** 1.55
            r = min(255, int(base[0] + glow * 16))
            g = min(255, int(base[1] + glow * 38))
            b = min(255, int(base[2] + glow * 64))
            px[x, y] = (r, g, b, 255)
    return img


def draw_arc_layer(
    size: int,
    box: tuple[int, int, int, int],
    width: int,
    color: tuple[int, int, int],
    blur: int,
    alpha: int,
) -> Image.Image:
    layer = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(layer)
    draw.arc(box, start=200, end=-20, fill=(*color, alpha), width=width)
    if blur > 0:
        layer = layer.filter(ImageFilter.GaussianBlur(radius=blur))
    return layer


def render_icon(size: int = 512) -> Image.Image:
    img = make_background(size)
    draw = ImageDraw.Draw(img)

    font_size = int(size * 0.36)
    font = load_font(font_size)
    label = "W:S"
    bbox = draw.textbbox((0, 0), label, font=font)
    tw = bbox[2] - bbox[0]
    tx = (size - tw) // 2 - bbox[0]
    ty = int(size * 0.34) - bbox[1]

    shadow = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    sdraw = ImageDraw.Draw(shadow)
    sdraw.text((tx, ty + max(2, size // 110)), label, font=font, fill=(0, 0, 0, 110))
    shadow = shadow.filter(ImageFilter.GaussianBlur(radius=max(2, size // 90)))
    img.alpha_composite(shadow)

    draw = ImageDraw.Draw(img)
    draw.text((tx, ty), label, font=font, fill=(248, 250, 252, 255))

    pad = int(size * 0.18)
    arc_top = int(size * 0.62)
    arc_bottom = int(size * 0.84)
    arc_box = (pad, arc_top, size - pad, arc_bottom)
    stroke = max(4, size // 38)

    img.alpha_composite(draw_arc_layer(size, arc_box, stroke + 12, (14, 165, 233), blur=max(3, size // 70), alpha=42))
    img.alpha_composite(draw_arc_layer(size, arc_box, stroke + 5, (56, 189, 248), blur=max(2, size // 110), alpha=90))
    draw = ImageDraw.Draw(img)
    draw.arc(arc_box, start=200, end=-20, fill=(125, 211, 252, 255), width=stroke)

    return img


def export_all() -> None:
    master = render_icon(512)
    targets = {
        ICONS / "icon-512.png": 512,
        ICONS / "icon-192.png": 192,
        ICONS / "apple-touch-icon.png": 180,
        PUBLIC / "favicon.png": 64,
    }
    for path, dim in targets.items():
        out = master.resize((dim, dim), Image.Resampling.LANCZOS)
        out.save(path, format="PNG", optimize=True)
        print(f"saved {path} ({dim}px)")


if __name__ == "__main__":
    export_all()
