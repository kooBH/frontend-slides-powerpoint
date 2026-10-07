#!/usr/bin/env python3
"""
make-background.py — Generate atmospheric background images for slides.

pptxgenjs cannot write gradient fills, noise, or pattern overlays. The skill's
design philosophy still wants "atmosphere and depth rather than solid colors",
so this script renders those effects to a PNG that becomes the slide (or layout)
background via `background: { path: "bg.png" }`.

Usage:
    python scripts/make-background.py out.png --gradient 1a1a1a,2d2d2d,1a1a1a --angle 135
    python scripts/make-background.py out.png --solid 0f0f0f --glow e8b4b8@80%,20%:0.35 --glow d4a574@15%,85%:0.25
    python scripts/make-background.py out.png --solid 1C2644 --grid 80:ffffff:0.03
    python scripts/make-background.py out.png --solid f5f3ee --noise 0.04
    python scripts/make-background.py out.png --solid ffffff --dots 40:1a1a1a:0.12
    python scripts/make-background.py out.png --solid 0a0f1c --scanlines 4:000000:0.18

Options:
    --size WxH            pixel size (default 1920x1080; always 16:9 for LAYOUT_WIDE)
    --solid HEX           base color
    --gradient HEX,HEX[,HEX...]   linear gradient between stops
    --angle DEG           gradient direction (0 = left→right, 90 = top→bottom, 135 = diagonal)
    --glow HEX@X%,Y%:ALPHA[:RADIUS%]   soft radial blob (repeatable) — the "abstract soft shapes" device
    --grid PX:HEX:ALPHA   square grid lines every PX pixels
    --dots PX:HEX:ALPHA   dot grid every PX pixels
    --noise ALPHA         film grain
    --scanlines PX:HEX:ALPHA   horizontal lines every PX pixels (CRT / terminal feel)
    --vignette ALPHA      darken the edges

Requires: pip install Pillow
"""

import argparse
import math
import random
import sys

try:
    from PIL import Image, ImageDraw, ImageFilter
except ImportError:
    print("Pillow is required: pip install Pillow")
    sys.exit(1)


def hex_rgb(h):
    h = h.strip().lstrip("#")
    if len(h) != 6:
        raise SystemExit(f"bad hex color: {h}")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def gradient(size, stops, angle):
    w, h = size
    img = Image.new("RGB", size)
    px = img.load()
    rad = math.radians(angle)
    dx, dy = math.cos(rad), math.sin(rad)
    # project corners to find the extent along the gradient axis
    corners = [(0, 0), (w, 0), (0, h), (w, h)]
    proj = [x * dx + y * dy for x, y in corners]
    lo, hi = min(proj), max(proj)
    n = len(stops) - 1
    # compute a 1-D lookup table to keep it fast
    lut = []
    for i in range(1024):
        t = i / 1023 * n
        k = min(int(t), n - 1)
        f = t - k
        a, b = stops[k], stops[k + 1]
        lut.append(tuple(round(a[c] + (b[c] - a[c]) * f) for c in range(3)))
    for y in range(h):
        for x in range(w):
            t = (x * dx + y * dy - lo) / (hi - lo) if hi > lo else 0
            px[x, y] = lut[int(t * 1023)]
    return img


def overlay(base, layer_fn, alpha):
    layer = Image.new("RGBA", base.size, (0, 0, 0, 0))
    layer_fn(ImageDraw.Draw(layer), layer)
    a = layer.split()[3].point(lambda v: int(v * alpha))
    layer.putalpha(a)
    return Image.alpha_composite(base.convert("RGBA"), layer)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("out")
    ap.add_argument("--size", default="1920x1080")
    ap.add_argument("--solid")
    ap.add_argument("--gradient")
    ap.add_argument("--angle", type=float, default=135)
    ap.add_argument("--glow", action="append", default=[])
    ap.add_argument("--grid")
    ap.add_argument("--dots")
    ap.add_argument("--noise", type=float, default=0)
    ap.add_argument("--scanlines")
    ap.add_argument("--vignette", type=float, default=0)
    args = ap.parse_args()

    w, h = (int(v) for v in args.size.lower().split("x"))
    size = (w, h)

    if args.gradient:
        img = gradient(size, [hex_rgb(s) for s in args.gradient.split(",")], args.angle)
    else:
        img = Image.new("RGB", size, hex_rgb(args.solid or "ffffff"))
    img = img.convert("RGBA")

    for spec in args.glow:
        # HEX@X%,Y%:ALPHA[:RADIUS%]
        color, rest = spec.split("@", 1)
        pos, _, tail = rest.partition(":")
        xs, ys = pos.split(",")
        parts = tail.split(":")
        alpha = float(parts[0]) if parts and parts[0] else 0.3
        radius = float(parts[1].rstrip("%")) / 100 if len(parts) > 1 else 0.35
        cx, cy = float(xs.rstrip("%")) / 100 * w, float(ys.rstrip("%")) / 100 * h
        r = radius * max(w, h)
        blob = Image.new("RGBA", size, (0, 0, 0, 0))
        ImageDraw.Draw(blob).ellipse([cx - r, cy - r, cx + r, cy + r], fill=hex_rgb(color) + (255,))
        blob = blob.filter(ImageFilter.GaussianBlur(r * 0.55))
        a = blob.split()[3].point(lambda v: int(v * alpha))
        blob.putalpha(a)
        img = Image.alpha_composite(img, blob)

    if args.grid:
        step, color, alpha = args.grid.split(":")
        step, rgb, alpha = int(step), hex_rgb(color), float(alpha)

        def draw_grid(d, _):
            for x in range(0, w, step):
                d.line([(x, 0), (x, h)], fill=rgb + (255,), width=1)
            for y in range(0, h, step):
                d.line([(0, y), (w, y)], fill=rgb + (255,), width=1)
        img = overlay(img, draw_grid, alpha)

    if args.dots:
        step, color, alpha = args.dots.split(":")
        step, rgb, alpha = int(step), hex_rgb(color), float(alpha)

        def draw_dots(d, _):
            for x in range(step // 2, w, step):
                for y in range(step // 2, h, step):
                    d.ellipse([x - 1.5, y - 1.5, x + 1.5, y + 1.5], fill=rgb + (255,))
        img = overlay(img, draw_dots, alpha)

    if args.scanlines:
        step, color, alpha = args.scanlines.split(":")
        step, rgb, alpha = int(step), hex_rgb(color), float(alpha)

        def draw_lines(d, _):
            for y in range(0, h, step):
                d.line([(0, y), (w, y)], fill=rgb + (255,), width=1)
        img = overlay(img, draw_lines, alpha)

    if args.noise > 0:
        rnd = random.Random(7)
        noise = Image.effect_noise(size, 64).convert("L")
        layer = Image.new("RGBA", size, (255, 255, 255, 0))
        layer.putalpha(noise.point(lambda v: int(abs(v - 128) * 2 * args.noise)))
        dark = Image.new("RGBA", size, (0, 0, 0, 0))
        dark.putalpha(noise.point(lambda v: int(max(0, 128 - v) * 2 * args.noise)))
        img = Image.alpha_composite(Image.alpha_composite(img, layer), dark)

    if args.vignette > 0:
        mask = Image.new("L", size, 0)
        d = ImageDraw.Draw(mask)
        d.ellipse([-w * 0.2, -h * 0.35, w * 1.2, h * 1.35], fill=255)
        mask = mask.filter(ImageFilter.GaussianBlur(min(w, h) * 0.25))
        shade = Image.new("RGBA", size, (0, 0, 0, int(255 * args.vignette)))
        shade.putalpha(mask.point(lambda v: int((255 - v) * args.vignette)))
        img = Image.alpha_composite(img, shade)

    if args.out.lower().endswith((".jpg", ".jpeg")):
        img.convert("RGB").save(args.out, "JPEG", quality=88, optimize=True)   # much smaller for noise/glow
    else:
        img.convert("RGB").save(args.out, "PNG", optimize=True)
    print(args.out)


if __name__ == "__main__":
    main()
