#!/usr/bin/env python3
"""Build the Quizboard app icon: a mortarboard with a bold gold Q on top.

Dev tool - needs Pillow.  Quizboard never imports it; the icon ships as plain
PNGs in icons/.  Colours are the board's own (tile blue, gold, navy).

Usage:
    python3 make_icon.py                     use DejaVu Sans Bold (or the best bold found)
    python3 make_icon.py --font PATH.ttf     use this font file

After rebuilding, run ./install-desktop.sh to refresh the dash icon.
"""
from __future__ import annotations

import argparse
import glob
import os
import sys

from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "icons")

TILE = (31, 61, 153)        # quizboard TILE
TILE_EDGE = (58, 79, 156)   # quizboard RULE
NAVY = (11, 18, 51)         # quizboard BACKGROUND
BOARD_TOP = (22, 32, 82)
BOARD_SIDE = (6, 10, 30)
GOLD = (255, 204, 51)       # quizboard GOLD
GOLD_DARK = (204, 150, 20)
SIZE = 512
SS = 4                      # supersampling factor
RADIUS = 84

SEARCH = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
    "/usr/share/fonts/**/*Sans*-Bold.[tTo][tT][fF]",
]


def find_font() -> str:
    for pattern in SEARCH:
        hits = sorted(glob.glob(pattern, recursive=True))
        if hits:
            return hits[0]
    sys.exit("no bold font found; pass --font PATH.ttf")


def s(*xy: float) -> list[float]:
    return [v * SS for v in xy]


def build(font_path: str) -> Image.Image:
    big = SIZE * SS
    img = Image.new("RGBA", (big, big), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    # Background: a quizboard tile.
    d.rounded_rectangle(s(8, 8, 504, 504), radius=RADIUS * SS, fill=TILE,
                        outline=TILE_EDGE, width=6 * SS)

    # Skullcap: a band under the board, curved at the bottom.
    d.polygon(s(140, 236, 372, 236, 372, 360, 140, 360), fill=NAVY)
    d.ellipse(s(140, 318, 372, 402), fill=NAVY)
    d.ellipse(s(140, 300, 372, 384), fill=BOARD_TOP)    # inner shadow line
    d.ellipse(s(146, 296, 366, 378), fill=NAVY)

    # The board: a square seen from above at an angle, with a thick edge.
    top, right, bottom, left = (256, 92), (480, 196), (256, 300), (32, 196)
    thick = 18
    d.polygon(s(*left, *bottom, *right, right[0], right[1] + thick,
                bottom[0], bottom[1] + thick, left[0], left[1] + thick), fill=BOARD_SIDE)
    d.polygon(s(*top, *right, *bottom, *left), fill=BOARD_TOP)
    d.line(s(*left, *top, *right), fill=(46, 62, 130), width=3 * SS)

    # The Q, foreshortened to lie on the board.
    q_size = 230 * SS
    font = ImageFont.truetype(font_path, q_size)
    q = Image.new("RGBA", (q_size * 2, q_size * 2), (0, 0, 0, 0))
    qd = ImageDraw.Draw(q)
    qd.text((q_size, q_size), "Q", font=font, fill=GOLD, anchor="mm")
    q = q.crop(q.getbbox())
    q = q.resize((q.width, int(q.height * 0.62)), Image.LANCZOS)
    cx, cy = 256 * SS, 192 * SS
    shadow = Image.new("RGBA", q.size, (0, 0, 0, 0))
    shadow.paste((*GOLD_DARK, 255), mask=q.getchannel("A"))
    img.alpha_composite(shadow, (cx - q.width // 2, cy - q.height // 2 + 4 * SS))
    img.alpha_composite(q, (cx - q.width // 2, cy - q.height // 2))

    # Tassel: button near the right corner, cord over the edge, gold fringe.
    d.ellipse(s(404, 186, 424, 202), fill=GOLD_DARK)
    d.line(s(414, 194, 448, 210, 450, 330), fill=GOLD, width=9 * SS, joint="curve")
    d.ellipse(s(438, 318, 462, 340), fill=GOLD)
    d.polygon(s(438, 334, 462, 334, 474, 420, 426, 420), fill=GOLD)
    for x in range(432, 470, 9):
        d.line(s(x, 352, x - 2 + (x - 450) * 0.15, 420), fill=GOLD_DARK, width=2 * SS)

    return img.resize((SIZE, SIZE), Image.LANCZOS)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--font", help="bold font file for the Q")
    args = ap.parse_args()
    font = args.font or find_font()
    os.makedirs(OUT, exist_ok=True)
    icon = build(font)
    icon.save(os.path.join(OUT, "quizboard.png"))
    for size in (256, 128, 64, 48):
        icon.resize((size, size), Image.LANCZOS).save(os.path.join(OUT, f"quizboard-{size}.png"))
    print(f"icons written to {OUT} (font: {font})")


if __name__ == "__main__":
    main()
