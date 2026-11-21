# -*- coding: utf-8 -*-
"""Generate the multi-size app icon used by the desktop shortcut.

Optional helper: the repository already ships assets/app.ico, so you only need
to run this when you want to change the artwork. Requires Pillow.
"""
from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw

OUT = Path(__file__).resolve().parent / "app.ico"
S = 512
SIZES = [16, 24, 32, 48, 64, 128, 256]

TOP = (47, 107, 255)
BOTTOM = (0, 180, 216)
WHITE = (255, 255, 255)
LINE = (35, 96, 210)

image = Image.new("RGBA", (S, S), (0, 0, 0, 0))
gradient = Image.new("RGBA", (S, S))
draw_gradient = ImageDraw.Draw(gradient)
for y in range(S):
    ratio = y / (S - 1)
    color = tuple(int(TOP[i] * (1 - ratio) + BOTTOM[i] * ratio) for i in range(3)) + (255,)
    draw_gradient.line([(0, y), (S, y)], fill=color)

mask = Image.new("L", (S, S), 0)
mask_draw = ImageDraw.Draw(mask)
mask_draw.rounded_rectangle((18, 18, S - 18, S - 18), radius=108, fill=255)
image.paste(gradient, (0, 0), mask)

draw = ImageDraw.Draw(image)

# White chat bubble.
draw.rounded_rectangle((96, 118, 416, 338), radius=62, fill=WHITE)
draw.polygon([(154, 322), (154, 398), (238, 322)], fill=WHITE)

# Three short blue lines inside the bubble suggest a conversation.
draw.rounded_rectangle((142, 182, 322, 202), radius=10, fill=LINE)
draw.rounded_rectangle((142, 226, 286, 246), radius=10, fill=(80, 145, 235))
draw.rounded_rectangle((142, 270, 248, 290), radius=10, fill=(132, 184, 245))

# Small export arrow badge in the lower-right corner.
draw.ellipse((318, 300, 438, 420), fill=(7, 156, 224), outline=WHITE, width=10)
draw.line([(378, 392), (378, 334)], fill=WHITE, width=18, joint="curve")
draw.polygon([(344, 360), (378, 326), (412, 360)], fill=WHITE)

OUT.parent.mkdir(parents=True, exist_ok=True)
image.save(OUT, format="ICO", sizes=[(size, size) for size in SIZES])
print(f"ICON={OUT}")
print(f"SIZE={OUT.stat().st_size}")
print("SIZES=" + ",".join(str(size) for size in SIZES))
