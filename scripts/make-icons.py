#!/usr/bin/env python3
"""Иконки надстройки Excel (16/32/64/80) без внешних зависимостей, кроме Pillow."""
from pathlib import Path

from PIL import Image, ImageDraw

OUT = Path(__file__).resolve().parent.parent / "addin" / "assets"
OUT.mkdir(parents=True, exist_ok=True)

BG = (17, 17, 24, 255)      # почти чёрный, как у Hermes
FG = (226, 178, 96, 255)    # золотой акцент


def icon(size: int) -> Image.Image:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    r = max(2, size // 6)
    d.rounded_rectangle([0, 0, size - 1, size - 1], radius=r, fill=BG, outline=FG,
                        width=max(1, size // 16))
    # буква H из трёх прямоугольников
    w = max(2, size // 6)
    top = int(size * 0.28)
    bot = int(size * 0.72)
    left = int(size * 0.3)
    right = int(size * 0.7)
    d.rectangle([left, top, left + w - 1, bot], fill=FG)
    d.rectangle([right - w + 1, top, right, bot], fill=FG)
    mid = (top + bot) // 2
    d.rectangle([left, mid - max(1, w // 3), right, mid + max(1, w // 3)], fill=FG)
    return img


for s in (16, 32, 64, 80):
    icon(s).save(OUT / f"icon-{s}.png")
    print("wrote", OUT / f"icon-{s}.png")
