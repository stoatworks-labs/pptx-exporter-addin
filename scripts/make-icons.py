#!/usr/bin/env python3
"""Draw the add-in's icons (public/icons/*.png and icon.svg).

A slide with an arrow leaving it, on the Stoatworks rust orange. Drawn, not
an SF Symbol or Fluent icon — neither licence allows them in an app's icon.

    .venv/bin/python scripts/make-icons.py
"""
import os
from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, '..', 'public', 'icons')
BG = (168, 84, 28)
BG2 = (196, 108, 46)
INK = (255, 255, 255)
SHADE = (250, 226, 208)

S = 1024  # drawn large, scaled down


def draw():
    im = Image.new('RGBA', (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    # Rounded square with a gentle vertical gradient.
    grad = Image.new('RGBA', (S, S))
    gd = ImageDraw.Draw(grad)
    for y in range(S):
        t = y / S
        c = tuple(int(BG2[i] * (1 - t) + BG[i] * t) for i in range(3)) + (255,)
        gd.line([(0, y), (S, y)], fill=c)
    mask = Image.new('L', (S, S), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, S - 1, S - 1], radius=int(S * 0.22), fill=255)
    im.paste(grad, (0, 0), mask)
    # The slide: a 16:9 card, lower left.
    x0, y0, w = int(S * 0.14), int(S * 0.40), int(S * 0.56)
    h = int(w * 9 / 16)
    d.rounded_rectangle([x0, y0, x0 + w, y0 + h], radius=int(S * 0.035), fill=INK)
    # Two lines of "text" on it.
    lw = int(S * 0.03)
    d.rounded_rectangle([x0 + int(w * 0.12), y0 + int(h * 0.25), x0 + int(w * 0.70), y0 + int(h * 0.25) + lw], radius=lw // 2, fill=BG)
    d.rounded_rectangle([x0 + int(w * 0.12), y0 + int(h * 0.50), x0 + int(w * 0.52), y0 + int(h * 0.50) + lw], radius=lw // 2, fill=SHADE)
    d.rounded_rectangle([x0 + int(w * 0.12), y0 + int(h * 0.68), x0 + int(w * 0.44), y0 + int(h * 0.68) + lw], radius=lw // 2, fill=SHADE)
    # The arrow: out of the slide's top-right corner, up and right.
    ax0, ay0 = x0 + int(w * 0.62), y0 + int(h * 0.05)
    ax1, ay1 = int(S * 0.84), int(S * 0.16)
    t = int(S * 0.075)
    d.line([(ax0, ay0), (ax1, ay1)], fill=INK, width=t)
    head = int(S * 0.20)
    d.polygon([(ax1 + int(t * 0.35), ay1 - int(t * 0.35)), (ax1 - head, ay1 + int(head * 0.08)), (ax1 - int(head * 0.08), ay1 + head)], fill=INK)
    d.ellipse([ax0 - t // 2, ay0 - t // 2, ax0 + t // 2, ay0 + t // 2], fill=INK)
    return im


def main():
    os.makedirs(OUT, exist_ok=True)
    big = draw()
    for n in (16, 32, 64, 80, 128, 192, 512):
        big.resize((n, n), Image.LANCZOS).save(os.path.join(OUT, f'icon-{n}.png'))
    print('wrote', OUT)


if __name__ == '__main__':
    main()
