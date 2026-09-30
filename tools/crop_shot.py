"""Crop a Power BI Desktop canvas screenshot to the report page and resize it to 1280 px wide.

The page background colour (#F4F6F6) marks the page; the crop is its bounding box.
Usage: python crop_shot.py in.png out.png
Author: Badr Abardazzou. License: MIT.
"""
import sys

from PIL import Image

BG = (0xF4, 0xF6, 0xF6)


def crop(src, dst, width=1280):
    im = Image.open(src).convert("RGB")
    w, h = im.size
    px = im.load()

    def is_bg(x, y):
        r, g, b = px[x, y]
        return (r, g, b) == BG

    # the bottom margin strip of the page is plain background across the full page width
    rows = [y for y in range(h) if sum(is_bg(x, y) for x in range(0, w, 5)) > w / 5 * 0.6]
    bottom = max(rows)
    y = bottom - 3
    xs = [x for x in range(w) if is_bg(x, y)]
    left, right = min(xs), max(xs)
    top = bottom - round((right - left + 1) * 720 / 1280) + 1  # page is 1280 x 720
    page = im.crop((left, top, right + 1, bottom + 1))
    ratio = width / page.width
    page = page.resize((width, round(page.height * ratio)), Image.LANCZOS)
    page.save(dst, optimize=True)
    print(dst, page.size, (left, top, right, bottom))


if __name__ == "__main__":
    crop(sys.argv[1], sys.argv[2])
