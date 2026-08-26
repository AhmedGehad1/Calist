"""Draw the Calist app icon.

    python docs/make_icon.py

Writes docs/calist.ico (embedded in the executable and used as the window
icon) and docs/calist-icon.png (the README header). Needs Pillow, which is a
development dependency only and is deliberately not in requirements.txt — the
app itself never draws the icon, it just ships the file.

Two things here are deliberate:

* **Each size is drawn at its own geometry, not downsampled from one master.**
  A four-row list with hairline bars is legible at 256px and turns to mush at
  16px. The small renditions drop to three rows with much thicker strokes, so
  the taskbar icon still reads as a list.

* **The tile carries a lifted rim.** The fill is near-black, and the Windows 11
  taskbar is too; without a lighter edge the icon dissolves into its own
  background.
"""
from pathlib import Path

from PIL import Image, ImageDraw

DOCS = Path(__file__).resolve().parent
ICO = DOCS / "calist.ico"
PNG = DOCS / "calist-icon.png"

# The app's own palette, so the icon and the window agree.
BG = (31, 31, 37, 255)        # SURFACE
EDGE = (72, 78, 96, 255)      # lifted off BORDER so a dark taskbar can't eat it
BLUE = (76, 141, 255, 255)    # PRIMARY
MUTED = (128, 132, 148, 255)  # the one row still outstanding

#: Sizes an .ico should carry, and how much detail each one can hold.
#: Windows picks 16 and 32 for the taskbar and Explorer's small views, which is
#: where all the legibility is won or lost.
SIZES = [256, 128, 64, 48, 32, 24, 16]

#: Row geometry as fractions of the tile, per level of detail.
#: `rows` is (bullet + bar) count; the last row is muted and short, which stops
#: the block reading as a solid rectangle and keeps a hint of "still compiling".
DETAILED = dict(rows=4, left=0.225, bar_x=0.375, bar_end=0.785, short_end=0.63,
                row_h=0.068, gap=0.133, radius=0.22, border=0.010)
SIMPLE = dict(rows=3, left=0.195, bar_x=0.395, bar_end=0.815, short_end=0.66,
              row_h=0.105, gap=0.235, radius=0.20, border=0.020)


def draw(size: int, *, supersample: int = 8) -> Image.Image:
    """Render one square icon at `size`, drawn large and reduced once."""
    g = DETAILED if size >= 48 else SIMPLE
    s = size * supersample
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    border = max(1, int(s * g["border"]))
    d.rounded_rectangle([0, 0, s - 1, s - 1], radius=int(s * g["radius"]),
                        fill=BG, outline=EDGE, width=border)

    row_h = int(s * g["row_h"])
    gap = int(s * g["gap"])
    # Centre the stack vertically rather than pinning it to a fixed top: with
    # the check badge gone there is nothing weighting the lower right.
    top = (s - ((g["rows"] - 1) * gap + row_h)) // 2
    left = int(s * g["left"])

    for i in range(g["rows"]):
        y = top + i * gap
        last = i == g["rows"] - 1
        colour = MUTED if last else BLUE
        d.ellipse([left, y, left + row_h, y + row_h], fill=colour)
        end = int(s * (g["short_end"] if last else g["bar_end"]))
        d.rounded_rectangle(
            [int(s * g["bar_x"]), y + int(row_h * 0.18),
             end, y + int(row_h * 0.82)],
            radius=row_h // 3, fill=colour)

    return img.resize((size, size), Image.LANCZOS)


def main() -> None:
    # save() re-renders from the base image for every entry in `sizes`, which
    # would throw away the per-size geometry. Hand it the finished renditions
    # via append_images instead, so each one is stored exactly as drawn.
    renditions = [draw(n) for n in SIZES]
    renditions[0].save(ICO, format="ICO", sizes=[(n, n) for n in SIZES],
                       append_images=renditions[1:])
    draw(512).save(PNG)
    print(f"wrote {ICO} ({', '.join(str(n) for n in SIZES)}) and {PNG}")


if __name__ == "__main__":
    main()
