"""Cuts the courier logo sheet (2 columns x 5 rows of logo strips) into one PNG per service group, the logos each group
uses until another is dropped for it (data/courier logos/<group id>.png). Only the size changes: each strip is cut out of the sheet
as it is (its own colours and design) and scaled to LOGO_WIDTH px wide, which prints sharp in the label's 43 mm logo
slot (as tall as the LABEL / PAGE chips) at 600 dpi and keeps every file small, so stitching stays quick.

    python -m engine.courier.logo_sheet SHEET.png          (writes the ten PNGs and a check sheet, _cut_check.png)

The sheet's order, left column then right column, top to bottom (group ids from courier_services.DEFAULT_GROUPS):
"""
import os
import sys

from PIL import Image

from engine.core.paths import DATA_DIR
from engine.courier.courier_services import LOGO_DIR_NAME

LOGO_DIR = os.path.join(DATA_DIR, LOGO_DIR_NAME)
LOGO_WIDTH = 1000  # px: the 43 mm slot at 600 dpi
ORDER = [  # (column, row) -> group id
    ('dhl', 0, 0), ('fedex-international', 0, 1), ('flash-couriers', 0, 2), ('startrack-premium-overnight', 0, 3),
    ('startrack-road-express', 0, 4),
    ('tnt-overnight', 1, 0), ('tnt-international', 1, 1), ('tnt-nsw-road-express', 1, 2),
    ('toll-ipec-road-express', 1, 3),  # Team Global Express IPEC Road Express
    ('toll-priority', 1, 4),           # Team Global Express Priority
]
__doc__ += "\n".join(f"  column {c + 1}, row {r + 1}: {gid}" for gid, c, r in ORDER) + "\n"


def _background(img):
    """The sheet's own background colour: the most common colour along its outer edge."""
    w, h = img.size
    edge = [img.getpixel((x, 0)) for x in range(w)] + [img.getpixel((x, h - 1)) for x in range(w)] + \
           [img.getpixel((0, y)) for y in range(h)] + [img.getpixel((w - 1, y)) for y in range(h)]
    return max(set(edge), key=edge.count)


def _runs(flags):
    """[(start, end)] of the runs of True in a list."""
    out, start = [], None
    for i, f in enumerate(flags + [False]):
        if f and start is None:
            start = i
        elif not f and start is not None:
            out.append((start, i))
            start = None
    return out


def cells(img, cols=2, rows=5):
    """The (left, top, right, bottom) of each strip, found from the gaps of background colour between them; falls back
    to an even split of the sheet's used area when the gaps can't be told apart (a strip with a background-coloured
    edge)."""
    rgb = img.convert('RGB')
    bg = _background(rgb)
    w, h = rgb.size
    near = lambda p: sum(abs(a - b) for a, b in zip(p, bg)) < 30
    px = rgb.load()
    col_used = [sum(not near(px[x, y]) for y in range(0, h, 2)) > h * 0.02 for x in range(w)]
    row_used = [sum(not near(px[x, y]) for x in range(0, w, 2)) > w * 0.02 for y in range(h)]
    xs = [r for r in _runs(col_used) if r[1] - r[0] > w * 0.1]
    ys = [r for r in _runs(row_used) if r[1] - r[0] > h * 0.05]
    if len(xs) != cols:
        x0, x1 = (xs[0][0], xs[-1][1]) if xs else (0, w)
        step = (x1 - x0) / cols
        xs = [(round(x0 + i * step), round(x0 + (i + 1) * step)) for i in range(cols)]
    if len(ys) != rows:
        y0, y1 = (ys[0][0], ys[-1][1]) if ys else (0, h)
        step = (y1 - y0) / rows
        ys = [(round(y0 + i * step), round(y0 + (i + 1) * step)) for i in range(rows)]
    return {(c, r): (xs[c][0], ys[r][0], xs[c][1], ys[r][1]) for c in range(cols) for r in range(rows)}


def cut(sheet_path, out_dir=LOGO_DIR, width=LOGO_WIDTH):
    img = Image.open(sheet_path)
    img.load()
    img = img.convert('RGBA')
    boxes = cells(img)
    os.makedirs(out_dir, exist_ok=True)
    made = []
    for gid, c, r in ORDER:
        strip = img.crop(boxes[(c, r)])
        h = max(1, round(strip.height * width / strip.width))
        strip = strip.resize((width, h), Image.LANCZOS) if strip.width != width else strip
        path = os.path.join(out_dir, f"{gid}.png")
        strip.save(path, 'PNG', optimize=True)
        made.append((gid, path, strip.size, os.path.getsize(path)))
    # A check sheet of what was cut, to look at before using them
    pad = 10
    check = Image.new('RGBA', (width * 2 + pad * 3, sum(max(s[1] for _, _, s, _ in made[i::5]) for i in range(5)) + pad * 6), 'white')
    y = [pad] * 2
    for n, (gid, path, size, _) in enumerate(made):
        col = 0 if n < 5 else 1
        check.paste(Image.open(path), (pad + col * (width + pad), y[col]))
        y[col] += size[1] + pad
    # Next to the sheet, not in the logo folder: only logos go there
    check.save(os.path.join(os.path.dirname(os.path.abspath(sheet_path)), '_cut_check.png'))
    return made


if __name__ == '__main__':
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    for gid, path, size, nbytes in cut(sys.argv[1]):
        print(f"{gid:30s} {size[0]} x {size[1]} px  {nbytes / 1024:.1f} KB  {path}")
    print("check sheet:", os.path.join(os.path.dirname(os.path.abspath(sys.argv[1])), '_cut_check.png'))
