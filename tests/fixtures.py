"""Sample files shaped like Packing Sheets' outputs, for the tests and for trying the horizontal labels by hand:

  build_final_file()       the 'Final <file>.xlsx': a horizontal distribution file with a 'Code for Pack <name>'
                           column at the start of each pack (and a sub-group code column in pack 2)
  build_signature_links()  the 'Signature links_<job>.xlsx': a tab of store codes per tab, the Divider Barcodes tab
                           and the Label Plugin tab

    python tests/fixtures.py OUT_DIR      writes both files there
"""
import io
import os
import re
import sys

import openpyxl
import pandas as pd
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image, ImageDraw

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from engine.core.core_math import (generate_pack_signatures, label_plugin_rows, LABEL_PLUGIN_SHEET, LABEL_PLUGIN_COLUMNS,  # noqa: E402
                       DIVIDER_BARCODE_SHEET, DIVIDER_BARCODE_COLUMNS)

ADDRESS = ["Store Name", "Attn", "Address 1", "Suburb", "State", "Postcode", "Install"]
TABS = {
    "Spring Sale": {
        "items": [  # pack, job, description, size, colour
            ("1", "J476699-01", "A1 poster", "594 x 841", (214, 69, 65)),
            ("1", "J476699-02", "A2 poster", "420 x 594", (232, 126, 4)),
            ("1", "J476699-03", "Window decal L", "1200 x 600", (41, 128, 185)),
            ("1", "J476699-03", "Window decal R", "1200 x 600", (52, 152, 219)),
            ("1", "J476699-04", "Ceiling wobbler", "300 dia", (142, 68, 173)),
            ("2", "J476699-05", "Counter card", "A4", (39, 174, 96)),
            ("2", "J476699-06", "Shelf strip", "900 x 40", (22, 160, 133)),
            ("2", "J476699-07", "Floor decal", "500 dia", (243, 156, 18)),
        ],
        "subgroups": {"2": [(1, 2)]},  # pack 2 split by Sub-Group into items 1-2
        "stores": [  # store, attn, address, suburb, state, postcode, install, spec per pack, qty per item
            ("Corio Village", "Store Manager", "Shop K001, Corio Village SC", "Corio", "VIC", "3214", "N", ["OB600100100", "CS450320150"], [2, 2, 1, 1, 4, 3, 10, 2]),
            ("Highpoint", "", "Shop 2102, 120-200 Rosamond Rd", "Maribyrnong", "VIC", "3032", "N", ["OB600100100", "CS450320150"], [2, 2, 1, 1, 4, 3, 10, 2]),
            ("Chadstone", "Sam Lee", "Shop B045, 1341 Dandenong Rd", "Malvern East", "VIC", "3145", "N", ["OB600100100", "2 x CS450320150"], [4, 4, 2, 2, 8, 6, 20, 4]),
            ("Westfield Parramatta", "", "Shop 4011, 159-175 Church St", "Parramatta", "NSW", "2150", "N", ["OB600100100", "CS450320150"], [1, 2, 1, 1, 2, 3, 10, 2]),
            ("Carindale", "", "Shop 1080, 1151 Creek Rd", "Carindale", "QLD", "4152", "N", ["OB600100100", "x"], [1, 2, 1, 1, 2, 0, 0, 0]),
            ("Rundle Mall", "Steven Priestley", "68 Ricketts Road", "Mount Waverley", "VIC", "3149", "Y", ["OB600100100", "CS450320150"], [2, 2, 1, 1, 4, 3, 10, 2]),
            ("Garden City", "", "Shop 2200, 2049 Logan Rd", "Upper Mount Gravatt", "QLD", "4122", "N", ["x", "CS450320150"], [0, 0, 0, 0, 0, 3, 10, 2]),
            ("Knox City", "", "Shop 3001, 425 Burwood Hwy", "Wantirna South", "VIC", "3152", "N", ["OB600100100", "CS450320150"], [2, 2, 1, 1, 4, 3, 5, 0]),
        ],
    },
    "Easter": {
        "items": [
            ("1", "J476700-01", "Egg hunt poster", "A2", (230, 81, 120)),
            ("1", "J476700-02", "Bunny standee", "1800 x 600", (120, 90, 200)),
        ],
        "subgroups": {},
        "installer": True,
        "stores": [
            ("Corio Village", "Store Manager", "Shop K001, Corio Village SC", "Corio", "VIC", "3214", "N", ["OB1800100100"], [1, 1]),
            ("Highpoint", "", "Shop 2102, 120-200 Rosamond Rd", "Maribyrnong", "VIC", "3032", "N", ["OB1800100100"], [2, 1]),
            ("Chadstone", "Sam Lee", "Shop B045, 1341 Dandenong Rd", "Malvern East", "VIC", "3145", "N", ["OB1800100100"], [1, 1]),
            ("Knox City", "", "Shop 3001, 425 Burwood Hwy", "Wantirna South", "VIC", "3152", "N", ["OB1800100100"], [2, 1]),
        ],
    },
}
JOB_ROW, PACK_ROW = 7, 8  # as in a Packing Sheets file: job numbers, then the merged pack headers just above the stores


def _art(text, colour):
    img = Image.new("RGB", (220, 300), colour)
    d = ImageDraw.Draw(img)
    d.rectangle([12, 12, 207, 287], outline="white", width=4)
    d.text((70, 140), text, fill="white")
    buf = io.BytesIO()
    img.save(buf, "PNG")
    buf.seek(0)
    return buf


def _layout(tab):
    """Raw (original file) and final (with code columns) positions of every item column, per pack."""
    spec = TABS[tab]
    packs = []
    for pack, *_ in spec["items"]:
        if not packs or packs[-1] != pack:
            packs.append(pack)
    first_raw = len(ADDRESS) + len(packs) + 1
    raw, col = {}, first_raw
    for i, (pack, *_rest) in enumerate(spec["items"]):
        raw.setdefault(pack, []).append(col)
        col += 1
    return packs, raw


def build_final_file(path, codes=True, change=None):
    """change: (tab, store row index, item index, new qty) applied after the codes were written (stale codes)."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    computed = {}
    for tab, spec in TABS.items():
        ws = wb.create_sheet(tab)
        packs, raw = _layout(tab)
        ws["A1"] = f"{spec['items'][0][1][:7]} - {tab} 2026 - Store rollout"
        ws["A1"].font = Font(bold=True, size=14)
        for c, h in enumerate(ADDRESS, 1):
            ws.cell(row=PACK_ROW, column=c, value=h)
        for n, pack in enumerate(packs):
            ws.cell(row=PACK_ROW, column=len(ADDRESS) + 1 + n, value=pack)  # the pack's Packing Spec column
        labels = {3: "Description", 4: "Size", 6: "Artwork", JOB_ROW: "Job Number"}
        for r, t in labels.items():
            ws.cell(row=r, column=len(ADDRESS), value=t).font = Font(bold=True)
        stores = spec["stores"]
        first_store = PACK_ROW + 1
        last_store = PACK_ROW + len(stores)
        # Codes as Packing Sheets works them out, on the original columns (before any code column went in)
        grid = [[None] * 60 for _ in range(last_store + 2)]
        for i, st in enumerate(stores):
            k = 0
            for pack in packs:
                for c in raw[pack]:
                    grid[first_store + i - 1][c - 1] = st[8][k] or None
                    k += 1
        computed[tab] = {}
        for pack in packs:
            _, _, _, _, ordered = generate_pack_signatures(grid, range(PACK_ROW + 1, last_store + 1), min(raw[pack]), max(raw[pack]))
            computed[tab][pack] = ordered
        # Write the columns left to right: [code col][sub-group code cols][items] per pack
        col = len(ADDRESS) + len(packs) + 1
        item_index = 0
        for pack in packs:
            start = col
            if codes:
                ws.cell(row=JOB_ROW, column=col, value=f"Code for Pack {pack}")
                for i in range(len(stores)):
                    ws.cell(row=first_store + i, column=col, value=computed[tab][pack][i] or None)
                col += 1
                for a, b in spec["subgroups"].get(pack, []):
                    cols = raw[pack][a - 1:b]
                    _, _, _, _, sub = generate_pack_signatures(grid, range(PACK_ROW + 1, last_store + 1), min(cols), max(cols))
                    ws.cell(row=JOB_ROW, column=col, value=f"Code for pack {pack}: {a}-{b}")
                    for i in range(len(stores)):
                        ws.cell(row=first_store + i, column=col, value=sub[i] or None)
                    col += 1
            for _ in raw[pack]:
                p, job, desc, size, colour = spec["items"][item_index]
                ws.cell(row=3, column=col, value=desc)
                ws.cell(row=4, column=col, value=size)
                ws.cell(row=JOB_ROW, column=col, value=job)
                img = XLImage(_art(job[-2:], colour))
                img.width, img.height = 60, 80
                ws.add_image(img, f"{get_column_letter(col)}6")
                for i, st in enumerate(stores):
                    q = st[8][item_index]
                    ws.cell(row=first_store + i, column=col, value=q or None)
                ws.column_dimensions[get_column_letter(col)].width = 14
                item_index += 1
                col += 1
            ws.cell(row=PACK_ROW, column=start, value=pack)
            if col - 1 > start:
                ws.merge_cells(start_row=PACK_ROW, start_column=start, end_row=PACK_ROW, end_column=col - 1)
        for i, st in enumerate(stores):
            for c, v in enumerate(st[:7], 1):
                ws.cell(row=first_store + i, column=c, value=v)
            for n, specv in enumerate(st[7]):
                ws.cell(row=first_store + i, column=len(ADDRESS) + 1 + n, value=specv)
        ws.cell(row=last_store + 1, column=1, value="TOTAL")
        ws.row_dimensions[6].height = 64
        for c in range(1, col):
            ws.cell(row=PACK_ROW, column=c).font = Font(bold=True)
            ws.cell(row=PACK_ROW, column=c).fill = PatternFill("solid", fgColor="DDDDDD")
            ws.cell(row=PACK_ROW, column=c).alignment = Alignment(horizontal="center")
        ws.column_dimensions["A"].width = 22
        ws.column_dimensions["C"].width = 30
    if change:
        tab, i, item, qty = change
        ws = wb[tab]
        cols = [c for c in range(1, ws.max_column + 1) if re.match(r'J\d', str(ws.cell(row=JOB_ROW, column=c).value or ''))]
        ws.cell(row=PACK_ROW + 1 + i, column=cols[item], value=qty)
    wb.save(path)
    return computed


def build_signature_links(path, computed, plugin=True, tamper=None):
    """tamper: (tab, pack, store index, code) written instead of the true code (files from different runs)."""
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        plugin_rows, barcode_rows = [], []
        for tab, spec in TABS.items():
            packs, raw = _layout(tab)
            data = {" Name": [s[0] for s in spec["stores"]]}
            for pack in packs:
                codes = list(computed[tab][pack])
                if tamper and tamper[0] == tab and tamper[1] == pack:
                    codes[tamper[2]] = tamper[3]
                data[f"Code for {pack}"] = codes
            pd.DataFrame(data).to_excel(writer, sheet_name=tab, index=False)
            info = {"pack_ranges": [{"name": p, "start": raw[p][0], "end": raw[p][-1]} for p in packs],
                    "store_col_idx": 1, "pack_group_row": PACK_ROW, "job_id_row": JOB_ROW,
                    "last_row": PACK_ROW + len(spec["stores"]), "stock_start_col": len(ADDRESS) + 1,
                    "last_stock_col": len(ADDRESS) + len(packs)}
            summaries = [{"name": p, "unique_sigs": sorted(set(c for c in computed[tab][p] if c)), "is_selected": True} for p in packs]
            plugin_rows += label_plugin_rows(tab, info, {"installer": spec.get("installer")}, summaries, f"Final {tab}.xlsx")
            if spec.get("installer"):
                for pack in packs:
                    for code in sorted(set(c for c in computed[tab][pack] if c)):
                        for (p, job, desc, size, colour) in spec["items"]:
                            if p == pack:
                                barcode_rows.append([tab, pack, code, job, "", 1, job])
        if barcode_rows:
            pd.DataFrame(barcode_rows, columns=DIVIDER_BARCODE_COLUMNS).to_excel(writer, sheet_name=DIVIDER_BARCODE_SHEET, index=False)
        if plugin:
            pd.DataFrame(plugin_rows, columns=LABEL_PLUGIN_COLUMNS).to_excel(writer, sheet_name=LABEL_PLUGIN_SHEET, index=False)


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "."
    os.makedirs(out, exist_ok=True)
    codes = build_final_file(os.path.join(out, "Final J476699 Spring rollout.xlsx"))
    build_signature_links(os.path.join(out, "Signature links_J476699.xlsx"), codes)
    print("written to", out)
