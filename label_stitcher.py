"""Stitch Labels: puts each courier label from the courier portal's label PDF onto its packing label.

How a courier label finds its packing label:
  1. Its text is read and the value after "Item Ref:" or "Item Reference:" (any capitals, with or without the
     colon; a label may carry both) is taken. In it, the unique code written by Generate ('01-QWXK-L-1-Provision Clayton' -> QWXK) is looked up in
     the project's label map. A code found elsewhere on the label in the same '-QWXK-L-' shape also counts.
  2. Only for projects generated before the codes existed: the Item Ref text itself ('10 Label 2 - Q Eyewear', the
     leading number dropped) is matched against the label map's item references. A project with codes matches by
     code only, so a label from another job can never land on the wrong packing label.

How it's placed: the courier label page is rendered as a lossless, high-resolution grayscale image (600 dpi), the
blank margin around the printed label is cropped off, the label is turned so its text reads upright (whatever way it
was printed on the page), a genuinely landscape label is then turned a quarter to fit the portrait courier region,
and it's scaled to fit the region (107 x 150 mm, top-left of the packing label's first page) without distortion.
Barcodes and QR codes stay sharp: at 600 dpi a 0.25 mm bar is 6 pixels wide.

Two steps: plan_stitch() reads the text of every courier label at once and works out where each goes (a fraction of
a second, nothing drawn); render_stitch() draws the Complete Labels PDF for that plan, once. A pack of several boxes
('2 x OB170170170') is one place per box (packing_slots()), filled in turn by the courier labels with its code.
"""
import io
import math
import os
import re
from collections import Counter

import pymupdf as fitz
from PIL import Image, ImageOps

import app_log
from packing_label_generator import _norm, draw_report_pages

log = app_log.get('stitch')
DPI = 600
MM = 72 / 25.4
PAD_MM = 1.0            # white kept around the cropped label
# 'Item Ref:' or 'Item Reference:' (the longer word is matched whole, so its 'erence' isn't taken as the value)
_REF = re.compile(r'\bitem\s*ref(?:erence)?\b\s*[:.\-]?', re.I)
_STOP = re.compile(r'item\s*desc|description|powered by', re.I)
_CODE_IN_REF = re.compile(r'\d+(?:\.\d+)?-([A-Z]{4,8})-L-')        # older Item References: '01-QWXK-L-1-Store'
_CODE_ANYWHERE = re.compile(r'(?<![A-Z])-?([A-Z]{4,8})-L-\d')
_CODE_TAB = re.compile(r'(?<![A-Za-z0-9])\d{1,4}T\d{1,2}([A-Z]{4,})')  # '01T1QWXK Store': count, T + tab, code


class StitchError(ValueError):
    """The project or the uploaded file can't be stitched."""


def _item_ref_texts(text):
    """The text after every 'Item Ref:' / 'Item Reference:' on the label (each up to 'Item Desc' or a few lines)."""
    found = []
    for m in _REF.finditer(text):
        tail = text[m.end():m.end() + 240]
        stop = _STOP.search(tail)
        nxt = _REF.search(tail)
        cut = min(x.start() for x in (stop, nxt) if x) if (stop or nxt) else len(tail)
        value = " ".join(tail[:cut].split())
        if value:
            found.append(value)
    return found


def _item_ref_text(text):
    """The first Item Ref / Item Reference value on the label, or ''."""
    refs = _item_ref_texts(text)
    return refs[0] if refs else ''


def _tab_code(text, known):
    """The code in a '01T1QWXK Store' Item Reference, if it's one of known: the serial with its code
    ('01T1QWXK'), or the code alone ('QWXK') for older projects. The capitals after the
    serial are tried from the longest down, so a store name the portal ran straight on ('01T1QWXKPROVISION')
    doesn't get in."""
    for m in _CODE_TAB.finditer(text):
        serial, run = text[m.start():m.start(1)], m.group(1)
        for n in range(len(run), 3, -1):
            for candidate in (serial + run[:n], run[:n]):
                if candidate in known:
                    return candidate
    return None


def find_code(text, known):
    """The unique code on a courier label, if it's one of this project's (known: set of codes). Item References
    of both layouts: '01T1QWXK Provision Clayton' and the older '01-QWXK-L-1-Provision Clayton'."""
    known = {k.upper() for k in known}
    for ref in _item_ref_texts(text) + [text]:  # in the Item Ref first, then anywhere on the label
        found = _tab_code(ref, known) or _tab_code(ref.replace(' ', ''), known)
        if found:
            return found
    in_refs = [c for ref in _item_ref_texts(text) for c in _CODE_IN_REF.findall(ref.replace(' ', ''))]
    for candidate in in_refs + _CODE_ANYWHERE.findall(text.replace(' ', '')):
        if candidate.upper() in known:
            return candidate.upper()
    return None


def _ref_key(text):
    """'10 Label 2 - Q Eyewear' -> 'label 2 q eyewear' (the row number the portal file adds is dropped)."""
    return _norm(re.sub(r'^\s*\d+(?:\.\d+)?\s+', '', text))


def label_targets(label_map):
    """{code: pack} and {item reference key: pack} for every packing label of the project that has a courier label."""
    by_code, by_ref = {}, {}
    for store, packs in label_map.get('stores', {}).items():
        for pk in packs:
            entry = {**pk, 'store': store}
            if pk.get('open360_code'):
                by_code[pk['open360_code'].upper()] = entry
            by_ref[_ref_key(pk['item_reference'])] = entry
    return by_code, by_ref


def text_angle(page):
    """Which way the label's text runs as the page is shown: 0 (upright), 90, 180 or 270 degrees clockwise.
    Decided by the text itself (most characters win), plus the page's own rotation flag."""
    votes = Counter()
    for block in page.get_text('dict')['blocks']:
        for line in block.get('lines', []):
            dx, dy = line['dir']
            angle = round(math.degrees(math.atan2(dy, dx)) / 90) % 4 * 90  # y points down, so this is clockwise
            votes[angle] += sum(len(s['text'].strip()) for s in line['spans'])
    return ((votes.most_common(1)[0][0] if votes else 0) + page.rotation) % 360


def render_label(page, dpi=DPI):
    """The courier label as a cropped, lossless grayscale PIL image, turned so its text reads upright
    (a 4x6 label printed sideways on a landscape page comes back the right way up)."""
    pix = page.get_pixmap(dpi=dpi, colorspace=fitz.csGRAY, alpha=False)
    img = Image.frombytes("L", (pix.width, pix.height), pix.samples)
    angle = text_angle(page)
    if angle:
        img = img.rotate(angle, expand=True, fillcolor=255)  # PIL turns anti-clockwise: undoes the clockwise angle
    box = ImageOps.invert(img).point(lambda v: 255 if v > 24 else 0).getbbox()  # ignore near-white noise
    if box:
        pad = round(PAD_MM / 25.4 * dpi)
        box = (max(0, box[0] - pad), max(0, box[1] - pad), min(img.width, box[2] + pad), min(img.height, box[3] + pad))
        img = img.crop(box)
    return img


def place_label(page, region, img):
    """Fits img (upright) into region (a fitz.Rect on the packing label page), aspect kept, centred across and
    top-aligned, covering the region's placeholder. A landscape label is turned a quarter (reading bottom to top) to
    suit the portrait region, so it stays near full size instead of shrinking to the region's width."""
    if (img.width > img.height) != (region.width > region.height):
        img = img.rotate(90, expand=True, fillcolor=255)
    scale = min(region.width / img.width, region.height / img.height)
    w, h = img.width * scale, img.height * scale
    target = fitz.Rect(region.x0 + (region.width - w) / 2, region.y0, region.x0 + (region.width + w) / 2, region.y0 + h)
    page.draw_rect(region, color=None, fill=(1, 1, 1), overlay=True)  # hide the dashed placeholder and its text
    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=True)
    page.insert_image(target, stream=buf.getvalue(), keep_proportion=False)
    return target, scale


# ---------- matching by hand: what's left after a stitch ----------

def page_key(page, text):
    """A courier label page's identity, the same when the portal's PDF is downloaded again: a hash of its text, or
    for a scanned page without text, of a small render of it. Matches made by hand are remembered by this key."""
    import hashlib
    words = " ".join(text.split()).lower()
    if words:
        return "t" + hashlib.sha1(words.encode('utf-8')).hexdigest()[:16]
    pix = page.get_pixmap(dpi=40, colorspace=fitz.csGRAY)
    return "i" + hashlib.sha1(pix.samples).hexdigest()[:16]


def address_snippet(text, lines=5):
    """The lines of a courier label that look like its delivery address: the ones up to the first 4-digit postcode."""
    rows = [" ".join(r.split()) for r in text.splitlines() if r.strip()]
    for i, r in enumerate(rows):
        if re.search(r'\b\d{4}\b', r) and not re.search(r'\d{5,}', r):
            return rows[max(0, i - lines + 1):i + 1]
    return rows[:lines]


def slot_name(item_reference, box, boxes):
    """The id of one place a courier label goes: the item reference, plus the box for a pack of several boxes
    ('Label 1 - Provision Clayton · box 2 of 2')."""
    return item_reference if boxes <= 1 else f"{item_reference} · box {box} of {boxes}"


def packing_slots(label_map):
    """Every place a courier label goes, in packing label order: one per box of every packing label ('2 x
    OB170170170' is two boxes, each with its own pages and courier label; both share one Open360 row and code).
    [{'slot', 'item_reference', 'box', 'boxes', 'page', 'pack', 'not_in_csv'}]; page is 1-based in the packing
    labels PDF (None when the label map doesn't say)."""
    out = []
    groups = [(store, pk, False) for store, packs in label_map.get('stores', {}).items() for pk in packs]
    groups += [(pk.get('store', ''), pk, True) for pk in label_map.get('not_in_courier_csv') or []]
    for store, pk, left_out in groups:
        boxes = max(1, int(pk.get('boxes') or 1))
        pages = pk.get('box_pages') or [pk.get('courier_label_page') or (pk.get('pdf_pages') or [None])[0]]
        for box in range(1, boxes + 1):
            out.append({'slot': slot_name(pk['item_reference'], box, boxes), 'item_reference': pk['item_reference'],
                        'box': box, 'boxes': boxes, 'page': pages[box - 1] if box - 1 < len(pages) else None,
                        'pack': {**pk, 'store': store}, 'not_in_csv': left_out})
    out.sort(key=lambda sl: (sl['page'] or 10 ** 9))
    return out


def waiting_packs(label_map, done):
    """The packing labels (each box of them) still without a courier label, with what's needed to recognise them:
    receiver, Attn and the full address of their consignment. done: the slots already taken. Packs left out of the
    courier CSV are included (they may have been booked another way), flagged. 'item_reference' is the slot id."""
    cons = {c['number']: c for c in label_map.get('consignments', [])}
    out = []
    for sl in packing_slots(label_map):
        if sl['slot'] in done:
            continue
        pk = sl['pack']
        common = {'item_reference': sl['slot'], 'pack_reference': sl['item_reference'], 'box': sl['box'],
                  'boxes': sl['boxes'], 'store': pk.get('store', ''), 'label': pk.get('label'), 'of': pk.get('of'),
                  'packing_spec': pk.get('packing_spec', ''), 'page': sl['page'],
                  'serial': (re.match(r'\s*(\d{1,4}T\d{1,2})[A-Z]', pk.get('open360_item_reference') or '') or [None, None])[1]}
        if sl['not_in_csv']:
            out.append({**common, 'receiver': pk.get('store', ''), 'contact': '',
                        'address': 'No complete address (left out of the courier CSV)', 'postcode': '', 'suburb': '',
                        'line1': '', 'line2': '', 'installer': False, 'consignment': None, 'not_in_csv': True})
            continue
        d = (cons.get(pk.get('consignment')) or {}).get('destination') or {}
        out.append({**common, 'receiver': d.get('receiver') or pk['headed_to']['receiver'],
                    'contact': d.get('contact', ''), 'address': pk['headed_to']['address'],
                    'postcode': d.get('postcode', ''), 'suburb': d.get('suburb', ''),
                    'line1': d.get('line1', ''), 'line2': d.get('line2', ''),
                    'installer': bool(pk['headed_to'].get('installer')), 'consignment': pk.get('consignment'),
                    'not_in_csv': False})
    return out


def address_score(text, pack):
    """0..1: how surely a courier label's text is addressed to this packing label's receiver. Postcode, suburb,
    street numbers and receiver name each count (the courier prints them; Attn and spacing vary)."""
    words = set(_norm(text).split())
    if not pack.get('postcode'):
        return 0.0
    score = 0.4 if pack['postcode'] in words else 0.0
    suburb = _norm(pack.get('suburb')).split()
    if suburb and all(w in words for w in suburb):
        score += 0.2
    numbers = [w for w in _norm(f"{pack.get('line1', '')} {pack.get('line2', '')}").split() if any(ch.isdigit() for ch in w)]
    if numbers and all(n in words for n in numbers):
        score += 0.2
    name = [w for w in _norm(pack.get('receiver')).split() if len(w) > 1]
    if name:
        score += 0.2 * sum(w in words for w in name) / len(name)
    return round(score, 2)


_OPEN360_CODE = re.compile(r'^\s*\d+(?:\s*-\s*([A-Z]{4,})\s*-\s*L\s*-|T\d{1,2}([A-Z]{4,}))')


def codes_elsewhere(project_dir):
    """{Open360 code: project name} for the other projects next to this one, to explain a courier label whose code
    isn't this project's: it was booked from another project's Open360 CSV (every Generate makes new codes)."""
    import json
    found = {}
    parent = os.path.dirname(os.path.abspath(project_dir))
    own = os.path.abspath(project_dir)
    try:
        folders = [os.path.join(parent, d) for d in os.listdir(parent)]
    except OSError:
        return found
    for folder in folders:
        if os.path.abspath(folder) == own or not os.path.isdir(folder):
            continue
        for f in os.listdir(folder):
            if not f.endswith('.labelmap.json'):
                continue
            try:
                with open(os.path.join(folder, f), encoding='utf-8') as fh:
                    lm = json.load(fh)
            except (OSError, ValueError):
                continue
            name = (lm.get('project') or {}).get('name') or re.sub(r'_\d{6}_\d{4}$', '', os.path.basename(folder))
            for packs in (lm.get('stores') or {}).values():
                for pk in packs:
                    if pk.get('open360_code'):
                        code = pk['open360_code'].upper()
                        found.setdefault(code, name)
                        job = re.match(r'^\d+T\d+([A-Z]{4,})$', code)  # '01T1QWXK': QWXK (one code per job in older projects)
                        if job:
                            found.setdefault(job.group(1), name)
    return found


def ref_parts(ref):
    """What a courier label's Item Ref says about its packing label, whichever job it's from:
    '42T1LCQK Blue Star Moorabbin' -> serial 42T1, store 'Blue Star Moorabbin';
    '01-QWXK-L-1-Provision Clayton' -> label 1, store; '10 Label 2 - Q Eyewear' -> label 2, store."""
    ref = " ".join(str(ref or '').split())
    m = re.match(r'^(\d{1,4}T\d{1,2})\s*[A-Z]{4,}\s+(.+)$', ref)
    if m:  # 'tail': all after the serial, for a store name the portal may have run into the code
        return {'serial': m.group(1), 'label': None, 'store': m.group(2), 'tail': ref[m.end(1):]}
    m = re.match(r'^(\d{1,4}T\d{1,2})\s*([A-Z]{4,}.*)$', ref)  # the store run straight on: the code is in front of it
    if m:
        return {'serial': m.group(1), 'label': None, 'store': m.group(2), 'tail': m.group(2)}
    m = re.match(r'^\d+(?:\.\d+)?-[A-Z]{4,}-L-(\d+)-(.+)$', ref)
    if m:
        return {'serial': None, 'label': int(m.group(1)), 'store': m.group(2)}
    m = re.match(r'^(?:\d+(?:\.\d+)?\s+)?Label\s+(\d+)\s*-\s*(.+)$', ref, re.I)
    if m:
        return {'serial': None, 'label': int(m.group(1)), 'store': m.group(2)}
    return {'serial': None, 'label': None, 'store': ''}


def _store_score(ref, store):
    """0..1: how surely an Item Ref's store is this packing label's store (spacing, capitals and punctuation aside;
    a store name the portal ran on after the code is found inside it)."""
    from difflib import SequenceMatcher
    a, b = _norm(ref.get('store')).replace(' ', ''), _norm(store).replace(' ', '')
    if not a or not b:
        return None
    tail = _norm(ref.get('tail')).replace(' ', '')
    # a store run into the code: the tail is the code (4-6 letters) then the store, nothing else in front
    if a == b or (tail.endswith(b) and len(tail) - len(b) <= 6):
        return 1.0
    return round(SequenceMatcher(None, a, b).ratio(), 2)


def pair_score(page, w):
    """0..1: how surely a courier label belongs to this waiting packing label. The address printed on the label
    (address_score) and, when the label has an Item Ref, the store in it: several stores' packs can go to one
    address (an installer), and the Item Ref names the store each label was booked for. Its serial ('42T1') or label
    number ('L-1', 'Label 1') adds a little, to tell one store's labels apart."""
    address = address_score(page['text'], w)
    ref = page.get('ref') or {}
    store = _store_score(ref, w.get('store'))
    score = address if store is None else 0.55 * address + 0.45 * store
    if ref.get('serial') and ref['serial'] == w.get('serial'):
        score += 0.05
    if ref.get('label') and ref['label'] == w.get('label'):
        score += 0.05
    return round(min(score, 1.0), 2)


def suggest_pairs(pages, waiting, keep=8):
    """A suggested packing label for each unmatched courier page ([{key, text, item_ref}]), in file order.
    Packing labels of one consignment (same address) and one store count as one group: its courier labels are given
    its packing labels in order (the one whose serial or label number the Item Ref names first, then Label 1, then
    Label 2). A group is only suggested when it's clearly the best one: by the address on the label, and by the store
    in its Item Ref when there's one (pair_score), so one installer's packs for different stores aren't mixed up.
    Each page also gets 'candidates': up to `keep` packing labels [item reference, score], best first, for the
    matching page's "closest match" next to it."""
    groups = {}
    for w in waiting:
        if w['consignment'] is not None:
            groups.setdefault((w['consignment'], _norm(w.get('store'))), []).append(w)
    free = {g: list(ws) for g, ws in groups.items()}
    for p in pages:
        p['ref'] = ref_parts(p.get('item_ref') or _item_ref_text(p.get('text', '')))
        scored = {g: max(pair_score(p, w) for w in ws) for g, ws in groups.items()}
        ranked = sorted(((sc, g) for g, sc in scored.items()), key=lambda x: -x[0])
        p['candidates'] = sorted(([w['item_reference'], pair_score(p, w)] for ws in groups.values() for w in ws),
                                 key=lambda c: -c[1])
        p['candidates'] = [c for c in p['candidates'] if c[1] > 0][:keep]
        p['suggestion'], p['score'] = None, 0
        if not ranked or ranked[0][0] < 0.6 or (len(ranked) > 1 and ranked[0][0] - ranked[1][0] < 0.12):
            p.pop('ref', None)
            continue
        best = ranked[0][1]
        if free[best]:
            # the packing label the Item Ref names (serial / label number) first, else the next in label order
            named = [w for w in free[best] if pair_score(p, w) > ranked[0][0] - 0.001 and (
                (p['ref'].get('serial') and w.get('serial') == p['ref']['serial']) or (p['ref'].get('label') and w.get('label') == p['ref']['label']))]
            pick = named[0] if named else free[best][0]
            free[best].remove(pick)
            p['suggestion'], p['score'] = pick['item_reference'], pair_score(p, pick)
        p.pop('ref', None)
    return pages


QUADRANTS = ("top-left", "top-right", "bottom-left", "bottom-right")


def four_up_order(count):
    """Cut-and-stack order for count labels 4-up: [(page, quadrant)] per label, in file order. With P pages, the
    top-left quarters hold labels 1..P, top-right P+1..2P, bottom-left, then bottom-right: cut the printed stack into
    quarters, put the top-left pile on the top-right pile and so on, and the labels are back in file order."""
    pages = math.ceil(count / 4)
    return [(i % pages, i // pages) for i in range(count)]


def add_four_up_pages(doc, images, page_rect, margin_pt):
    """Appends the images 4-up on pages the size of the packing labels, with the same margin round each quarter and
    dashed cut guides on the centre lines. Returns [(page index in doc, quadrant name)] per image."""
    if not images:
        return []
    order = four_up_order(len(images))
    first = len(doc)
    W, H = page_rect.width, page_rect.height
    mx, my = W / 2, H / 2
    quads = [fitz.Rect(margin_pt, margin_pt, mx - margin_pt, my - margin_pt),
             fitz.Rect(mx + margin_pt, margin_pt, W - margin_pt, my - margin_pt),
             fitz.Rect(margin_pt, my + margin_pt, mx - margin_pt, H - margin_pt),
             fitz.Rect(mx + margin_pt, my + margin_pt, W - margin_pt, H - margin_pt)]
    for _ in range(max(p for p, _ in order) + 1):
        page = doc.new_page(width=W, height=H)
        for a, b in (((mx, 0), (mx, H)), ((0, my), (W, my))):
            page.draw_line(a, b, color=(0.65, 0.65, 0.65), width=0.5, dashes="[4 3] 0")
    for img, (p, q) in zip(images, order):
        place_label(doc[first + p], quads[q], img)
    return [(first + p, QUADRANTS[q]) for p, q in order]


def _span(pages):
    """'5' or '5-12' for a run of 1-based page numbers."""
    return '' if not pages else str(pages[0]) if len(pages) == 1 else f"{pages[0]}-{pages[-1]}"


def _pages(pages):
    """'Page 5' or 'Pages 5-12'."""
    return ("Page " if len(pages) == 1 else "Pages ") + _span(pages)


def _compact(section):
    """A section Generate saved in a project's label map, in the short form: projects made before the compact report
    listed every installer pack ('Wilson Storage: Label 1 - A, Label 2 - B'); now it's 'Wilson Storage x2'."""
    if not section['title'].startswith("Going to installers"):
        return section
    items = []
    for item in section.get('items', []):
        who, sep, refs = item.partition(': Label ')
        items.append(f"{who} x{len(re.findall(r'(?:^|, )Label ' + chr(92) + 'd+', 'Label ' + refs))}" if sep else item)
    title = re.sub(r', labelled for the installer, not the store$', ' labelled for the installer', section['title'])
    return {**section, 'title': title, 'items': items}


def _region(label_map):
    region_mm = label_map.get('courier_label_region_mm') or {'x': 4, 'y': 4, 'width': 107, 'height': 150}
    rect = fitz.Rect(region_mm['x'] * MM, region_mm['y'] * MM,
                     (region_mm['x'] + region_mm['width']) * MM, (region_mm['y'] + region_mm['height']) * MM)
    return rect, region_mm


def _packing_pdf(project_dir, label_map):
    pdf_name = (label_map.get('files') or {}).get('packing_labels_pdf')
    if not pdf_name or not os.path.isfile(os.path.join(project_dir, pdf_name)):
        raise StitchError("This project's packing labels PDF is missing.")
    return os.path.join(project_dir, pdf_name)


def plan_stitch(project_dir, label_map, courier_pdfs, manual=None):
    """Which courier label goes on which packing label, from the text of every courier label page at once: nothing
    is drawn or written, so it takes a moment however many labels there are.
      1. The unique code in each label's Item Ref (or, for a project generated before codes, the Item Ref text)
         picks the pack; a pack of several boxes takes its labels in turn, box 1 first.
      2. For the rest, the matches made by hand (manual: {courier page key: slot or item reference}).
    Returns {'placed', 'unmatched' (with each page's text), 'duplicates', 'without_label', 'slots', 'files',
    'complete'}: complete when every courier label found its packing label and every packing label that should get
    a courier label has one (so the Complete Labels can be made straight away)."""
    packing_path = _packing_pdf(project_dir, label_map)
    by_code, by_ref = label_targets(label_map)
    if not by_code and not by_ref:
        raise StitchError("This project's label map lists no packing labels.")
    with fitz.open(packing_path) as packing:
        packing_pages = len(packing)
    manual = manual or {}
    slots = packing_slots(label_map)
    slots_of, slot_by_id = {}, {}
    for sl in slots:
        slots_of.setdefault(sl['item_reference'], []).append(sl)
        slot_by_id[sl['slot']] = sl
    pages, unmatched, duplicates, placed, done, by_hand = [], [], [], [], {}, []
    seen_keys = Counter()
    elsewhere = None  # other projects' codes, read the first time a label's code isn't this project's
    for name, data in courier_pdfs:
        try:
            src = fitz.open(stream=data, filetype="pdf")
        except Exception as e:
            log.warning("%s: not a PDF that can be read (%s)", name, e)
            unmatched.append({'file': name, 'page': None, 'reason': "not a PDF that can be read", 'why': "not a PDF that can be read"})
            continue
        log.info("  %s: %d page(s)", name, len(src))
        for i, page in enumerate(src):
            text = page.get_text()
            base_key = page_key(page, text)
            seen_keys[base_key] += 1
            key = base_key if seen_keys[base_key] == 1 else f"{base_key}-{seen_keys[base_key]}"  # identical pages
            code = find_code(text, by_code.keys())
            pack, method = (by_code[code], 'code') if code else (None, None)
            if pack is None and not by_code:  # older project without codes: match the Item Ref text
                pack = next((by_ref[_ref_key(r)] for r in _item_ref_texts(text) if _ref_key(r) in by_ref), None)
                method = 'item reference' if pack else None
            pages.append({'file': name, 'page': i + 1, 'key': key, 'text': text, 'code': code,
                          'pack': pack['item_reference'] if pack else None, 'store': pack.get('store', '') if pack else '',
                          'method': method})
        src.close()

    def take(p, sl, method):
        if not sl['page'] or sl['page'] > packing_pages:
            log.warning("    %s p%d: %s matched (%s) but its packing label page %s is missing", p['file'], p['page'],
                        sl['slot'], method, sl['page'])
            unmatched.append({**_why_not(p, project_dir, by_code, elsewhere), 'reason': f"{sl['slot']}: its packing label page is missing",
                              'why': "its packing label page is missing"})
            return
        done[sl['slot']] = f"{p['file']} p{p['page']}"
        placed.append({'file': p['file'], 'page': p['page'], 'key': p['key'], 'item_reference': sl['slot'],
                       'pack_reference': sl['item_reference'], 'box': sl['box'], 'boxes': sl['boxes'],
                       'store': sl['pack'].get('store', ''), 'packing_page': sl['page'], 'method': method,
                       'code': p['code'] if method == 'code' else None})
        log.debug("    %s p%d: %s (matched %s%s) -> packing label page %d", p['file'], p['page'], sl['slot'], method,
                  f" {p['code']}" if method == 'code' else '', sl['page'])

    # 1. Codes (or Item Ref text): each label takes its pack's next box still free
    rest = []
    for p in pages:
        if not p['pack']:
            rest.append(p)
            continue
        free = next((sl for sl in slots_of.get(p['pack'], []) if sl['slot'] not in done), None)
        if free:
            take(p, free, p['method'])
        else:
            first = done.get(slots_of[p['pack']][0]['slot'], '')
            log.debug("    %s p%d: %s again (all its boxes have a label) - skipped", p['file'], p['page'], p['pack'])
            duplicates.append({'file': p['file'], 'page': p['page'], 'item_reference': p['pack'], 'first': first})
    # 2. Matches made by hand, for the labels the codes didn't place
    for p in rest:
        wanted = manual.get(p['key'])
        sl = slot_by_id.get(wanted) if wanted else None
        if wanted and sl is None:  # an older match by item reference: that pack's next free box
            sl = next((s for s in slots_of.get(wanted, []) if s['slot'] not in done), None)
        if sl is not None and sl['slot'] not in done:
            take(p, sl, 'by hand')
            if sl['slot'] in done:
                by_hand.append({**_why_not(p, project_dir, by_code, elsewhere), 'matched_to': sl['slot']})
            continue
        if elsewhere is None and by_code and _item_ref_texts(p['text']):
            elsewhere = codes_elsewhere(project_dir)
        unmatched.append(_why_not(p, project_dir, by_code, elsewhere))

    left_out = {sl['slot'] for sl in slots if sl['not_in_csv']}
    missing = [sl['slot'] for sl in slots if sl['slot'] not in done]
    without_label = [s + (" (not in the courier CSV)" if s in left_out else '') for s in missing]
    complete = not unmatched and all(s in left_out for s in missing)
    log.info("  plan: %d placed, %d courier label(s) unmatched, %d packing label(s) without one, %d duplicate(s)%s",
             len(placed), len(unmatched), len(missing), len(duplicates), " - complete" if complete else '')
    return {'placed': placed, 'unmatched': unmatched, 'duplicates': duplicates, 'without_label': without_label,
            'missing': missing, 'slots': slots, 'done': done, 'by_hand': by_hand,
            'files': [n for n, _ in courier_pdfs], 'complete': complete}


def _why_not(p, project_dir, by_code, elsewhere):
    """An unmatched courier label page, with why it didn't match (a code of another project is named)."""
    text = p['text']
    refs = _item_ref_texts(text)
    code_seen = _OPEN360_CODE.match(refs[-1]) if refs and by_code else None
    other = None
    if code_seen and elsewhere:
        run = (code_seen.group(1) or code_seen.group(2)).upper()
        # the old layout's code stands alone; in '01T1QWXK Store' the capitals may run into the store name
        other = elsewhere.get(run) or next((elsewhere[run[:n]] for n in range(len(run), 3, -1) if run[:n] in elsewhere), None)
    reason = ("no text on the page (a scanned image?)" if not text.strip() else
              "no Item Ref / Item Reference on the label" if not refs else
              f"Item Ref '{refs[-1][:60]}' is project {other}'s - booked from that project's Open360 CSV, "
              f"not this one's" if other else
              f"Item Ref '{refs[-1][:60]}' has no code of this project" if by_code else
              f"Item Ref '{refs[0][:60]}' isn't a label of this project")
    why = ("no text on the page (a scanned image?)" if not text.strip() else
           "no Item Ref on the label" if not refs else
           f"code of project {other} (booked from its Open360 CSV)" if other else
           "code isn't this project's" if by_code else "not a label of this project")
    log.debug("    %s p%d: not placed - %s", p['file'], p['page'], reason)
    return {'file': p['file'], 'page': p['page'], 'reason': reason, 'key': p['key'], 'text': text,
            'item_ref': refs[-1][:80] if refs else '', 'why': why, 'other_project': other}


def match_session(label_map, plan):
    """What the matching page needs from a plan: every courier page the codes didn't place (with suggested packing
    labels), those already matched by hand carrying 'matched_to' (so they show as pairs and can be changed), and the
    packing labels (each box) without a courier label from a code."""
    hand_slots = {u['matched_to'] for u in plan.get('by_hand', [])}
    waiting = waiting_packs(label_map, {s: v for s, v in plan['done'].items() if s not in hand_slots})
    order = {name: n for n, name in enumerate(plan['files'])}
    pages = sorted([u for u in plan['unmatched'] if u.get('key')] + plan.get('by_hand', []),
                   key=lambda u: (order.get(u['file'], 0), u['page'] or 0))
    left = suggest_pairs([{'key': u['key'], 'file': u['file'], 'page': u['page'], 'reason': u['reason'],
                           'why': u.get('why', u['reason']), 'item_ref': u.get('item_ref', ''),
                           'other_project': u.get('other_project'), 'matched_to': u.get('matched_to'),
                           'snippet': address_snippet(u['text']), 'text': u['text']}
                          for u in pages], waiting)
    for u in left:
        u.pop('text')
    return {'files': plan['files'], 'unmatched': left, 'waiting': waiting, 'placed': len(plan['placed']),
            'duplicates': len(plan['duplicates']), 'complete': plan['complete'],
            'expected': sum(1 for sl in plan['slots'] if not sl['not_in_csv'])}


def render_stitch(project_dir, label_map, courier_pdfs, plan, output_path, add_unmatched=True):
    """Draws the Complete Labels PDF for a plan (plan_stitch), in this order:
      1. the job report (stitching results, then what Generate found), replacing Generate's report page
      2. every packing label in the packing labels' own order, with its courier label where one was found
      3. courier labels that matched no packing label (no Item Ref / Item Reference, or a code of another job),
         4-up in cut-and-stack order (four_up_order), if add_unmatched
    Each courier label is rendered at 600 dpi only here. Returns {'placed', 'unmatched', 'duplicates',
    'without_label', 'report_pages', 'groups'}."""
    from datetime import datetime
    region, region_mm = _region(label_map)
    out = fitz.open(_packing_pdf(project_dir, label_map))
    sources = {}
    for name, data in courier_pdfs:
        try:
            sources[name] = fitz.open(stream=data, filetype="pdf")
        except Exception:
            pass
    project = label_map.get('project') or {}
    old_report = min(project.get('report_pages') or 0, len(out) - 1)
    placed = [dict(p) for p in plan['placed']]
    unmatched = [dict(u) for u in plan['unmatched']]
    duplicates, without_label = plan['duplicates'], plan['without_label']
    for p in placed:
        _, scale = place_label(out[p['packing_page'] - 1], region, render_label(sources[p['file']][p['page'] - 1]))
        p['scale_pct'] = round(scale * DPI / 72 * 100, 1)
    leftovers = [(u, render_label(sources[u['file']][u['page'] - 1]) if add_unmatched else None)
                 for u in unmatched if u.get('page') and u['file'] in sources]
    for doc in sources.values():
        doc.close()

    # The packing labels stay in their own order; only Generate's report page goes (the new report replaces it)
    order = list(range(old_report, len(out)))
    page_rect = out[order[0]].rect if order else fitz.paper_rect("a4-l")
    out.select(order)
    new_index = {old: n for n, old in enumerate(order)}

    # The courier labels that matched nothing: 4-up at the end, in file order for cut and stack
    margin = (region_mm.get('x') or 4) * MM
    spots = add_four_up_pages(out, [img for _, img in leftovers], page_rect, margin) if add_unmatched else []
    four_up_pages = sorted({p for p, _ in spots})
    slots = plan['slots']
    files = {name: 0 for name, _ in courier_pdfs}  # every PDF read, even one that placed nothing
    for p in placed:
        files[p['file']] = files.get(p['file'], 0) + 1

    def build(report_pages):
        at = lambda idx: idx + 1 + report_pages  # 0-based page in the reordered labels -> 1-based page in the file
        for (entry, _), (page_idx, quadrant) in zip(leftovers, spots):
            entry['four_up'] = f"page {at(page_idx)}, {quadrant}"
        groups = [
            f"{_pages([at(n) for n in range(len(order))])}: the packing labels in their own order - {len(placed)} with their "
            f"courier label, {len(slots) - len(placed)} without" if order else '',
            (f"{_pages([at(p) for p in four_up_pages])}: courier labels that matched no packing label, 4-up "
             f"({len(leftovers)}) - cut the stack into quarters, then stack top-left, top-right, bottom-left, bottom-right "
             f"to have them in file order" if add_unmatched else
             f"{len(leftovers)} courier label{'s' if len(leftovers) != 1 else ''} that matched no packing label: left out of "
             f"this file, as chosen") if leftovers else '',
        ]
        booked_elsewhere = Counter(u['other_project'] for u in unmatched if u.get('other_project'))
        boxed = sorted({(p['pack_reference'], p['boxes']) for p in placed if p.get('boxes', 1) > 1})
        sections = [
            {'title': "Courier labels booked from another project's Open360 CSV - their codes aren't this project's", 'level': 'bad',
             'items': [f"{n} label{'s' if n != 1 else ''} of project {name}: book this job from this project's own "
                       f"Open360 CSV (every Generate makes new codes), or match them by hand" for name, n in booked_elsewhere.items()]},
            {'title': "Packing labels without a courier label", 'level': 'bad',
             'items': without_label},
            {'title': ("Courier labels matching no packing label - 4-up at the end; cut into quarters, stack TL, TR, BL, BR"
                       if add_unmatched else "Courier labels matching no packing label - LEFT OUT of this file"), 'level': 'bad',
             'items': [f"{u['file']}{' p' + str(u['page']) if u['page'] else ''}: {u['reason']}"
                       + (f" ({u['four_up']})" if u.get('four_up') else '') for u in unmatched]},
            {'title': "Same courier label more than its pack has boxes - the first ones were used", 'level': 'warn',
             'items': [f"{d['file']} p{d['page']}: {d['item_reference']} (used {d['first']})" for d in duplicates]},
            {'title': "Packs of several boxes - one packing label and one courier label per box", 'level': 'info',
             'items': [f"{ref}: {n} boxes" for ref, n in boxed]},
        ] + [_compact(s) for s in (project.get('report') or {}).get('sections', []) if s.get('kind') != 'left_out'
             and not s['title'].startswith("Not in the courier CSV")]  # already in "without a courier label"
        missing = len(without_label)
        facts = [("Courier labels placed", f"{len(placed)} of {len(slots)} packing labels"),
                 ("Without a courier label", missing, 'bad' if missing else None)]
        if unmatched:
            facts.append(("Courier labels not matched", f"{len(unmatched)} ({'4-up at the end' if add_unmatched else 'left out'})", 'bad'))
        if duplicates:
            facts.append(("Duplicates", len(duplicates), 'warn'))
        facts.append(("Labels", _pages([at(n) for n in range(len(order))]).lower() if order else '-'))
        if add_unmatched and four_up_pages:
            facts.append(("4-up", _pages([at(p) for p in four_up_pages]).lower()))
        facts += [(name, f"{n} label{'s' if n != 1 else ''}") for name, n in files.items()]
        return {
            'title': "Complete Labels", 'job': label_map.get('consignment_reference'),
            'project': project.get('name') or re.sub(r'_\d{6}_\d{4}$', '', os.path.basename(project_dir)),
            'when': "Stitched " + datetime.now().strftime("%d/%m/%Y %H:%M"),
            'facts': facts, 'sections': sections}, groups

    # The report's own length moves every page number it quotes: settle it before drawing it for real
    report_pages = 1
    for _ in range(3):
        trial = fitz.open()
        needed = draw_report_pages(trial, build(report_pages)[0], at=0)
        trial.close()
        if needed == report_pages:
            break
        report_pages = needed
    report, groups = build(report_pages)
    report_pages = draw_report_pages(out, report, at=0)
    for p in placed:
        p['packing_page'] = new_index[p['packing_page'] - 1] + 1 + report_pages
    out.save(output_path, garbage=3, deflate=True)
    out.close()
    for u in unmatched:
        u.pop('text', None)
    return {'placed': placed, 'unmatched': unmatched, 'duplicates': duplicates, 'without_label': without_label,
            'report_pages': report_pages, 'groups': [g for g in groups if g]}


def stitch(project_dir, label_map, courier_pdfs, output_path, manual=None, add_unmatched=True):
    """plan_stitch, then render_stitch: the whole stitch in one go. Returns render_stitch's result plus 'session'
    (match_session: what the matching page needs) and 'complete'."""
    plan = plan_stitch(project_dir, label_map, courier_pdfs, manual)
    session = match_session(label_map, plan)
    result = render_stitch(project_dir, label_map, courier_pdfs, plan, output_path, add_unmatched)
    session['output'] = os.path.basename(output_path)
    return {**result, 'session': session, 'complete': plan['complete']}
