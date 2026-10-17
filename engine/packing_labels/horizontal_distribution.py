"""Packing Labels from a HORIZONTAL distribution file (the Packing Sheets layout).

Vertical files (parse_packing_data) have one row per item with the Packing Spec merged down a box's rows.
Horizontal files have one row per store and one column per item, the items of a pack under a merged pack header:

            Description | Code for Pack 1  A1 poster   A2 poster |  Code for Pack 2  Code for pack 2: 1-2  Card  Strip
  Job row   Job Number  | Code for Pack 1  J476699-01  J476699-02 | Code for Pack 2  Code for pack 2: 1-2  J…-05 J…-06
  Pack row  Store  Addr…|  <--------------- 1 ------------------> |  <------------------- 2 ---------------------->
            Corio  …    |  B               2           2          |  A                A                     3     10

The packing-sheet code columns ("Code for Pack <name>", sub-groups "Code for pack <name>: 1-5") are headers like any
job number: their cells hold each store's code instead of a quantity. A pack may have one, several or none:
  - in the file:            used, and checked against the quantities and the Signature Links file
  - only in Signature Links: taken from there (its "Code for <pack>" columns)
  - nowhere:                worked out exactly as Packing Sheets does (core_math.generate_pack_signatures)

The Signature Links file can carry a "Label Plugin" tab (written by Packing Sheets, core_math.label_plugin_rows):
the rows and columns Packing Sheets was given for each tab, so the Packing Labels user needs no mapping.

read_horizontal() returns pack groups shaped like parse_packing_data()'s, ordered like the packing sheets
(pack, then code A, B … Z, AA, then sub-group codes, then the store's row), each with its serial
'<store's number in the tab>T<tab>P<pack>' ('03T1P2'), the start of its Item Reference '03T1P2 Chadstone'.
Positions are found by pattern and header text, never fixed rows or columns.
"""
import io
import re
from collections import Counter, defaultdict

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import column_index_from_string

from engine.packing_labels.packing_label_generator import (extract_rich_value_images, _floating_image_cell, _issue, _row_ranges, _norm,
                                                            with_stores,
                                     is_install_flag, PackCheckError, ADDRESS_FIELDS, _header_field)
from engine.courier.courier_import import (detect_import_columns, detect_pack_columns, apply_choices, _merged_spans, _header_cells,
                            NO_PACK, TOTAL_ROW)
from engine.courier.courier_export import attn_text
from engine.core.core_math import (generate_pack_signatures, sanitize_cell, clean_store_name, LABEL_PLUGIN_SHEET,
                       DIVIDER_BARCODE_SHEET, read_divider_barcodes)
from engine.label_shuffler.pdf_engine import code128_modules
from engine.courier.packing_specs import split_count

JOB = re.compile(r'^[A-Z]{1,3}\d{4,8}(-\d{1,4})?[A-Z]?$', re.I)
CODE_HEADER = re.compile(r'^\s*code\s+for\s+(.+?)\s*$', re.I)
SUB_RANGE = re.compile(r'^(.*?):\s*(\d+\s*-\s*\d+)$')
ROW_LABELS = (  # label text left of the items, above the header -> item field
    ('desc', r'desc|item\s*name|^item$|^name$'),
    ('dimension', r'size|dimension|dims'),
    ('material', r'material|substrate'),
    ('notes', r'note|comment'),
    ('thumbnail', r'art\s*work|image|thumb|visual|picture'),
    ('pack_spec', r'pack(ing)?\s*spec|carton|box\s*type'),
)


# ---------------------------------------------------------------- codes and names

def code_key(code):
    """Packing-sheet order of a code: A … Z, AA … (length, then letters); no code last."""
    code = (code or '').strip().upper()
    return (len(code), code) if code else (99, '')


def parse_code_header(text):
    """'Code for Pack 2' -> ('Pack 2', None); 'Code for pack 2: 1-5' -> ('pack 2', '1-5'); anything else None."""
    m = CODE_HEADER.match(str(text or ''))
    if not m:
        return None
    rest = m.group(1)
    sub = SUB_RANGE.match(rest)
    return (sub.group(1).strip(), re.sub(r'\s+', '', sub.group(2))) if sub else (rest, None)


def names_pack(header_name, pack_name):
    """A code header's pack name against a pack's: Packing Sheets writes 'Code for Pack 1' in the Final file and
    'Code for 1' in Signature Links, so 'Pack 1' and '1' both name pack '1' (and 'Pack Pack A' names 'Pack A')."""
    h, p = _norm(header_name), _norm(pack_name)
    return h == p or h == _norm(f"pack {pack_name}")


# ---------------------------------------------------------------- the Signature Links file

def read_signature_file(path):
    """{'plugin': {tab: {...}}, 'codes': {tab: {pack header name: {store key: code}}}, 'barcodes': {...}} from a
    Signature Links file. Every part may be empty: an older file has no Label Plugin tab."""
    out = {'plugin': {}, 'codes': {}, 'barcodes': {}, 'file': path}
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            rows = [[c for c in r] for r in ws.iter_rows(values_only=True)]
            if not rows:
                continue
            header = [str(h).strip() if h is not None else '' for h in rows[0]]
            if ws.title == LABEL_PLUGIN_SHEET:
                out['plugin'] = _plugin_from_rows(header, rows[1:])
            elif ws.title != DIVIDER_BARCODE_SHEET:
                tab = {}
                for c, h in enumerate(header[1:], start=1):
                    parsed = parse_code_header(h)
                    if not parsed or parsed[1]:
                        continue
                    tab[parsed[0]] = {_norm(clean_store_name(r[0])): str(r[c]).strip()
                                      for r in rows[1:] if r and r[0] and c < len(r) and r[c] not in (None, '')}
                if tab:
                    out['codes'][ws.title] = tab
    finally:
        wb.close()
    try:
        import pandas as pd
        with pd.ExcelFile(path) as xls:
            out['barcodes'] = read_divider_barcodes(xls)
    except Exception:
        out['barcodes'] = {}
    return out


def _plugin_from_rows(header, rows):
    idx = {h: i for i, h in enumerate(header)}
    need = {'Tab', 'Pack', 'Pack Order', 'Pack Group Row', 'Job Number Row', 'First Store Row', 'Last Store Row', 'Store Column'}
    if not need <= set(idx):
        return {}
    get = lambda r, k: r[idx[k]] if k in idx and idx[k] < len(r) else None
    num = lambda v: int(float(v)) if v not in (None, '') and str(v).strip().replace('.', '', 1).isdigit() else None
    tabs = {}
    for r in rows:
        tab = str(get(r, 'Tab') or '').strip()
        if not tab:
            continue
        t = tabs.setdefault(tab, {
            'pack_group_row': num(get(r, 'Pack Group Row')), 'job_id_row': num(get(r, 'Job Number Row')),
            'first_row': num(get(r, 'First Store Row')), 'last_row': num(get(r, 'Last Store Row')),
            'store_col': str(get(r, 'Store Column') or '').strip().upper(), 'packs': []})
        t['packs'].append({
            'name': str(get(r, 'Pack') or '').strip(), 'order': num(get(r, 'Pack Order')),
            'stock_col': str(get(r, 'Stock Column') or '').strip().upper(),
            'stock_col_final': str(get(r, 'Stock Column (Final file)') or '').strip().upper(),
        })
    return tabs


# ---------------------------------------------------------------- the distribution tab

def _row_field(text):
    t = " ".join(str(text).lower().split())
    for field, pattern in ROW_LABELS:
        if re.search(pattern, t):
            return field
    return None


def _image_map(ws, path):
    images = {cell: [data] for cell, data in extract_rich_value_images(path, ws.title).items()}
    for image in getattr(ws, '_images', []):
        try:
            r, c = _floating_image_cell(ws, image.anchor)
            images.setdefault((r, c), []).append(io.BytesIO(image._data()))
        except Exception:
            continue
    return images


def _letter_to_col(letter):
    try:
        return column_index_from_string(letter) if letter else None
    except ValueError:
        return None


def read_horizontal(excel_path, header_row, sheet_name=None, columns=None, tab_no=1, signature=None):
    """Pack groups of one tab, in packing-sheet order. Returns (pack_groups, last_row, found, warnings) like
    parse_packing_data(); raises PackCheckError(errors, warnings) when the labels can't be trusted.
    header_row: the row just above the first store. columns: the Column mapping set by hand (Courier Import fields:
    address columns, and 'pack_<n>' for a pack's Packing Spec column). signature: read_signature_file()."""
    signature = signature or {}
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    try:
        ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active
        return _read(ws, excel_path, header_row, columns or {}, tab_no, signature)
    finally:
        wb.close()


def _read(ws, excel_path, header_row, columns, tab_no, signature):
    spans = _merged_spans(ws)
    plugin = (signature.get('plugin') or {}).get(ws.title)
    sig_codes = (signature.get('codes') or {}).get(ws.title, {})
    errors, warnings = [], []
    L = get_column_letter

    def val(r, c):
        top = spans.get((r, c))
        v = ws.cell(row=top[0], column=top[1]).value if top else ws.cell(row=r, column=c).value
        return '' if v is None else str(v).strip()

    # 1. The job number row: the Label Plugin's, else the row with the most job numbers / code headers
    top_rows = range(1, header_row + 1)
    if plugin and plugin.get('job_id_row'):
        job_row = plugin['job_id_row']
    else:
        score = lambda r: sum(1 for c in range(1, ws.max_column + 1)
                              if JOB.match(val(r, c)) or parse_code_header(val(r, c)))
        job_row = max(top_rows, key=score)
        if not score(job_row):
            raise PackCheckError([_issue('', f"Rows 1-{header_row}", "No Job Number row found above the stores",
                                         "Check the header row, or upload the Signature Links file with its Label Plugin tab.")], warnings)
    marked = [c for c in range(1, ws.max_column + 1) if JOB.match(val(job_row, c)) or parse_code_header(val(job_row, c))]
    first_col, last_col = min(marked), max(marked)

    # 2. The pack row: the Label Plugin's, else the row whose headers / merged cells cover the most item columns
    if plugin and plugin.get('pack_group_row'):
        pack_row = plugin['pack_group_row']
    else:
        # Pack headers are merged over their items (a description or size row isn't), and sit nearest the stores
        def merged(r):
            return sum(1 for c in range(first_col, last_col + 1) if spans.get((r, c)) and spans[(r, c)][2] > 1)

        def filled(r):
            return sum(1 for c in range(first_col, last_col + 1) if val(r, c) or (r, c) in spans)
        candidates = [r for r in range(1, max(header_row, job_row) + 1) if r != job_row and filled(r)]
        if not candidates:
            raise PackCheckError([_issue('', f"Rows 1-{header_row}", "No pack headers found over the item columns")], warnings)
        pack_row = max(candidates, key=lambda r: (merged(r) > 0, merged(r), r == header_row, r))

    # 3. Packs, left to right: a merged cell in the pack row, or one header over one column (as Packing Sheets reads them)
    packs, col_pack = [], {}
    c = first_col
    while c <= last_col:
        top = spans.get((pack_row, c))
        if top:
            start, end = top[1], top[1] + top[2] - 1
            name = val(pack_row, start) or f"Pack_{start}"
        elif val(pack_row, c):
            start = end = c
            name = val(pack_row, c)
        else:
            c += 1
            continue
        packs.append({'name': name, 'order': len(packs) + 1, 'cols': list(range(start, end + 1))})
        for x in range(start, end + 1):
            col_pack[x] = packs[-1]
        c = end + 1
    if plugin:
        known = [p['name'] for p in plugin['packs']]
        for p in packs:
            hit = next((q for q in plugin['packs'] if _norm(q['name']) == _norm(p['name'])), None)
            if hit and hit.get('order'):
                p['order'] = hit['order']
            p['plugin'] = hit
        missing = [n for n in known if not any(_norm(n) == _norm(p['name']) for p in packs)]
        if missing:
            warnings.append(_issue('', f"Row {pack_row}", f"Packs in the Label Plugin not found here: {', '.join(missing)}", major=True))

    # 4. Inside each pack: code columns (headers 'Code for …') and item columns (job numbers)
    for p in packs:
        p['code_cols'], p['sub_cols'], p['items'] = None, [], []
        for x in p['cols']:
            parsed = parse_code_header(val(job_row, x))
            if parsed:
                hname, rng = parsed
                if not names_pack(hname, p['name']):
                    warnings.append(_issue('', f"{L(x)}{job_row}", f"'{val(job_row, x)}' sits in pack {p['name']}"))
                if rng:
                    p['sub_cols'].append((rng, x))
                elif p['code_cols'] is None:
                    p['code_cols'] = x
            elif val(job_row, x):
                p['items'].append(x)
    item_cols = [x for p in packs for x in p['items']]
    if not item_cols:
        raise PackCheckError([_issue('', f"Row {job_row}", "No item columns (job numbers) under the packs")], warnings)

    # 5. Store rows: as Packing Sheets finds them (last store above a TOTAL row), unless the Label Plugin says
    cols = detect_import_columns(_header_cells(ws, header_row))
    spec_cols, spec_warn = detect_pack_columns(ws, header_row, spans)
    spec_cols = {n: c for n, c in spec_cols.items() if c not in col_pack}
    cols, spec_cols = apply_choices(cols, spec_cols, columns)
    warnings += [w for w in spec_warn if not columns]
    if plugin and _letter_to_col(plugin.get('store_col')):
        cols['store_name'] = _letter_to_col(plugin['store_col'])
    store_col = cols.get('store_name') or cols.get('receiver_name')
    if not store_col:
        raise PackCheckError([_issue('', f"Row {header_row}", "No Store Name column found",
                                     "Pick it in the Column mapping, or upload the Signature Links file.")], warnings)
    install_col = next((x for x, v in _header_cells(ws, header_row) if v and _header_field(v) == 'install'), None)
    last_row = plugin.get('last_row') if plugin and plugin.get('last_row') else None
    if not last_row:
        last_row = ws.max_row
        while last_row > header_row and not val(last_row, store_col):
            last_row -= 1
        if TOTAL_ROW.match(val(last_row, store_col)):
            last_row -= 1
    first_row = plugin.get('first_row') if plugin and plugin.get('first_row') else header_row + 1
    store_rows = [r for r in range(first_row, last_row + 1) if val(r, store_col)]
    totals_row = next((r for r in range(last_row + 1, min(ws.max_row, last_row + 3) + 1)
                       if TOTAL_ROW.match(val(r, store_col))), None)

    # 6. Item details from labelled rows above (Description, Size, Artwork …)
    attr_rows = {'job_no': job_row}
    for r in range(1, header_row + 1):
        if r == job_row:
            continue
        label = next((val(r, x) for x in range(first_col - 1, 0, -1) if x not in spec_cols.values() and val(r, x)), '')
        field = _row_field(label) if label else None
        if field and field not in attr_rows:
            attr_rows[field] = r
    # Every other row above the stores with something under the items, as a block of its own for the label's item box,
    # named by what's written in the label column ('Kinds', 'Item Number', a second 'Description') or by its number
    # ('Row 6'): 'row_5' -> that row's value for the item. A label written twice gives a block per row.
    used = set(attr_rows.values()) | {pack_row, header_row}
    extra_rows = []
    for r in range(1, header_row):
        if r in used or not any(val(r, x) for x in item_cols):
            continue
        label = next((val(r, x) for x in range(first_col - 1, 0, -1) if x not in spec_cols.values() and val(r, x)), '')
        extra_rows.append((r, label))
    images = _image_map(ws, excel_path)
    info, no_image = {}, []
    for x in item_cols:
        info[x] = {f: (val(attr_rows[f], x) if f in attr_rows else '') for f in ('job_no', 'desc', 'dimension', 'material', 'notes')}
        info[x].update({f"row_{r}": val(r, x) for r, _ in extra_rows})
        info[x]['thumbnails'] = images.get((attr_rows['thumbnail'], x), []) if 'thumbnail' in attr_rows else []
        if 'thumbnail' in attr_rows and not info[x]['thumbnails']:
            no_image.append(L(x))
    if no_image:
        warnings.append(_issue('', f"{', '.join(no_image)} · row {attr_rows['thumbnail']}", "No artwork image"))

    # 7. Each pack's codes per store: from its code column, else Signature Links, else worked out as Packing Sheets does
    raw = [[ws.cell(row=r, column=x).value for x in range(1, ws.max_column + 1)] for r in range(1, ws.max_row + 1)]
    sig_rows = list(range(pack_row + 1, last_row + 1))  # the rows Packing Sheets coded
    store_key = {r: _norm(clean_store_name(val(r, store_col))) for r in store_rows}
    code_problems, code_warnings = [], []
    for p in packs:
        computed = {}
        if p['items']:
            _, _, _, _, ordered = generate_pack_signatures(raw, sig_rows, min(p['items']), max(p['items']))
            computed = dict(zip(sig_rows, ordered))
        sheet = next((codes for name, codes in sig_codes.items() if names_pack(name, p['name'])), None)
        if sheet is not None:
            # Signature Links attached: its codes make the groups (they're what the packing sheets were made with).
            # Every check against it is a warning, on the preview and the pack's audit page; Generate goes ahead.
            p['codes'] = {r: (sheet.get(store_key[r]) or '').upper() for r in store_rows}
            p['code_source'] = "the Signature Links file"
            has_qty = [r for r in store_rows if computed.get(r)]
            missing = [r for r in has_qty if not p['codes'][r]]
            if missing:  # a store Signature Links doesn't list: its code is worked out from the quantities
                for r in missing:
                    p['codes'][r] = computed.get(r, '')
                code_warnings.append(_issue(_row_ranges(missing), f"Pack {p['name']}", "Store not in Signature Links",
                                            "Its code was worked out from the quantities instead.", major=True))
            stale = [r for r in has_qty if r not in missing and p['codes'][r] != computed.get(r)]
            if stale:
                code_warnings.append(_issue(_row_ranges(stale), f"Pack {p['name']}",
                                            "Signature Links code doesn't match the quantities",
                                            "The Signature Links codes are used for the grouping. The quantities may have "
                                            "changed after Packing Sheets, or the files are from different runs.", major=True))
            if p['code_cols']:
                own = {r: val(r, p['code_cols']).upper() for r in store_rows}
                off = [r for r in store_rows if own[r] != p['codes'][r]]
                if off:
                    code_warnings.append(_issue(_row_ranges(off), f"{L(p['code_cols'])} · Pack {p['name']}",
                                                "Code differs from the Signature Links file",
                                                "The Signature Links codes are used for the grouping; the two files are from "
                                                "different Packing Sheets runs.", major=True))
                p['code_source'] += f", checked against the file's '{val(job_row, p['code_cols'])}' column"
            p['code_source'] += " and the quantities"
        elif p['code_cols']:
            p['codes'] = {r: val(r, p['code_cols']).upper() for r in store_rows}
            p['code_source'] = f"the file's '{val(job_row, p['code_cols'])}' column (column {L(p['code_cols'])})"
            stale = [r for r in store_rows if p['codes'][r] != (computed.get(r) or '')]
            if stale:
                code_problems.append(_issue(_row_ranges(stale), f"{L(p['code_cols'])} · Pack {p['name']}",
                                            "Code doesn't match the quantities",
                                            "The quantities changed after Packing Sheets: run Packing Sheets again."))
        else:
            p['codes'] = {r: computed.get(r, '') for r in store_rows}
            p['code_source'] = "worked out from the quantities (no code column, no Signature Links)"
        p['subs'] = {rng: {r: val(r, x).upper() for r in store_rows} for rng, x in p['sub_cols']}
        p['code_header'] = (val(job_row, p['code_cols']) if p['code_cols'] else
                            f"Code for Pack {p['name']}" if not str(p['name']).lower().startswith('pack') else f"Code for {p['name']}")
        p['sub_headers'] = [(rng, val(job_row, x)) for rng, x in p['sub_cols']]
    errors += code_problems
    warnings += code_warnings

    # 8. One pack group per store and pack that has quantities
    n_stores = len(store_rows)
    width = max(2, len(str(n_stores)))
    names = Counter(store_key.values())
    groups, bad_qty, x_but_qty, no_spec, spec_no_qty = {}, [], [], [], []
    grid_units = label_units = 0
    lost = []  # (row, column, units): quantities in the grid that no label carries, for the check's coordinates
    for number, r in enumerate(store_rows, start=1):
        receiver = val(r, cols['receiver_name']) if 'receiver_name' in cols else ''
        store = clean_store_name(val(r, cols['store_name'])) if 'store_name' in cols else receiver
        store = store or receiver
        parts = {k: val(r, cols[k]) for k in ADDRESS_FIELDS + ('address_3',) if k in cols}
        line2 = ", ".join(v for v in (parts.get('address_2', ''), parts.get('address_3', '')) if v)
        install = is_install_flag(val(r, install_col)) if install_col else False
        for p in packs:
            items = []
            for x in p['items']:
                q = sanitize_cell(ws.cell(row=r, column=x).value)
                if isinstance(q, str):
                    bad_qty.append((r, x))
                    continue
                if q:
                    grid_units += q
                    items.append({**{k: v for k, v in info[x].items()}, 'qty': f"{q:g}" if isinstance(q, float) else str(q),
                                  'install': "INSTALLER" if install else "", 'source_key': (ws.title, x)})
            spec = _spec_for(p, r, val, spec_cols, attr_rows, any(p['code_cols'] or p['sub_cols'] for p in packs))
            if items and spec.lower() in NO_PACK:
                (x_but_qty if spec else no_spec).append((r, p))
                lost += [(r, i['source_key'][1], float(i['qty'])) for i in items]
                continue
            if not items:
                if spec and spec.lower() not in NO_PACK and p.get('spec_from') == 'column':
                    spec_no_qty.append((r, p))
                continue
            label_units += sum(float(i['qty']) for i in items)
            code = p['codes'].get(r, '')
            subs = [(rng, codes.get(r, '')) for rng, codes in p['subs'].items()]
            # Item boxes for the codes, first on the label (where the code columns sit in the file): the code
            # column's header where a job number goes, the store's code where the quantity goes
            code_cells = [_code_cell(p['code_header'], code)] + [_code_cell(h, c) for (rng, h), c in
                                                                 zip(p['sub_headers'], (c for _, c in subs)) if c]
            # With Signature Links the code is in the label's own code box: the packing spec box just names the pack
            tag = (f"PACK {p['order']}" if sig_codes else
                   f"PACK {p['order']} \u00b7 CODE {code or '?'}" + "".join(f" \u00b7 {rng}: {c}" for rng, c in subs if c))
            serial = f"{number:0{width}d}T{tab_no}P{p['order']}"
            groups[f"ROW_{r}_P{p['order']}"] = {
                'pack_spec_name': spec, 'store_name': store, 'receiver_name': receiver,
                'contact': attn_text(val(r, cols['contact'])) if 'contact' in cols else '', 'install': install,
                'address_1': parts.get('address_1', ''), 'address_2': line2,
                **{k: parts.get(k, '') for k in ('suburb', 'state', 'postcode', 'country')},
                'pack_no': p['order'], 'pack_name': p['name'], 'pack_columns': len(p['items']), 'pack_code': code, 'sub_codes': subs,
                'pack_tag': tag, 'code_cells': code_cells, 'code_source': p['code_source'], 'serial': serial, 'store_no': number,
                'tab_no': tab_no, 'tab_name': ws.title, 'excel_rows': [r], 'items': items,
            }

    # 9. Checks, each with its Excel row and column
    for r, x in bad_qty:
        errors.append(_issue(str(r), L(x), "Quantity isn't a number"))
    for r, p in x_but_qty:
        errors.append(_issue(str(r), f"Pack {p['name']}", f"Pack {p['name']} marked x but has quantities",
                             "Give it a Packing Spec, or clear its quantities."))
    for r, p in no_spec:
        errors.append(_issue(str(r), f"Pack {p['name']}", "No Packing Spec for this pack",
                             "Add the pack's Packing Spec column (headed with the pack's number), or a Packing Spec row."))
    for r, p in spec_no_qty:
        warnings.append(_issue(str(r), f"Pack {p['name']}", "Packing Spec but no quantities: no label made",
                               "The store has this pack's spec but no units in it. Check the distribution file.", major=True))
    uncoded = [g['excel_rows'][0] for g in groups.values() if not g['pack_code']]
    if uncoded:
        errors.append(_issue(_row_ranges(uncoded), 'Code', "Store gets a pack but has no packing-sheet code"))
    dupes = [s for s, k in names.items() if k > 1 and s]
    if dupes:
        warnings.append(_issue('', L(store_col), f"Store listed twice: {', '.join(dupes)}",
                               "The packing sheets and the Label Shuffler can't tell them apart.", major=True))
    no_barcode = sorted({i['job_no'] for g in groups.values() for i in g['items'] if i['job_no'] and code128_modules(i['job_no']) is None})
    if no_barcode:
        warnings.append(_issue('', f"Row {job_row}", f"Job number can't be a barcode: {', '.join(no_barcode)}", major=True))
    if abs(grid_units - label_units) > 1e-9:
        # Where they are: each cell as Excel names it (J12) with its units, and its rows and columns
        cells = ", ".join(f"{L(x)}{r} ({q:g})" for r, x, q in sorted(lost, key=lambda c: (c[0], c[1]))[:12])
        more = f" and {len(lost) - 12} more" if len(lost) > 12 else ''
        cols_of = ", ".join(L(x) for x in sorted({x for _, x, _ in lost}))
        errors.append(_issue(_row_ranges(sorted({r for r, _, _ in lost})) if lost else '',
                             f"{cols_of} · Quantities" if lost else 'Quantities',
                             f"{grid_units - label_units:g} units in the grid are on no label",
                             (f"In cell{'s' if len(lost) != 1 else ''} {cells}{more}. " if lost else '') +
                             "Fix the problems above: every unit must land on a packing label."))
    totals = [sanitize_cell(raw[totals_row - 1][x - 1]) for x in item_cols] if totals_row else []
    totals = [t for t in totals if isinstance(t, (int, float)) and t]
    if totals:  # a TOTAL row with its numbers in: they must add up to the stores' quantities
        sheet_total = sum(totals)
        if abs(sheet_total - grid_units) > 1e-9:
            warnings.append(_issue(str(totals_row), 'TOTAL', f"TOTAL row says {sheet_total:g}, the stores add up to {grid_units:g}", major=True))
    with_stores(errors + warnings, lambda r: val(r, store_col))  # 'Row 11 Macpac Adelaide', not only 'Row 11'
    if errors:
        raise PackCheckError(errors, warnings)

    # 10. Packing-sheet order: pack, code (A … Z, AA …), sub-group codes, the store's row
    order = sorted(groups, key=lambda k: (groups[k]['pack_no'], code_key(groups[k]['pack_code']),
                                          [code_key(c) for _, c in groups[k]['sub_codes']], groups[k]['store_no']))
    found = {'packs': {f"{p['order']}. {p['name']}": f"{L(p['cols'][0])}-{L(p['cols'][-1])}" for p in packs},
             # Where the codes came from, in a word or two for the tab's card (each pack's detail is on its audit page;
             # a difference is a warning of its own, so this stays quiet)
             'codes': sorted({'Signature Links' if p['code_source'].startswith('the Signature Links') else
                              "the file's code columns" if p['code_source'].startswith("the file's") else
                              'worked out from the quantities' for p in packs}),
             'rows': {'job': job_row, 'pack': pack_row, 'stores': f"{first_row}-{last_row}"},
             'plugin': bool(plugin), 'stores': n_stores,
             # Each pack's Packing Specs and how many labels each, for the tab's card: [{'pack', 'specs': [[spec, n]]}]
             'specs_by_pack': [{'pack': f"Pack {p['order']}" + ('' if str(p['name']).strip() in ('', str(p['order'])) else f" ({p['name']})"),
                                'specs': [[s, n] for s, n in Counter(groups[k]['pack_spec_name'] for k in order
                                                                     if groups[k]['pack_no'] == p['order']).most_common()]}
                               for p in packs]}
    # The cell layout's blocks by the names parse_packing_data() gives them: the page offers a block only when
    # its details were found (job numbers and quantities always are; the rest come from the labelled rows above)
    found.update({
        'Packing Spec': True, 'Store Name': True, 'Job Number': True, 'Quantity': True,
        'Address Info': any(k in cols for k in ADDRESS_FIELDS),
        'Install': bool(install_col),
        'Description': 'desc' in attr_rows, 'Dimensions': 'dimension' in attr_rows, 'Material': 'material' in attr_rows,
        'Notes': 'notes' in attr_rows, 'Thumbnail': 'thumbnail' in attr_rows,
        # The other rows, each a block the page offers: [{'id': 'row_5', 'label': 'Kinds (row 5)'}]
        'Extra rows': [{'id': f"row_{r}", 'label': f"{label} (row {r})" if label else f"Row {r}"} for r, label in extra_rows],
    })
    return {k: groups[k] for k in order}, last_row, found, warnings


def _spec_for(p, r, val, spec_cols, attr_rows, final_file):
    """A store's Packing Spec for a pack: its column headed with the pack's number (Courier Import's), else the
    Label Plugin's stock column for the pack, else a Packing Spec row above the pack's items."""
    n = int(p['name']) if p['name'].isdigit() else p['order']
    if n in spec_cols:
        p['spec_from'] = 'column'
        return val(r, spec_cols[n])
    plug = p.get('plugin') or {}
    col = _letter_to_col(plug.get('stock_col_final') if final_file else plug.get('stock_col')) or _letter_to_col(plug.get('stock_col'))
    if col:
        p['spec_from'] = 'column'
        return val(r, col)
    if 'pack_spec' in attr_rows:
        p['spec_from'] = 'row'
        return next((val(attr_rows['pack_spec'], x) for x in p['items'] if val(attr_rows['pack_spec'], x)), '')
    return ''


def _code_cell(header, code):
    """A packing-sheet code as an item box: the header in the job number's place, the code in the quantity's.
    No barcode (it isn't a piece to scan), and never counted as units or job numbers."""
    return {'job_no': header, 'qty': code or '?', 'desc': '', 'dimension': '', 'material': '', 'install': '',
            'notes': '', 'thumbnails': [], 'barcode': '', 'is_code': True}


def item_references(pack_groups):
    """[(Item Reference, code)] per pack in order: ('03T1P2 Chadstone', '03T1P2'). The code is what Stitch Labels
    reads back off the courier label (label_stitcher.find_code)."""
    return [(f"{g['serial']} {g['store_name']}", g['serial']) for g in pack_groups.values()]


def pack_blocks(pack_groups):
    """The packs in order as [((tab_no, pack_no), [keys])], one block per tab and pack (an audit page each)."""
    blocks = []
    for k, g in pack_groups.items():
        key = (g.get('tab_no', 1), g['pack_no'])
        if not blocks or blocks[-1][0] != key:
            blocks.append((key, []))
        blocks[-1][1].append(k)
    return blocks


def pack_summary(pack_groups, keys):
    """What a pack's audit page shows: its item columns, and per code the stores, boxes and quantity per item."""
    cols, seen = [], set()
    for k in keys:
        for i in pack_groups[k]['items']:
            if i['source_key'] not in seen:
                seen.add(i['source_key'])
                cols.append((i['source_key'], i['job_no'], i['desc']))
    cols.sort(key=lambda c: c[0][1])
    by_code = defaultdict(lambda: {'stores': 0, 'boxes': 0, 'qty': Counter(), 'serials': []})
    for k in keys:
        g = pack_groups[k]
        row = by_code[g['pack_code']]
        row['stores'] += 1
        row['boxes'] += split_count(g['pack_spec_name'])[0]
        row['serials'].append(g['serial'])
        for i in g['items']:
            row['qty'][i['source_key']] = float(i['qty'])
    codes = sorted(by_code, key=code_key)
    return cols, [(c, by_code[c]) for c in codes]
