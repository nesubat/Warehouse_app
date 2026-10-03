import re
import pandas as pd
import openpyxl
import io
import math
import zipfile
import xml.etree.ElementTree as ET
import pymupdf as fitz
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string, get_column_letter

def extract_rich_value_images(excel_path, sheet_name):
    ns = {
        'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'xlrd': 'http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel'
    }
    image_map = {}
    try:
        with zipfile.ZipFile(excel_path, 'r') as z:
            wb_xml = z.read('xl/workbook.xml')
            wb_root = ET.fromstring(wb_xml)
            sheet_id = None
            for sheet in wb_root.findall('.//main:sheet', ns):
                if sheet.get('name') == sheet_name:
                    sheet_id = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    break
            if not sheet_id: return {}
            
            wb_rels_xml = z.read('xl/_rels/workbook.xml.rels')
            wb_rels_root = ET.fromstring(wb_rels_xml)
            sheet_path = None
            for rel in wb_rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                if rel.get('Id') == sheet_id:
                    sheet_path = 'xl/' + rel.get('Target')
                    break
            if not sheet_path: return {}
            
            sheet_xml = z.read(sheet_path)
            sheet_root = ET.fromstring(sheet_xml)
            cell_vm_map = {} 
            for row in sheet_root.findall('.//main:row', ns):
                r_idx = int(row.get('r'))
                for c in row.findall('main:c', ns):
                    vm = c.get('vm')
                    if vm:
                        ref = c.get('r')
                        col_letter = "".join(filter(str.isalpha, ref))
                        cell_vm_map[(r_idx, col_letter)] = vm
            if not cell_vm_map: return {}
                
            meta_xml = z.read('xl/metadata.xml')
            meta_root = ET.fromstring(meta_xml)
            vm_nodes = meta_root.find('.//main:valueMetadata', ns)
            vm_to_rvb = {}
            if vm_nodes is not None:
                for i, bk in enumerate(vm_nodes.findall('main:bk', ns)):
                    rc = bk.find('main:rc', ns)
                    if rc is not None: vm_to_rvb[str(i+1)] = rc.get('v')
                        
            rv_xml = z.read('xl/richData/richValueRel.xml')
            rv_root = ET.fromstring(rv_xml)
            rvb_to_rid = {}
            for i, rel in enumerate(rv_root.findall('.//xlrd:rel', ns)):
                rvb_to_rid[str(i)] = rel.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                
            rv_rels_xml = z.read('xl/richData/_rels/richValueRel.xml.rels')
            rv_rels_root = ET.fromstring(rv_rels_xml)
            rid_to_path = {}
            for rel in rv_rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                rid_to_path[rel.get('Id')] = rel.get('Target').replace('../', 'xl/')
                
            for (r, col_letter), vm in cell_vm_map.items():
                rvb = vm_to_rvb.get(vm)
                if rvb:
                    rid = rvb_to_rid.get(rvb)
                    if rid:
                        media_path = rid_to_path.get(rid)
                        if media_path:
                            col_idx = column_index_from_string(col_letter)
                            image_map[(r, col_idx)] = io.BytesIO(z.read(media_path))
    except Exception as e:
        pass
    return image_map

EMU_PER_PT = 12700
EMU_PER_PX = 9525


def _row_height_emu(ws, row):
    h = ws.row_dimensions[row].height if row in ws.row_dimensions else None
    return (h or ws.sheet_format.defaultRowHeight or 15) * EMU_PER_PT


def _col_width_emu(ws, col):
    letter = get_column_letter(col)
    w = ws.column_dimensions[letter].width if letter in ws.column_dimensions else None
    w = w or ws.sheet_format.defaultColWidth or 8.43
    return (w * 7 + 5) * EMU_PER_PX


def _cell_at(offset_emu, size_of):
    """1-based row/col index containing a sheet-absolute EMU offset."""
    idx, edge = 1, 0
    while edge + size_of(idx) <= offset_emu and idx < 100000:
        edge += size_of(idx)
        idx += 1
    return idx


def _start_emu(index0, offset, size_of):
    return sum(size_of(i) for i in range(1, index0 + 1)) + offset


def _floating_image_cell(ws, anchor):
    """Cell holding the image's centre, so an image that pokes into a neighbouring row still maps to its own row."""
    row_size = lambda r: _row_height_emu(ws, r)
    col_size = lambda c: _col_width_emu(ws, c)
    top = _start_emu(anchor._from.row, anchor._from.rowOff, row_size)
    left = _start_emu(anchor._from.col, anchor._from.colOff, col_size)
    if getattr(anchor, 'to', None) is not None:
        bottom = _start_emu(anchor.to.row, anchor.to.rowOff, row_size)
        right = _start_emu(anchor.to.col, anchor.to.colOff, col_size)
    elif getattr(anchor, 'ext', None) is not None:
        bottom = top + anchor.ext.height
        right = left + anchor.ext.width
    else:
        return anchor._from.row + 1, anchor._from.col + 1
    return _cell_at((top + bottom) / 2, row_size), _cell_at((left + right) / 2, col_size)


def _row_ranges(rows):
    """[5, 6, 7, 10] -> '5-7, 10'"""
    parts, start, prev = [], None, None
    for r in sorted(rows):
        if start is None:
            start = prev = r
        elif r == prev + 1:
            prev = r
        else:
            parts.append(f"{start}-{prev}" if prev > start else str(start))
            start = prev = r
    if start is not None:
        parts.append(f"{start}-{prev}" if prev > start else str(start))
    return ", ".join(parts)


ADDRESS_FIELDS = ('address_1', 'address_2', 'suburb', 'state', 'postcode', 'country')
INSTALL_YES = {'y', 'yes', 'true'}


def is_install_flag(value):
    return str(value or '').strip().lower() in INSTALL_YES


def _norm(text):
    """Comparison key: ignores capitals, spacing and punctuation ('12 Bourke Rd.' == '12  BOURKE RD')."""
    return " ".join(re.sub(r"[^0-9a-z]+", " ", str(text or '').lower()).split())


class PackCheckError(ValueError):
    """Blocking problems found in the sheet; carries the structured issues for the preview table."""
    def __init__(self, issues, warnings):
        self.issues, self.warnings = issues, warnings
        super().__init__("\n".join(_issue_line(i) for i in issues))


def _issue(rows, col, text, detail=''):
    return {'rows': rows, 'col': col, 'text': text, 'detail': detail}


def _issue_line(i):
    return " · ".join(x for x in (f"Row {i['rows']}" if i['rows'] else '', i['col'], i['text']) if x)


def _span(rows):
    return _row_ranges(list(range(rows[0]['row'], rows[-1]['row'] + 1)))


def _short(text, limit=40):
    return text if len(text) <= limit else text[:limit - 1].rstrip() + "…"


STREET_TYPES = {
    'street': 'st', 'road': 'rd', 'avenue': 'ave', 'boulevard': 'blvd', 'highway': 'hwy',
    'drive': 'dr', 'parade': 'pde', 'place': 'pl', 'court': 'ct', 'crescent': 'cres',
    'lane': 'ln', 'terrace': 'tce', 'arcade': 'arc',
}


def _address_key(text):
    """Like _norm, plus 'Street' == 'St' etc. Only after a name, so '12 St Kilda Rd' keeps its Saint."""
    words = _norm(text).split()
    return " ".join(
        STREET_TYPES.get(w, w) if i and words[i - 1].isalpha() else w
        for i, w in enumerate(words)
    )


def _distinct(rows, key):
    found = {}
    for r in rows:
        if r[key]:
            found.setdefault(r[key], []).append(r['row'])
    return found


def _variants(rows, key, display):
    """'4-5: 49 Church Street… · 58: 49 Church St…'"""
    by_value = {}
    for r in rows:
        if r[key]:
            by_value.setdefault(r[key], (r[display], []))[1].append(r['row'])
    # Skip the leading words a variant shares with the first one, so the detail shows where they differ
    def shared(a, b):
        n = 0
        while n < min(len(a), len(b)) and _norm(a[n]) == _norm(b[n]):
            n += 1
        return max(0, n - 1)

    variants = [(text.split(), rs) for text, rs in by_value.values()]
    ref = variants[0][0]
    parts = []
    for idx, (words, rs) in enumerate(variants):
        same = max(shared(ref, w) for w, _ in variants[1:]) if idx == 0 else shared(ref, words)
        parts.append(f"{_row_ranges(rs)}: {'…' if same else ''}{_short(' '.join(words[same:]))}")
    return " · ".join(parts)


def _check_pack_consistency(pack_groups, pack_rows, has_address, col):
    errors, warnings = [], []
    no_address_packs = []

    for pack_id, rows in pack_rows.items():
        span = _span(rows)
        pack = pack_groups[pack_id]['pack_spec_name']

        stores = _distinct(rows, 'store_key')
        if len(stores) > 1:
            errors.append(_issue(span, col['store_name'], f"Store name differs in pack {pack}", _variants(rows, 'store_key', 'store')))
        elif stores and any(not r['store_key'] for r in rows):
            blank = [r['row'] for r in rows if not r['store_key']]
            errors.append(_issue(_row_ranges(blank), col['store_name'], f"Store name blank in pack {pack}"))

        if len({r['install'] for r in rows}) > 1:
            flagged = [r['row'] for r in rows if r['install']]
            errors.append(_issue(span, col['install'], f"Install mixed in pack {pack}", f"Y on {_row_ranges(flagged)}"))

        if has_address:
            addresses = _distinct(rows, 'address_key')
            if len(addresses) > 1:
                errors.append(_issue(span, col['address'], f"Address differs in pack {pack}", _variants(rows, 'address_key', 'address')))
            elif addresses and any(not r['address_key'] for r in rows):
                blank = [r['row'] for r in rows if not r['address_key']]
                errors.append(_issue(_row_ranges(blank), col['address'], f"Address blank in pack {pack}"))
            elif not addresses:
                no_address_packs.append(rows[0]['row'])

            if any(r['crosses_pack'] for r in rows):
                warnings.append(_issue(span, col['address'], f"Merged address shared with next pack"))

    if no_address_packs:
        warnings.append(_issue(_row_ranges(no_address_packs), col['address'], "No address"))

    # Across packs of the same store: store deliveries must agree, installer deliveries may differ
    if has_address:
        by_store = {}
        for rows in pack_rows.values():
            first = rows[0]
            if first['store_key'] and first['address_key']:
                by_store.setdefault(first['store_key'], []).append(first)
        for store_rows in by_store.values():
            for install in (False, True):
                group = [r for r in store_rows if r['install'] == install]
                if len(_distinct(group, 'address_key')) > 1:
                    rows_text = _row_ranges([r['row'] for r in group])
                    name = group[0]['store']
                    if install:
                        warnings.append(_issue(rows_text, col['address'], f"{name}: installer addresses differ", _variants(group, 'address_key', 'address')))
                    else:
                        errors.append(_issue(rows_text, col['address'], f"{name}: store address differs between packs", _variants(group, 'address_key', 'address')))

    # Neighbouring packs with the same spec and store usually mean the Packing Spec cells weren't merged
    ids = list(pack_rows)
    for prev_id, next_id in zip(ids, ids[1:]):
        prev, nxt = pack_groups[prev_id], pack_groups[next_id]
        if _norm(prev['pack_spec_name']) == _norm(nxt['pack_spec_name']) and _norm(prev['store_name']) == _norm(nxt['store_name']):
            span = _row_ranges(list(range(pack_rows[prev_id][0]['row'], pack_rows[next_id][-1]['row'] + 1)))
            warnings.append(_issue(span, col['packing_spec'], f"Same spec {nxt['pack_spec_name']} split into separate packs — merge if one box"))

    return errors, warnings


def parse_packing_data(excel_path, header_row, sheet_name=None):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active
    
    image_map = extract_rich_value_images(excel_path, ws.title)
    stacked_image_rows = set()
    
    for image in getattr(ws, '_images', []):
        try:
            r, c = None, None
            if hasattr(image, 'anchor'):
                if hasattr(image.anchor, '_from'):
                    r, c = _floating_image_cell(ws, image.anchor)
                elif isinstance(image.anchor, str):
                    c_letter, r_str = coordinate_from_string(image.anchor)
                    c = column_index_from_string(c_letter)
                    r = int(r_str)
                    
            if r and c and (r, c) in image_map:
                stacked_image_rows.add(r)
            elif r and c:
                img_bytes = None
                if hasattr(image, 'ref') and hasattr(image.ref, 'getvalue'):
                    img_bytes = image.ref.getvalue()
                elif hasattr(image, '_data'):
                    img_bytes = image._data() if callable(image._data) else image._data
                
                if img_bytes: image_map[(r, c)] = io.BytesIO(img_bytes)
        except Exception:
            continue

    cols = {}
    for col_idx in range(1, ws.max_column + 1):
        val = ws.cell(row=header_row, column=col_idx).value
        if not val: continue
        val_str = str(val).strip().lower()
        
        if 'packing spec' in val_str: cols['packing_spec'] = col_idx
        elif 'store' in val_str or 'retailer' in val_str: cols['store_name'] = col_idx
        elif 'job' in val_str: cols['job_no'] = col_idx
        elif 'qty' in val_str or 'quantity' in val_str: cols['qty'] = col_idx
        elif 'desc' in val_str: cols['desc'] = col_idx
        elif 'thumb' in val_str or 'image' in val_str or 'picture' in val_str or 'art' in val_str: cols['thumbnail'] = col_idx
        elif 'dimension' in val_str: cols['dim_combined'] = col_idx
        elif 'width' in val_str or val_str == 'w': cols['dim_w'] = col_idx
        elif 'height' in val_str or val_str == 'h': cols['dim_h'] = col_idx
        elif 'address line 1' in val_str or 'street address' in val_str or val_str == 'address': cols['address_1'] = col_idx
        elif 'address line 2' in val_str: cols['address_2'] = col_idx
        elif 'suburb' in val_str: cols['suburb'] = col_idx
        elif 'state' in val_str: cols['state'] = col_idx
        elif 'postcode' in val_str or 'post code' in val_str: cols['postcode'] = col_idx
        elif 'country' in val_str: cols['country'] = col_idx
        elif 'material' in val_str: cols['material'] = col_idx
        elif 'note' in val_str: cols['notes'] = col_idx
        elif 'install' in val_str: cols['install'] = col_idx

    found_headers = {
        'Packing Spec': 'packing_spec' in cols,
        'Store Name': 'store_name' in cols,
        'Job Number': 'job_no' in cols,
        'Quantity': 'qty' in cols,
        'Description': 'desc' in cols,
        'Thumbnail': 'thumbnail' in cols,
        'Dimensions': 'dim_combined' in cols or 'dim_w' in cols or 'dim_h' in cols,
        'Address Info': 'address_1' in cols or 'suburb' in cols or 'state' in cols,
        'Material': 'material' in cols,
        'Notes': 'notes' in cols,
        'Install': 'install' in cols
    }

    if 'packing_spec' not in cols:
        raise ValueError(f"Could not find a 'Packing Spec' column in Row {header_row}.")
    if 'job_no' not in cols:
        raise ValueError(f"Could not find a 'Job Number' column in Row {header_row}.")

    def get_cell_info(row, col):
        for merged_range in ws.merged_cells.ranges:
            if merged_range.min_col <= col <= merged_range.max_col and merged_range.min_row <= row <= merged_range.max_row:
                val = ws.cell(row=merged_range.min_row, column=merged_range.min_col).value
                return val, merged_range.min_row, merged_range.max_row
        return ws.cell(row=row, column=col).value, row, row

    pack_groups = {}
    pack_rows = {}
    rows_missing_job = []
    rows_without_image = []
    rows_ambiguous_image = []
    address_keys = [k for k in ADDRESS_FIELDS if k in cols]
    col_ref = {k: f"{get_column_letter(c)} · {str(ws.cell(row=header_row, column=c).value or '').strip().title()}" for k, c in cols.items()}
    addr_letters = [get_column_letter(cols[k]) for k in address_keys]
    col_ref['address'] = f"{addr_letters[0]}–{addr_letters[-1]} · Address" if len(addr_letters) > 1 else (f"{addr_letters[0]} · Address" if addr_letters else "Address")
    col_ref.setdefault('store_name', "Store Name")
    col_ref.setdefault('install', "Install")
    col_ref.setdefault('thumbnail', "Thumbnail")
    current_row = header_row + 1
    last_processed_row = current_row

    while True:
        pack_spec_val, pack_top_row, pack_bottom_row = get_cell_info(current_row, cols['packing_spec'])

        if pack_spec_val is None or str(pack_spec_val).strip() == "": break

        packing_spec = str(pack_spec_val).strip()
        pack_id = f"ROW_{pack_top_row}_{packing_spec}"

        def get_val(key):
            if key in cols:
                val, _, _ = get_cell_info(current_row, cols[key])
                return str(val).strip() if val is not None else ""
            return ""

        store_name = get_val('store_name')
        address_parts = {k: get_val(k) for k in address_keys}
        install = is_install_flag(get_val('install'))

        # Address cells merged past this pack's edges (shared with a neighbouring pack)
        crosses_pack = any(
            lo < pack_top_row or hi > pack_bottom_row
            for _, lo, hi in (get_cell_info(current_row, cols[k]) for k in address_keys)
        )

        if pack_id not in pack_groups:
            pack_groups[pack_id] = {
                'pack_spec_name': packing_spec,
                'store_name': store_name,
                'install': install,
                **{k: address_parts.get(k, '') for k in ADDRESS_FIELDS},
                'items': []
            }
            pack_rows[pack_id] = []

        pack_rows[pack_id].append({
            'row': current_row,
            'store': store_name,
            'store_key': _norm(store_name),
            'address': ", ".join(v for v in address_parts.values() if v),
            'address_key': _address_key(" ".join(address_parts.values())),
            'install': install,
            'crosses_pack': crosses_pack,
        })

        dim_w = get_val('dim_w')
        dim_h = get_val('dim_h')
        dim_combined = get_val('dim_combined')

        if dim_combined: dimension = dim_combined
        elif dim_w and dim_h: dimension = f"{dim_w} x {dim_h}"
        elif dim_w: dimension = dim_w
        elif dim_h: dimension = dim_h
        else: dimension = ""

        thumbnail_bytes = None
        if 'thumbnail' in cols:
            tc = cols['thumbnail']
            # Same row only: looking at neighbouring rows let an image-less row borrow the image above it
            search_coords = [(current_row, tc), (current_row, tc - 1), (current_row, tc + 1)]
            matches = [coord for coord in search_coords if coord in image_map]
            if matches:
                thumbnail_bytes = image_map[matches[0]]
            else:
                rows_without_image.append(current_row)
            if len(matches) > 1 or current_row in stacked_image_rows:
                rows_ambiguous_image.append(current_row)

        item = {
            'job_no': get_val('job_no'),
            'qty': get_val('qty') or "1",
            'desc': get_val('desc'),
            'dimension': dimension,
            'material': get_val('material'),
            'install': "INSTALLER" if install else "",
            'notes': get_val('notes'),
            'thumbnail_bytes': thumbnail_bytes
        }

        if not item['job_no']:
            rows_missing_job.append(current_row)
        pack_groups[pack_id]['items'].append(item)
        last_processed_row = current_row
        current_row += 1

    rows_missing_spec = []
    for row in range(current_row, ws.max_row + 1):
        spec_val, _, _ = get_cell_info(row, cols['packing_spec'])
        job_val, _, _ = get_cell_info(row, cols['job_no'])
        if (spec_val is None or str(spec_val).strip() == "") and job_val is not None and str(job_val).strip():
            rows_missing_spec.append(row)

    wb.close()

    errors = []
    if rows_missing_spec:
        errors.append(_issue(_row_ranges(rows_missing_spec), col_ref['packing_spec'], "Packing Spec empty"))
    if rows_missing_job:
        errors.append(_issue(_row_ranges(rows_missing_job), col_ref['job_no'], "Job Number empty"))
    if not pack_groups and not errors:
        errors.append(_issue('', '', f"No data found below row {header_row}"))

    warnings = []
    if rows_without_image:
        warnings.append(_issue(_row_ranges(rows_without_image), col_ref['thumbnail'], "No image"))
    if rows_ambiguous_image:
        warnings.append(_issue(_row_ranges(rows_ambiguous_image), col_ref['thumbnail'], "More than one image"))

    pack_errors, pack_warnings = _check_pack_consistency(pack_groups, pack_rows, bool(address_keys), col_ref)
    errors += pack_errors
    warnings += pack_warnings

    if errors:
        raise PackCheckError(errors, warnings)

    return pack_groups, last_processed_row, found_headers, warnings

TEXT_STYLES = {
    'desc': (9, "helv", (0, 0, 0)),
    'dimension': (9, "hebo", (0, 0, 0)),
    'material': (9, "helv", (0, 0, 0)),
    'install': (9, "hebo", (0.8, 0, 0)),
    'notes': (9, "helv", (0, 0, 0)),
}
LINE = 1.2


def _clean(val):
    val = str(val or '').strip()
    return '' if val.lower() == 'none' else val


def _wrap(text, fontname, size, width):
    lines = []
    for paragraph in text.splitlines() or ['']:
        line = ''
        for word in paragraph.split():
            while fitz.get_text_length(word, fontname=fontname, fontsize=size) > width:
                cut = len(word)
                while cut > 1 and fitz.get_text_length(word[:cut], fontname=fontname, fontsize=size) > width:
                    cut -= 1
                if line:
                    lines.append(line)
                    line = ''
                lines.append(word[:cut])
                word = word[cut:]
            candidate = f"{line} {word}".strip()
            if line and fitz.get_text_length(candidate, fontname=fontname, fontsize=size) > width:
                lines.append(line)
                line = word
            else:
                line = candidate
        if line:
            lines.append(line)
    return lines


def _draw_lines(page, lines, x0, x1, y, fontname, size, color):
    for line in lines:
        w = fitz.get_text_length(line, fontname=fontname, fontsize=size)
        page.insert_text((x0 + (x1 - x0 - w) / 2, y + size * 0.92), line,
                         fontname=fontname, fontsize=size, color=color)
        y += size * LINE
    return y


def _layout_cell(item, attribute_order, cell, text_scale, thumb_scale, draw_page=None):
    """Stacks the chosen blocks top-down; returns the total height used. Draws only when draw_page is given."""
    pad = 4
    inner_x0, inner_x1 = cell.x0 + pad, cell.x1 - pad
    inner_w = inner_x1 - inner_x0
    gap = 3 * text_scale
    y = cell.y0 + pad
    for attr in attribute_order:
        if attr == 'thumbnail':
            if not item.get('thumbnail_bytes'):
                continue
            h = cell.height * 0.35 * thumb_scale
            w = cell.width * 0.7
            if draw_page:
                rect = fitz.Rect(cell.x0 + (cell.width - w) / 2, y, cell.x0 + (cell.width + w) / 2, y + h)
                try:
                    draw_page.insert_image(rect, stream=item['thumbnail_bytes'].getvalue())
                except Exception:
                    draw_page.insert_textbox(rect, "Image Error", fontsize=8, align=1)
            y += h + gap

        elif attr == 'job_no':
            val = _clean(item.get('job_no'))
            if not val:
                continue
            size = 14 * text_scale
            lines = _wrap(val, "hebo", size, inner_w - 4)
            h = len(lines) * size * LINE + 6 * text_scale
            if draw_page:
                _round_rect(draw_page, fitz.Rect(cell.x0 + 3, y, cell.x1 - 3, y + h), 3, fill=(0, 0, 0))
                _draw_lines(draw_page, lines, inner_x0, inner_x1, y + 3 * text_scale, "hebo", size, (1, 1, 1))
            y += h + gap

        elif attr == 'qty':
            val = _clean(item.get('qty'))
            if not val:
                continue
            size = 20 * text_scale
            border = 2.5 * text_scale
            lines = _wrap(val, "hebo", size, inner_w - 2 * border - 8)
            widest = max(fitz.get_text_length(l, fontname="hebo", fontsize=size) for l in lines)
            box_w = min(inner_w, max(70 * text_scale, widest + 16 * text_scale))
            h = len(lines) * size * LINE + 6 * text_scale
            if draw_page:
                bx0 = cell.x0 + (cell.width - box_w) / 2
                half = border / 2
                draw_page.draw_rect(fitz.Rect(bx0 + half, y + half, bx0 + box_w - half, y + h - half), color=(0, 0, 0), width=border)
                _draw_lines(draw_page, lines, bx0, bx0 + box_w, y + 3 * text_scale, "hebo", size, (0, 0, 0))
            y += h + gap

        elif attr in TEXT_STYLES:
            val = _clean(item.get(attr))
            if not val:
                continue
            base, fontname, color = TEXT_STYLES[attr]
            size = base * text_scale
            lines = _wrap(val, fontname, size, inner_w)
            if draw_page:
                _draw_lines(draw_page, lines, inner_x0, inner_x1, y, fontname, size, color)
            y += len(lines) * size * LINE + gap
    return y - gap - cell.y0 + pad


def _draw_cell(page, item, attribute_order, cell):
    # Shrink text first (down to 65%), then the thumbnail, until the whole stack fits inside the cell
    text_scale = thumb_scale = 1.0
    for step in range(71):
        s = 1.0 - step * 0.01
        text_scale = max(s, 0.65)
        thumb_scale = 1.0 if s >= 0.65 else s / 0.65
        if _layout_cell(item, attribute_order, cell, text_scale, thumb_scale) <= cell.height:
            break
    _layout_cell(item, attribute_order, cell, text_scale, thumb_scale, draw_page=page)


def _round_rect(page, rect, r, color=None, fill=None, width=0):
    radius = min(0.5, r / max(1, min(rect.width, rect.height)))
    page.draw_rect(rect, color=color, fill=fill, width=width, radius=radius)


def _spec_panel(page, rect, spec, max_size=22):
    """Black rounded panel with the Packing Spec in large white type. Returns its bottom y."""
    inner_w = rect.width - 16
    size = max_size
    lines = _wrap(spec, "hebo", size, inner_w)
    while len(lines) > 2 and size > 11:
        size -= 1
        lines = _wrap(spec, "hebo", size, inner_w)
    lines = lines[:3]
    h = 14 + len(lines) * size * LINE + 6
    panel = fitz.Rect(rect.x0, rect.y0, rect.x1, rect.y0 + h)
    _round_rect(page, panel, 6, fill=(0, 0, 0))
    page.insert_text((panel.x0 + 8, panel.y0 + 11), "PACKING SPEC", fontname="hebo", fontsize=7, color=(1, 1, 1))
    _draw_lines(page, lines, panel.x0 + 8, panel.x0 + 8 + inner_w, panel.y0 + 15, "hebo", size, (1, 1, 1))
    return panel.y1


def _counter_chip(page, x0, y0, text, size=11):
    w = fitz.get_text_length(text, fontname="hebo", fontsize=size) + 16
    h = size * LINE + 8
    chip = fitz.Rect(x0, y0, x0 + w, y0 + h)
    _round_rect(page, chip, h / 2, color=(0, 0, 0), width=1.5)
    page.insert_text((x0 + 8, y0 + 4 + size * 0.92), text, fontname="hebo", fontsize=size)
    return chip


def generate_packing_labels(pack_groups, output_pdf_path="packing_labels.pdf", attribute_order=None):
    doc = fitz.open()
    MM2PT = 2.83465
    MARGIN = 4 * MM2PT
    GAP = 2.5 * MM2PT
    CELL_RADIUS = 3 * MM2PT
    A4_W, A4_H = fitz.paper_size("a4-l")

    ZONE_A_W = 107 * MM2PT
    ZONE_A_H = 150 * MM2PT
    CHIP_H = 11 * LINE + 8
    ROWS = 3

    # One box size for every page: width from page 1's 4 columns beside the courier region,
    # height from the later pages, which lose the footer strip at the bottom
    FIRST_COLS = 4
    CELL_W = (A4_W - 2 * MARGIN - ZONE_A_W - GAP - (FIRST_COLS - 1) * GAP) / FIRST_COLS
    CELL_H = (A4_H - 2 * MARGIN - CHIP_H - GAP - (ROWS - 1) * GAP) / ROWS
    NEXT_COLS = int((A4_W - 2 * MARGIN + GAP) // (CELL_W + GAP))
    FIRST_PAGE_CELLS, NEXT_PAGE_CELLS = FIRST_COLS * ROWS, NEXT_COLS * ROWS

    if not attribute_order:
        attribute_order = ['thumbnail', 'desc', 'dimension', 'job_no', 'qty']

    # "Label X of Y" counts pack groups per store, in the order they appear
    store_totals = {}
    for g in pack_groups.values():
        store_totals[g['store_name']] = store_totals.get(g['store_name'], 0) + 1
    store_seen = {}

    for group_data in pack_groups.values():
        items = group_data['items']
        total_items = len(items)
        store = group_data['store_name']
        spec = group_data.get('pack_spec_name', '')
        store_seen[store] = store_seen.get(store, 0) + 1
        label_text = f"LABEL {store_seen[store]} OF {store_totals[store]}"

        total_pages = 1 if total_items <= FIRST_PAGE_CELLS else 1 + math.ceil((total_items - FIRST_PAGE_CELLS) / NEXT_PAGE_CELLS)
        current_item_idx = 0
        page_num = 1

        while current_item_idx < total_items or (total_items == 0 and page_num == 1):
            page = doc.new_page(width=A4_W, height=A4_H)
            page_text = f"PAGE {page_num} OF {total_pages}"

            if page_num == 1:
                zone_a_rect = fitz.Rect(MARGIN, MARGIN, MARGIN + ZONE_A_W, MARGIN + ZONE_A_H)
                page.draw_rect(zone_a_rect, color=(0.8, 0.8, 0.8), width=0.5, dashes="[3] 0")
                page.insert_textbox(zone_a_rect, "Courier Label Region\n(107mm x 150mm)", fontsize=10, color=(0.6, 0.6, 0.6), align=1)

                zone_b = fitz.Rect(MARGIN, zone_a_rect.y1 + GAP, MARGIN + ZONE_A_W, A4_H - MARGIN)
                y = _spec_panel(page, zone_b, spec) + 8

                if store and store.lower() != 'none':
                    for line in _wrap(store, "hebo", 13, zone_b.width)[:2]:
                        page.insert_text((zone_b.x0, y + 12), line, fontname="hebo", fontsize=13)
                        y += 13 * LINE

                addr_parts = [_clean(group_data.get(k)) for k in ('address_1', 'address_2', 'suburb', 'state', 'postcode', 'country')]
                address = ", ".join(p for p in addr_parts if p)
                if address:
                    address_lines = _wrap(address, "helv", 9, zone_b.width)
                    room = int((zone_b.y1 - CHIP_H - 4 - y) // (9 * LINE))
                    for line in address_lines[:max(0, min(3, room))]:
                        page.insert_text((zone_b.x0, y + 9), line, fontname="helv", fontsize=9)
                        y += 9 * LINE

                chip_y = zone_b.y1 - (11 * LINE + 8)
                chip = _counter_chip(page, zone_b.x0, chip_y, label_text)
                _counter_chip(page, chip.x1 + 6, chip_y, page_text)

                grid_x0 = zone_a_rect.x1 + GAP
                cols = FIRST_COLS
            else:
                footer_y = A4_H - MARGIN - CHIP_H
                x = MARGIN
                if store and store.lower() != 'none':
                    page.insert_text((x, footer_y + 4 + 12 * 0.92), store, fontname="hebo", fontsize=12)
                    x += fitz.get_text_length(store, fontname="hebo", fontsize=12) + 10
                chip = _counter_chip(page, x, footer_y, label_text)
                _counter_chip(page, chip.x1 + 6, footer_y, page_text)

                grid_x0 = MARGIN
                cols = NEXT_COLS


            for r in range(ROWS):
                for c in range(cols):
                    if current_item_idx >= total_items: break
                    cx0 = grid_x0 + c * (CELL_W + GAP)
                    cy0 = MARGIN + r * (CELL_H + GAP)
                    cell = fitz.Rect(cx0, cy0, cx0 + CELL_W, cy0 + CELL_H)
                    _round_rect(page, cell, CELL_RADIUS, color=(0, 0, 0), width=0.75)
                    _draw_cell(page, items[current_item_idx], attribute_order, cell)
                    current_item_idx += 1
            page_num += 1
            if total_items == 0: break

    doc.save(output_pdf_path)
    doc.close()
    return output_pdf_path
