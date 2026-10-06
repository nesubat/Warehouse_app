"""Courier Import: a courier CSV straight from an address list, with no packing labels.

The distribution file has one row per store / receiver (its address in columns found by their headers, as in
Packing Labels) and a column per pack group, headed 1, 2, 3 … (or P1, P2 …). A store's cell under a pack column
holds that pack's Packing Spec ('OB600100100', 'CS1070580140', 'Box', '2 x OB…' for two boxes); blank or x / X
means that store doesn't get that pack.

  Store  Attn  Address 1  …  Postcode  Country  |  1             2              |  1  2  3 …  (item quantities)
  Corio  …     Shop K001  …  3214      AU       |  OB600100100   CS1070580140   |  1  10 5

Columns headed by a number that hold quantities (only numbers), or a number heading several columns at once (one
merged cell over the items of a pack), aren't pack columns.

read_allocation() returns pack groups shaped like packing_label_generator.parse_packing_data()'s, one per store
and pack, so the consignment preview, the address book, the service codes and both courier CSV writers work the
same as for Packing Labels. Each pack's Item Reference is T<tab>P<pack>: 'T1P2' is pack 2 of the first selected tab.
"""
import re
from collections import Counter

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import column_index_from_string

from packing_label_generator import (_header_field, _issue, _row_ranges, _norm, _address_key, PackCheckError,
                                     ADDRESS_FIELDS)
from courier_export import job_series, _tidy, attn_text

# The columns the Column mapping can set by hand: (field, label, required). The pack columns are added per file
# as 'pack_<number>' fields (pack_fields()).
IMPORT_FIELDS = (
    ('store_name', 'Store Name', False), ('receiver_name', 'Receiver Name', False), ('contact', 'Attn', False),
    ('address_1', 'Address Line 1', False), ('address_2', 'Address Line 2', False), ('address_3', 'Address Line 3', False),
    ('suburb', 'Suburb', False), ('state', 'State', False), ('postcode', 'Postcode', False), ('country', 'Country', False),
)
FIELD_NAMES = {f for f, _, _ in IMPORT_FIELDS}
PACK_FIELD = re.compile(r'pack_(\d{1,3})$')

# A pack column's header: '1', 1, 1.0, 'P1', 'Pack 1', 'Pack group 1', 'Box 1', '#1'
PACK_HEADER = re.compile(r'^(?:p|pk|pack|packs|pack\s*group|packing\s*group|box|carton|group)?\s*[#.:-]?\s*(\d{1,3})(?:\.0+)?$', re.I)
# A cell that means "no pack for this store"
NO_PACK = {'', 'x', '-', '–', '0', 'n/a', 'na', 'none', 'nil'}
TOTAL_ROW = re.compile(r'^(grand\s+)?totals?\b', re.I)


def import_field(header):
    """Which address-list column a header is: Store Name, Receiver Name, Attn and the address columns as Packing
    Labels reads them (_header_field), plus a third address line."""
    h = " ".join(str(header).strip().lower().replace('_', ' ').split())
    if re.search(r'(line|address|addr)\s*3\b', h):
        return 'address_3'
    field = _header_field(header)
    return field if field in FIELD_NAMES else None


def pack_number(header):
    """The pack number a header names ('1', 'P2', 'Pack 3' -> 1, 2, 3), or None."""
    if header is None or isinstance(header, bool):
        return None
    if isinstance(header, (int, float)):
        return int(header) if float(header).is_integer() and 0 < header < 1000 else None
    m = PACK_HEADER.match(str(header).strip())
    return int(m.group(1)) if m and int(m.group(1)) > 0 else None


def _is_number(value):
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return True
    return bool(re.fullmatch(r'\s*\d+(\.\d+)?\s*', str(value)))


def _merged_spans(ws):
    """{(row, column): (top-left row, column)} for every cell of a merged range, so a merged cell reads its value."""
    spans = {}
    for m in ws.merged_cells.ranges:
        for r in range(m.min_row, m.max_row + 1):
            for c in range(m.min_col, m.max_col + 1):
                spans[(r, c)] = (m.min_row, m.min_col, m.max_col - m.min_col + 1)
    return spans


def _header_cells(ws, header_row):
    return [(i, ws.cell(row=header_row, column=i).value) for i in range(1, ws.max_column + 1)]


def detect_pack_columns(ws, header_row, spans=None):
    """{pack number: column} of the pack columns found by their headers, and the warnings about the ones left out.
    A numbered header counts when it heads one column (not a merged cell over several) whose cells aren't all
    numbers (quantities). Two columns with the same number: the first one is used."""
    spans = _merged_spans(ws) if spans is None else spans
    found, warnings, dropped = {}, [], []
    for col, value in _header_cells(ws, header_row):
        n = pack_number(value)
        if n is None:
            continue
        span = spans.get((header_row, col))
        if span and span[2] > 1:
            continue  # one number over several columns: a pack's items, not its spec
        values = [ws.cell(row=r, column=col).value for r in range(header_row + 1, ws.max_row + 1)]
        filled = [v for v in values if v is not None and str(v).strip()]
        if filled and all(_is_number(v) for v in filled):
            continue  # quantities
        if n in found:
            dropped.append((n, col))
            continue
        found[n] = col
    for n, col in dropped:
        warnings.append(_issue('', f"{get_column_letter(col)}{header_row}",
                               f"Pack {n} heads two columns: {get_column_letter(found[n])} is used, {get_column_letter(col)} isn't",
                               "Pick the right one in the Column mapping, or number the packs differently."))
    return dict(sorted(found.items())), warnings


def detect_import_columns(headers):
    """{field: column number} for the address-list fields, by their header names; the first matching column wins."""
    cols = {}
    for col, value in headers:
        if value is None or not str(value).strip():
            continue
        field = import_field(value)
        if field and field not in cols:
            cols[field] = col
    return cols


def apply_choices(cols, packs, choices):
    """The columns found by their headers with the ones set by hand on top. choices: {field or 'pack_<n>': column
    letter, or '-' for not used}."""
    cols, packs = dict(cols), dict(packs)
    for field, letter in (choices or {}).items():
        letter = str(letter or '').strip().upper()
        m = PACK_FIELD.match(field)
        if not letter or not (field in FIELD_NAMES or m):
            continue
        target, key = (packs, int(m.group(1))) if m else (cols, field)
        if letter == '-':
            target.pop(key, None)
            continue
        try:
            target[key] = column_index_from_string(letter)
        except ValueError:
            pass
    return cols, dict(sorted(packs.items()))


def mapping_fields(packs, choices):
    """The fields the Column mapping lists: the address-list fields, then one 'Pack <n>' per pack column found or
    picked, and one more to add a pack the headers didn't show ('Pack 1', needed, when none was found)."""
    numbers = sorted(set(packs) | {int(m.group(1)) for f in (choices or {}) for m in [PACK_FIELD.match(f)] if m})
    fields = list(IMPORT_FIELDS) + [(f"pack_{n}", f"Pack {n}", False) for n in numbers]
    fields.append((f"pack_{max(numbers, default=0) + 1}", f"Pack {max(numbers, default=0) + 1} (another)" if numbers else "Pack 1",
                   not numbers))
    return fields


def scan_columns(excel_path, header_row, sheet_name=None):
    """(address-list columns found by their headers {field: column}, pack columns {number: column}, warnings),
    before the choices made by hand."""
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    try:
        ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active
        cols = detect_import_columns(_header_cells(ws, header_row))
        packs, warnings = detect_pack_columns(ws, header_row)
        return cols, packs, warnings
    finally:
        wb.close()


def read_allocation(excel_path, header_row, sheet_name=None, columns=None):
    """Pack groups from an address list with pack columns: {pack id: group} as parse_packing_data() gives them
    (store_name, install False, address fields, excel_rows, items []), plus 'pack_no' and 'contact' (the Attn
    column). Returns (pack_groups, last row read, {'packs': {number: column letter}}, warnings).
    columns: the Column mapping set by hand ({field or 'pack_<n>': letter or '-'})."""
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    try:
        ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active
        spans = _merged_spans(ws)
        auto_cols = detect_import_columns(_header_cells(ws, header_row))
        auto_packs, warnings = detect_pack_columns(ws, header_row, spans)
        cols, packs = apply_choices(auto_cols, auto_packs, columns)
        if any(PACK_FIELD.match(f) for f in (columns or {})):
            warnings = [w for w in warnings if not w['text'].startswith('Pack ')]  # set by hand: settled

        address_keys = [k for k in ADDRESS_FIELDS + ('address_3',) if k in cols]
        if 'store_name' not in cols and 'receiver_name' not in cols and not address_keys:
            raise ValueError(f"Could not find a Store Name, Receiver Name or address column in Row {header_row}. "
                             f"Pick them in the Column mapping below.")
        if not packs:
            raise ValueError(f"Could not find the pack columns in Row {header_row}: columns headed 1, 2, 3 … (or P1, P2 …) "
                             f"with each store's Packing Spec in them. Pick them in the Column mapping below.")

        def value(row, col):
            top = spans.get((row, col))
            v = ws.cell(row=top[0], column=top[1]).value if top else ws.cell(row=row, column=col).value
            return '' if v is None else str(v).strip()

        letter = {k: get_column_letter(c) for k, c in cols.items()}
        pack_ref = lambda n: f"{get_column_letter(packs[n])} · Pack {n}"
        store_ref = (f"{letter['store_name']} · Store" if 'store_name' in letter
                     else f"{letter['receiver_name']} · Receiver" if 'receiver_name' in letter else "Store")
        address_ref = (f"{letter[address_keys[0]]}–{letter[address_keys[-1]]} · Address" if len(address_keys) > 1
                       else (f"{letter[address_keys[0]]} · Address" if address_keys else "Address"))

        pack_groups = {}
        no_packs, no_receiver, no_address, numbers_only = [], [], [], {}
        seen = {}  # (store, address) -> first row: a store listed twice
        twice = []
        last_row, blank_run = header_row, 0
        for row in range(header_row + 1, ws.max_row + 1):
            receiver = value(row, cols['receiver_name']) if 'receiver_name' in cols else ''
            store = (value(row, cols['store_name']) if 'store_name' in cols else '') or receiver  # else the receiver's
            parts = {k: value(row, cols[k]) for k in address_keys}
            cells = {n: value(row, c) for n, c in packs.items()}
            if not store and not any(parts.values()) and not any(v.lower() not in NO_PACK for v in cells.values()):
                blank_run += 1
                if blank_run > 200:
                    break  # the end of the list (a sheet formatted far below it)
                continue
            blank_run = 0
            if TOTAL_ROW.match(store) and not any(parts.values()):
                continue
            last_row = row
            wanted = {n: v for n, v in cells.items() if v.lower() not in NO_PACK}
            if not store and not any(parts.values()):
                no_receiver.append(row)
                continue
            if not wanted:
                no_packs.append(row)
                continue
            if address_keys and not any(parts.values()):
                no_address.append(row)
            for n, v in wanted.items():
                if _is_number(v):
                    numbers_only.setdefault(n, []).append(row)
            key = (_norm(store), _address_key(" ".join(parts.values())))
            if key in seen:
                twice.append((store or parts.get('address_1', ''), seen[key], row))
            else:
                seen[key] = row
            contact = attn_text(value(row, cols['contact'])) if 'contact' in cols else ''
            line2 = ", ".join(x for x in (parts.get('address_2', ''), parts.get('address_3', '')) if x)
            for n, spec in wanted.items():
                pack_groups[f"ROW_{row}_P{n}"] = {
                    'pack_spec_name': spec, 'store_name': store, 'receiver_name': receiver, 'install': False,
                    'address_1': parts.get('address_1', ''), 'address_2': line2,
                    **{k: parts.get(k, '') for k in ('suburb', 'state', 'postcode', 'country')},
                    'contact': contact, 'pack_no': n, 'excel_rows': [row], 'items': [],
                }
    finally:
        wb.close()

    if not pack_groups:
        detail = f"Rows {_row_ranges(no_packs)} have no pack (blank or x in every pack column)." if no_packs else ''
        raise PackCheckError([_issue('', '', f"No packs found below row {header_row}", detail)], warnings)
    if no_packs:
        warnings.append(_issue(_row_ranges(no_packs), store_ref, "No pack for this store (blank or x in every pack column): left out"))
    if no_receiver:
        warnings.append(_issue(_row_ranges(no_receiver), store_ref, "Pack spec with no store and no address: left out"))
    if no_address:
        warnings.append(_issue(_row_ranges(no_address), address_ref, "No address"))
    for n, rows in numbers_only.items():
        warnings.append(_issue(_row_ranges(rows), pack_ref(n), "A number, not a Packing Spec",
                               "The pack's cell should hold its spec (e.g. OB600100100), or x when the store doesn't get it."))
    for name, first, row in twice:
        warnings.append(_issue(f"{first}, {row}", store_ref, f"{name} listed twice at the same address",
                               "Their packs go in one consignment."))
    return pack_groups, last_row, {'packs': {n: get_column_letter(c) for n, c in packs.items()}}, warnings


def reference_in_sheets(excel_path, tabs_and_rows, filename=''):
    """(the Consignment Reference, every job seen): the job (J + 6 digits) most job numbers above the header rows
    belong to ('J466307-11' -> 'J466307'), else the one in the file name."""
    found = Counter()
    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    try:
        for tab, header_row in tabs_and_rows:
            if tab not in wb.sheetnames:
                continue
            for row in wb[tab].iter_rows(min_row=1, max_row=max(1, header_row - 1), values_only=True):
                for v in row:
                    for part in re.split(r'[\s,;/]+', str(v or '')):
                        job = job_series(part)
                        if job:
                            found[job] += 1
    finally:
        wb.close()
    if not found:
        job = next((job_series(m) for m in re.findall(r'J\d{6}', str(filename), re.I)), '')
        if job:
            found[job] += 1
    return (found.most_common(1)[0][0] if found else ''), sorted(found)
