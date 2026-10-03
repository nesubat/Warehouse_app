"""Courier consignment CSV and label-map JSON for the Packing Labels tool.

One CSV row per pack (= one packing label = one carton). Packs going to the same delivery address
form one consignment: the first row of a consignment carries the consignment reference, later rows
leave it blank and repeat the exact same address so the courier joins them.
"""
import csv
import json
import re
from collections import Counter
from datetime import date

from packing_label_generator import _address_key, _norm, assign_label_numbers

STATES = {'VIC', 'NSW', 'QLD', 'SA', 'WA', 'TAS', 'NT', 'ACT'}
COMPANY_WORDS = {
    'sign', 'signs', 'signage', 'storage', 'pty', 'ltd', 'group', 'services', 'service', 'print', 'printing',
    'display', 'displays', 'install', 'installs', 'installations', 'solutions', 'co', 'company', 'graphics',
    'media', 'studio', 'industries', 'trading', 'enterprises', 'logistics', 'optical', 'eyewear', 'shop', 'store',
}
LINE_LIMIT = 30  # longer street text is split into Address Line 1 / Line 2 at its last separator

# Sizes (cm) for packing specs that don't spell out their dimensions like OB1370170170
KNOWN_SIZES = {
    'p7 jiffy bag': (48, 36, 3),
    'a4 box': (31, 22, 18),
}
LIGHT_SPEC_PREFIXES = ('P1', 'P5', 'P7')  # bags: 1 kg; everything else 2 kg

CSV_HEADER = [
    'Dispatch Date', 'Reference', 'Receiver Name', 'Receiver Address Line 1', 'Receiver Address Line 2',
    'Receiver Suburb', 'Receiver State', 'Receiver Postcode', 'Receiver Country', 'Receiver Contact Name',
    'Receiver Phone', 'Receiver Email', 'Authority To Leave Flag', 'Special Instructions', 'Who Pays',
    'Charge Account (3rd Party)', 'Service Code', 'No Items', 'Total Weight', 'Total Cubic', 'Item Type',
    'Item Length', 'Item Width', 'Item Height', 'Item EDN', 'Commercial Value (0/1)', 'Export Desc',
    'Export Origin', 'Contents Desc', 'Contents Qty', 'Contents $AUD', 'Contents Weight', 'Tariff Code',
    'Sender Name', 'Sender Address 1', 'Sender Address 2', 'Sender Town', 'Sender State', 'Sender Postcode',
    'Sender Country', 'Sender Contact', 'Sender Phone', 'Sender Email', 'Reference',
]

DEFAULT_FIXED = {'who_pays': 'S', 'charge_account': '', 'service_code': 'IPECX'}


# ---------- address parsing ----------

def _tidy(text):
    return re.sub(r'\s+', ' ', str(text or '')).strip(' ,;.-')


def _is_company(name):
    words = re.findall(r"[a-z&]+", name.lower())
    return len(words) > 3 or any(w in COMPANY_WORDS or w == '&' for w in words) or bool(re.search(r'\d', name))


def _split_name_address(raw):
    """'Steven Priestley - Wilson Storage, 68 Ricketts Rd, Mount Waverley, Vic, 3149 - Attn Ben ATL'
    -> ('Steven Priestley', 'Wilson Storage, 68 Ricketts Rd, Mount Waverley, Vic, 3149', 'Ben', True)"""
    text = re.sub(r'\s+', ' ', str(raw or '')).strip()
    atl = bool(re.search(r'\bATL\b', text, re.I))
    text = re.sub(r'\bATL\b', '', text, flags=re.I)
    attn = ''
    m = re.search(r'\(?\b(?:attn|atnn|attention|att)\b[:.]?\s*([^()]*?)\s*\)?\s*$', text, re.I)
    if m:
        attn = _tidy(m.group(1))
        text = text[:m.start()]
    text = re.sub(r'[\s\-–,()]+$', '', text)
    head = ''
    m = re.match(r'^([^\d]+?)\s+[-–]\s+(.*)$', text)
    if m:
        head, text = _tidy(m.group(1)), m.group(2)
    return head, text, attn, atl


def _split_lines(street):
    street = _tidy(street)
    if len(street) <= LINE_LIMIT:
        return street, ''
    cuts = [m.end() for m in re.finditer(r'[,;]\s|\.\s', street)]
    if not cuts:
        return street, ''
    cut = cuts[-1]
    return _tidy(street[:cut]), _tidy(street[cut:])


def resolve_destination(group):
    """Delivery address for a pack group, or None when it has no address at all."""
    raw = " ".join(_tidy(group.get(k)) for k in ('address_1', 'address_2') if _tidy(group.get(k)))
    col_suburb, col_state, col_postcode = (_tidy(group.get(k)) for k in ('suburb', 'state', 'postcode'))
    if not raw and not (col_suburb or col_postcode):
        return None

    head, text, attn, atl = _split_name_address(raw)
    tokens = [t.strip() for t in text.split(',') if t.strip()]
    suburb, state, postcode = col_suburb, col_state, col_postcode

    # Whole address typed into one cell: take suburb/state/postcode from its tail
    if tokens:
        m = re.fullmatch(r'(?i)(?:(.*?)\s+)?(VIC|NSW|QLD|SA|WA|TAS|NT|ACT)\s+(\d{4})', tokens[-1])
        if m:
            tokens.pop()
            postcode, state = m.group(3), m.group(2)
            suburb = m.group(1) or (tokens.pop() if len(tokens) > 1 else suburb)
        elif re.fullmatch(r'\d{4}', tokens[-1]):
            postcode = tokens.pop()
            if tokens and tokens[-1].upper() in STATES:
                state = tokens.pop()
            if len(tokens) > 1:
                suburb = tokens.pop()

    company = ''
    if head and len(tokens) > 1 and not re.search(r'\d', tokens[0]):
        company = tokens.pop(0)  # 'Steven Priestley - Wilson Storage, 68 Ricketts Road'

    person = ''
    if head:
        if _is_company(head) and not company:
            company = head
        elif not _is_company(head):
            person = head
    receiver = company or person or _tidy(group.get('store_name'))
    contact = ", ".join(x for x in (person if company else '', attn) if x)

    line1, line2 = _split_lines(", ".join(tokens))
    postcode = re.sub(r'\.0$', '', str(postcode))  # 3149.0 from numeric cells
    return {
        'receiver': receiver, 'contact': contact,
        'line1': line1, 'line2': line2,
        'suburb': _tidy(suburb), 'state': _tidy(state).upper(), 'postcode': postcode, 'country': 'AU',
        'authority_to_leave': atl, 'raw': raw,
    }


def one_line(dest):
    return ", ".join(x for x in (dest['line1'], dest['line2'], dest['suburb'], dest['state'], dest['postcode']) if x)


# ---------- cartons ----------

def package_size(spec):
    """(length, width, height) in cm, or None when unknown."""
    s = re.sub(r'\s+', '', str(spec)).upper()
    m = re.fullmatch(r'OB(\d+)(\d{3})(\d{3})', s)
    if m:
        return tuple(int(v) / 10 for v in m.groups())
    return KNOWN_SIZES.get(_norm(spec))


def item_type(spec):
    return 'Pallet' if str(spec).strip().lower().startswith('pallet') else 'Carton'


def package_weight(spec):
    return 1 if str(spec).strip().upper().startswith(LIGHT_SPEC_PREFIXES) else 2


def _num(v):
    return int(v) if float(v).is_integer() else round(v, 1)


# ---------- consignments ----------

def detect_series(pack_groups):
    found = Counter()
    for g in pack_groups.values():
        for item in g['items']:
            m = re.match(r'\s*(J\d+)', str(item.get('job_no', '')), re.I)
            if m:
                found[m.group(1).upper()] += 1
    return found.most_common(1)[0][0] if found else '', sorted(found)


def build_consignments(pack_groups):
    """Groups packs by delivery address. Returns (consignments, cartons, warnings)."""
    assign_label_numbers(pack_groups)
    consignments, cartons, warnings = {}, [], []
    skipped, unsized = [], []

    for key, g in pack_groups.items():
        store = _tidy(g['store_name'])
        dest = resolve_destination(g)
        rows = g.get('excel_rows', [])
        if dest is None or not dest['postcode']:
            skipped.append(f"{store or '(no store)'} ({_tidy(g['pack_spec_name'])}, row {g.get('excel_rows', ['?'])[0]})")
            continue
        size = package_size(g['pack_spec_name'])
        if size is None:
            unsized.append(_tidy(g['pack_spec_name']))
        addr_key = _address_key(one_line(dest))
        if addr_key not in consignments:
            consignments[addr_key] = {'number': len(consignments) + 1, 'destination': dest, 'cartons': []}
        con = consignments[addr_key]
        carton = {
            'pack_key': key, 'store': store, 'packing_spec': _tidy(g['pack_spec_name']),
            'label_no': g['label_no'], 'label_total': g['label_total'],
            'item_reference': f"Label {g['label_no']} - {store}",
            'install': bool(g.get('install')), 'excel_rows': rows,
            'job_numbers': [i['job_no'] for i in g['items']],
            'size_cm': size, 'weight_kg': package_weight(g['pack_spec_name']),
            'item_type': item_type(g['pack_spec_name']),
            'consignment': con['number'], 'first_in_consignment': not con['cartons'],
        }
        con['cartons'].append(carton)
        cartons.append(carton)

    cons = list(consignments.values())
    if skipped:
        warnings.append(f"No usable address (no postcode), left out of the courier CSV: {'; '.join(skipped)}")
    if unsized:
        warnings.append(f"Unknown carton size for packing spec {', '.join(sorted(set(unsized)))}. Dimensions are left blank.")
    for c in cons:
        d = c['destination']
        if not (d['suburb'] and d['state'] and d['postcode']):
            warnings.append(f"Consignment {c['number']} ({d['receiver']}): incomplete address '{one_line(d)}'.")

    # Possible duplicates the tool shouldn't merge on its own
    by_receiver, by_street = {}, {}
    for c in cons:
        d = c['destination']
        by_receiver.setdefault(_norm(d['receiver']), []).append(c)
        by_street.setdefault(_address_key(f"{d['line1']} {d['line2']}"), []).append(c)
    for group in by_receiver.values():
        if len(group) > 1 and any(c['cartons'][0]['install'] for c in group):
            warnings.append(f"{group[0]['destination']['receiver']} has {len(group)} different addresses: "
                            + " · ".join(f"#{c['number']} {one_line(c['destination'])}" for c in group))
    for street, group in by_street.items():
        if street and len(group) > 1:
            warnings.append(f"Same street, different suburb/postcode: "
                            + " · ".join(f"#{c['number']} {one_line(c['destination'])}" for c in group))
    return cons, cartons, warnings


# ---------- output ----------

def dispatch_date(today=None):
    today = today or date.today()
    return f"{today.day}-{today.strftime('%b')}"


def write_courier_csv(path, cartons, consignments, reference, fixed):
    by_number = {c['number']: c for c in consignments}
    with open(path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(CSV_HEADER)
        for carton in cartons:
            d = by_number[carton['consignment']]['destination']
            size = carton['size_cm']
            row = [''] * len(CSV_HEADER)
            row[0] = dispatch_date()
            row[1] = reference if carton['first_in_consignment'] else ''
            row[2:10] = [d['receiver'], d['line1'], d['line2'], d['suburb'], d['state'], d['postcode'], d['country'], d['contact']]
            row[12] = 'Y' if d['authority_to_leave'] else ''
            row[14] = fixed['who_pays']
            row[15] = fixed['charge_account']
            row[16] = fixed['service_code']
            row[17] = 1
            row[18] = carton['weight_kg']
            row[20] = carton['item_type']
            if size:
                row[19] = round(size[0] * size[1] * size[2] / 1_000_000, 3)
                row[21:24] = [_num(v) for v in size]
            row[-1] = carton['item_reference']
            writer.writerow(row)


def write_label_map(path, cartons, consignments, reference, pdf_name, csv_name, page_info):
    """Everything needed to match courier labels to packing-label pages before stitching."""
    stores = {}
    for csv_row, carton in enumerate(cartons, start=2):  # row 1 is the header
        pages = page_info['pages'].get(carton['pack_key'], [])
        dest = next(c['destination'] for c in consignments if c['number'] == carton['consignment'])
        stores.setdefault(carton['store'], []).append({
            'item_reference': carton['item_reference'],
            'label': carton['label_no'], 'of': carton['label_total'],
            'packing_spec': carton['packing_spec'],
            'pdf_pages': [p + 1 for p in pages],
            'courier_label_page': pages[0] + 1 if pages else None,
            'consignment': carton['consignment'],
            'headed_to': {'receiver': dest['receiver'], 'address': one_line(dest), 'installer': carton['install']},
            'csv_row': csv_row,
            'excel_rows': carton['excel_rows'],
            'job_numbers': carton['job_numbers'],
        })
    data = {
        'consignment_reference': reference,
        'dispatch_date': date.today().isoformat(),
        'files': {'packing_labels_pdf': pdf_name, 'courier_csv': csv_name},
        'page_size_mm': page_info['page_size_mm'],
        'courier_label_region_mm': page_info['courier_region_mm'],
        'consignments': [{
            'number': c['number'],
            'destination': {k: v for k, v in c['destination'].items() if k != 'raw'},
            'source_address': c['destination']['raw'],
            'cartons': [x['item_reference'] for x in c['cartons']],
        } for c in consignments],
        'stores': stores,
    }
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
