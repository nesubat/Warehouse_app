"""Courier consignment CSV and label-map JSON for the Packing Labels tool.

One CSV row per pack (= one packing label = one carton). Packs going to the same delivery address
form one consignment: the first row of a consignment carries the consignment reference, later rows
leave it blank and repeat the exact same address so the courier joins them.
"""
import csv
import json
import os
import re
from collections import Counter
from datetime import date

from packing_label_generator import _address_key, _norm, assign_label_numbers, assign_barcodes, normalize_postcode, STREET_TYPES
from packing_specs import SpecStore, MAX_KG, _num as _round_to

AU_STATES = {'VIC', 'NSW', 'QLD', 'SA', 'WA', 'TAS', 'NT', 'ACT'}
# New Zealand regions: the official ISO 3166-2:NZ codes, plus the codes Toll's portal uses
# (AUC Auckland, CHR Christchurch, WEL Wellington, MOU Mount Maunganui). TAS is also Tasmania,
# so on its own it means Australia.
NZ_REGIONS = {'AUK', 'BOP', 'CAN', 'CIT', 'GIS', 'HKB', 'MBH', 'MWT', 'NSN', 'NTL', 'OTA', 'STL', 'TAS',
              'TKI', 'WGN', 'WKO', 'WTC',
              'AUC', 'CHR', 'WEL', 'MOU'}
STATES = AU_STATES | NZ_REGIONS
_STATE_PATTERN = "|".join(sorted(STATES))
COUNTRIES = {'australia': 'AU', 'au': 'AU', 'aus': 'AU', 'new zealand': 'NZ', 'nz': 'NZ', 'nzl': 'NZ'}


def _country(column_value, state):
    """Country column if given, else NZ for a New Zealand-only region code, else AU."""
    named = COUNTRIES.get(_norm(column_value))
    if named:
        return named
    if _tidy(column_value):
        return _tidy(column_value).upper()
    return 'NZ' if state in NZ_REGIONS - AU_STATES else 'AU'
COMPANY_WORDS = {
    'sign', 'signs', 'signage', 'storage', 'pty', 'ltd', 'group', 'services', 'service', 'print', 'printing',
    'display', 'displays', 'install', 'installs', 'installations', 'solutions', 'co', 'company', 'graphics',
    'media', 'studio', 'industries', 'trading', 'enterprises', 'logistics', 'optical', 'eyewear', 'shop', 'store',
}
LINE_LIMIT = 30  # longer street text is split into Address Line 1 / Line 2 at its last separator

# Carton size, weight and item type come from the Packing Specs page (packing_specs.py)

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

DEFAULT_FIXED = {'who_pays': 'S', 'charge_account': '', 'service_code': ''}

# Courier service codes offered in the preview.
SERVICE_CODES = [
    ('BORDERP', 'BORDER EXPRESS PARCEL'),
    ('BORDERB', 'BORDER EXPRESS BULK'),
    ('DTRLOF', 'DETRACK LOCAL FREIGHT'),
    ('DHLWPXOD', 'DHL DDP EXPRESS PARCELS>30kG'),
    ('DHLWPXUD', 'DHL DDP EXPRESS PARCELS<30KG'),
    ('FXIEOIE', 'FEDEX INTERNATIONAL ECONOMY ECONOMY'),
    ('FXPOIP', 'FEDEX INTERNATIONAL PRIORITY PRIORITY'),
    ('FLASHFW', 'FLASH COURIERS EXP WAGON'),
    ('FLASHVIP', 'FLASH COURIERS VIP CAR'),
    ('FLASHFV', 'FLASH COURIERS EXP VAN'),
    ('FLASHVW', 'FLASH COURIERS VIP WAGON'),
    ('FLASHVV', 'FLASH COURIERS VIP VAN'),
    ('FLASHS', 'FLASH COURIERS STD CAR'),
    ('FLASHSF1', 'FLASH COURIERS STD FLATTOP'),
    ('FLASHSF2', 'FLASH COURIERS STD FLATTOP 2SP'),
    ('FLASHVF1', 'FLASH COURIERS VIP FLATTOP'),
    ('FLASHVF2', 'FLASH COURIERS VIP FLATT2SP'),
    ('FLASHSV', 'FLASH COURIERS STD VAN'),
    ('FLASHFX', 'FLASH COURIERS EXP CAR'),
    ('FLASHSW', 'FLASH COURIERS STD WAGON'),
    ('STEFPP', 'STARTRACK PREMIUM (SATCHEL)'),
    ('STEROAD', 'STARTRACK ROAD EXPRESS'),
    ('STERET', 'STARTRACK ROAD EXPRESS TAILGATE'),
    ('STERE2', 'STARTRACK ROAD EXPRESS 2 MAN SERVICE'),
    ('STEPRM', 'STARTRACK PREMIUM'),
    ('STNPRM', 'STARTRACK NATIONAL PREMIUM'),
    ('STNFPP', 'STARTRACK NATIONAL PREMIUM(SATCHEL)'),
    ('STNROAD', 'STARTRACK NATIONAL ROAD EXPRESS'),
    ('TNTN', 'TNT (NSW) ROAD EXPRESS'),
    ('TNTNONFC', 'TNT (NSW) OVERNIGHT EXPRESS'),
    ('TNTN9AM', 'TNT (NSW) OVERNIGHT 9AM'),
    ('TNTNSWSPEC', 'TNT (NSW) SPECIALISED (TAILGATE)'),
    ('TNTIXP', 'TNT INTERNATIONAL EXPRESS PARCELS'),
    ('IPECX', 'TOLL IPEC ROAD EXPRESS'),
]
SERVICE_CODE_SET = {code for code, _ in SERVICE_CODES}


def load_service_usage(path):
    try:
        with open(path, encoding='utf-8') as f:
            data = json.load(f)
        return {k: int(v) for k, v in data.items() if k in SERVICE_CODE_SET}
    except (OSError, ValueError, AttributeError):
        return {}


def service_options(usage):
    """Service codes for the dropdowns, most used first; unused ones keep the catalogue order."""
    order = {code: i for i, (code, _) in enumerate(SERVICE_CODES)}
    ranked = sorted(SERVICE_CODES, key=lambda cd: (-usage.get(cd[0], 0), order[cd[0]]))
    return [{'code': code, 'description': desc, 'used': usage.get(code, 0)} for code, desc in ranked]


def record_service_usage(path, consignments):
    """After a Generate: count each consignment's service code, so the most used rise to the top."""
    usage = load_service_usage(path)
    for c in consignments:
        code = c['service_code_used']
        if code in SERVICE_CODE_SET:
            usage[code] = usage.get(code, 0) + 1
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(usage, f, indent=2)


# ---------- address parsing ----------

def _tidy(text):
    return re.sub(r'\s+', ' ', str(text or '')).strip(' ,;.-')


def _is_company(name):
    words = re.findall(r"[a-z&]+", name.lower())
    return len(words) > 3 or any(w in COMPANY_WORDS or w == '&' for w in words) or bool(re.search(r'\d', name))


STATE_NAMES = {
    'new south wales': 'NSW', 'victoria': 'VIC', 'queensland': 'QLD', 'south australia': 'SA',
    'western australia': 'WA', 'tasmania': 'TAS', 'northern territory': 'NT', 'australian capital territory': 'ACT',
}
# Words that end a street name ('27 Sirius Road | Lane Cove'), full and short forms
STREET_WORDS = set(STREET_TYPES) | set(STREET_TYPES.values()) | {
    'way', 'close', 'cl', 'circuit', 'cct', 'grove', 'gr', 'square', 'sq', 'walk', 'esplanade', 'esp', 'rise', 'row',
    'track', 'trail', 'av', 'bvd', 'cr', 'crt', 'parkway', 'pkwy', 'loop', 'link', 'gardens', 'gdns', 'mews',
}
# Words that make a part of the cell an address rather than a name ('Westfield Knox', 'Shop 12', 'PO Box 9')
ADDRESS_WORDS = {
    'shop', 'shops', 'unit', 'level', 'lvl', 'suite', 'lot', 'tenancy', 'po', 'gpo', 'locked', 'box', 'building', 'bldg',
    'westfield', 'centre', 'center', 'plaza', 'mall', 'arcade', 'shopping', 'warehouse', 'floor', 'cnr', 'corner',
    'kiosk', 'dfo', 'outlet', 'outlets', 'marketplace', 'complex', 'precinct', 'factory', 'dock',
}
_PHONE = re.compile(r'(?<![\d/])(?:\+?61\s?|\+?64\s?|0)[2-9](?:[\s-]?\d){7,8}(?![\d/])')
_ATTN = re.compile(r'^\(?\s*(?:attn|atnn|attention|att)\b[:.]?\s*(.*?)\s*\)?$', re.I)


def _split_name_address(raw):
    """'Steven Priestley - Wilson Storage, 68 Ricketts Rd, Mount Waverley, Vic, 3149 - Attn Ben ATL'
    -> ('Steven Priestley', 'Wilson Storage, 68 Ricketts Rd, Mount Waverley, Vic, 3149', 'Ben', True)
    Line breaks are kept: they separate parts of the address like commas do."""
    text = re.sub(r'[^\S\n]+', ' ', str(raw or '')).strip()
    atl = bool(re.search(r'\bATL\b', text, re.I))
    text = re.sub(r'\bATL\b', '', text, flags=re.I)
    attn = ''
    m = re.search(r'\(?\b(?:attn|atnn|attention|att)\b[:.]?\s*([^()\n,;|]*?)\s*\)?\s*$', text, re.I)
    if m:
        attn = _tidy(m.group(1))
        text = text[:m.start()]
    text = re.sub(r'[\s\-–,()|;/]+$', '', text)
    head = ''
    m = re.match(r'^([^\d\n,;|]+?)\s+[-–]\s+(.*)$', text, re.S)
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


def _segments(text):
    """Splits a whole address in one cell into its parts. Commas, semicolons, |, line breaks, tabs, ' - ' and
    ' / ' all separate; '4-10', 'Shop 1 - 3' and '3/27' stay whole."""
    text = re.sub(r'(\d)\s+[-–—]\s+(\d)', r'\1-\2', text)
    parts = re.split(r'\s*(?:[,;|\n\r\t]+|\s[-–—]\s|\s/\s)\s*', text)
    return [_tidy(p) for p in parts if _tidy(p)]


def _state_code(text):
    t = _norm(text)
    if t.upper() in STATES:
        return t.upper()
    return STATE_NAMES.get(t, '')


def _ends_with_state(text):
    """('Lane Cove', 'NSW') for 'Lane Cove NSW' / 'Lane Cove New South Wales', else (text, '')."""
    words = text.split()
    for size in (3, 2, 1):
        if len(words) >= size:
            code = _state_code(" ".join(words[-size:]))
            if code:
                return " ".join(words[:-size]), code
    return text, ''


def _split_street_suburb(text):
    """'27 Sirius Road Lane Cove' -> ('27 Sirius Road', 'Lane Cove'): the suburb starts after the last street
    word that follows a name word (so 'Lane Cove' and 'St Kilda' stay suburbs). ('', '') if there's no street."""
    words = text.split()
    cut = None
    for i, w in enumerate(words[:-1]):
        prev = words[i - 1].lower() if i else ''
        if w.lower().strip('.') in STREET_WORDS and i and prev.isalpha() and prev not in STREET_WORDS:
            cut = i
    if cut is None:
        return '', ''
    return " ".join(words[:cut + 1]), " ".join(words[cut + 1:])


def _is_name(part):
    """A part of an address cell that names someone ('Andrew', 'Acme Signs Pty Ltd'), not a place."""
    words = [w.lower().strip('.') for w in part.split()]
    return (bool(words) and not re.search(r'\d', part) and not any(w in ADDRESS_WORDS for w in words)
            and words[-1] not in STREET_WORDS and not _state_code(part) and _norm(part) not in COUNTRIES)


def parse_address_text(raw, names=True):
    """A whole address in one cell -> its parts:
    'Andrew, 27 sirius road, lan cove NSW, 2067' / 'Andrew | 27 Sirius Rd | Lane Cove NSW 2067' /
    'Andrew 27 Sirius Road Lane Cove New South Wales 2067' -> person Andrew, line 27 Sirius Road, suburb Lane Cove,
    state NSW, postcode 2067.

    Read from the end: [country], postcode, state (code or full name), suburb. Then from the front: the parts
    naming someone (only when names=True, i.e. the file has no store/receiver column; a 'Name - address' head is
    always read). Phone numbers are dropped. Returns {head, names, attn, atl, street, suburb, state, postcode, country}."""
    head, text, attn, atl = _split_name_address(raw)
    text = _PHONE.sub(' ', text)
    parts = []
    for part in _segments(text):
        m = _ATTN.match(part)
        if m:
            attn = ", ".join(x for x in (attn, _tidy(m.group(1))) if x)
        else:
            parts.append(part)

    country = postcode = state = suburb = ''
    if len(parts) > 1 and _norm(parts[-1]) in COUNTRIES:
        country = parts.pop()
    # Postcode: a part of its own, or the end of the last part ('Lane Cove NSW 2067')
    if parts:
        m = re.fullmatch(r'(.*?)[\s,]*\b(\d{4})', parts[-1])
        if m:
            postcode = m.group(2)
            parts[-1] = _tidy(m.group(1))
            if not parts[-1]:
                parts.pop()
    # State: a part of its own, or the end of the last part
    if parts and (postcode or len(parts) > 1):
        rest, code = _ends_with_state(parts[-1])
        if code:
            state = code
            parts[-1] = _tidy(rest)
            if not parts[-1]:
                parts.pop()
    # Suburb: what's left of the last part, unless it's the street itself
    if parts and (postcode or state):
        last = parts[-1]
        street, after = _split_street_suburb(last)
        if after:
            parts[-1], suburb = street, after          # '27 Sirius Road Lane Cove'
        elif len(parts) > 1 and not re.match(r'^\d', last) and not street:
            suburb = parts.pop()                       # 'Lane Cove' on its own
        elif len(parts) > 1 and not re.search(r'\d', last):
            suburb = parts.pop()

    # Names at the front
    found = []
    if names:
        while len(parts) > 1 and _is_name(parts[0]):
            found.append(parts.pop(0))
        if not found and parts:
            # 'Andrew 27 Sirius Road': the words before the street number
            m = re.match(r'^([^\d]+?)\s+(\d.*)$', parts[0])
            if m and _is_name(m.group(1)):
                found.append(_tidy(m.group(1)))
                parts[0] = m.group(2)
    elif head and len(parts) > 1 and not re.search(r'\d', parts[0]):
        found.append(parts.pop(0))  # 'Steven Priestley - Wilson Storage, 68 Ricketts Road'
    return {'head': head, 'names': found, 'attn': attn, 'atl': atl, 'street': ", ".join(parts),
            'suburb': suburb, 'state': state, 'postcode': postcode, 'country': country}


def _looks_like_address(text):
    """A whole address ('Andrew, 27 Sirius Rd, Lane Cove NSW 2067') rather than a store name."""
    p = parse_address_text(text)
    return bool(p['postcode'] and (p['state'] or p['street']))


def fill_store_names(pack_groups):
    """For files whose only address information is one combined cell. Run right after reading the file.

    - A store / receiver column that holds whole addresses ('Ship To') with no address column is read as the address.
    - A pack with no store name takes the receiver named in its address cell ('Andrew, 27 Sirius Rd…' -> Andrew),
      or failing that its street, so labels are numbered per receiver and headed with a name."""
    for g in pack_groups.values():
        has_address = any(_tidy(g.get(k)) for k in ('address_1', 'address_2', 'suburb', 'postcode'))
        if not has_address and _looks_like_address(g.get('store_name')):
            g['address_1'], g['store_name'] = g['store_name'], ''
        if _tidy(g.get('store_name')) or not has_address and not _tidy(g.get('address_1')):
            continue
        g['store_from_address'] = True
        dest = resolve_destination(g)
        if dest:
            g['store_name'] = dest['receiver'] or dest['line1'] or dest['suburb']
            g['receiver_unknown'] = not dest['receiver']
    return pack_groups


def resolve_destination(group):
    """Delivery address for a pack group, or None when it has no address at all. The address can be in
    separate columns, all in one cell ('Andrew, 27 Sirius Rd, Lane Cove NSW 2067'), or a mix; columns win."""
    texts = [str(group.get(k) or '').strip() for k in ('address_1', 'address_2')]
    raw = " ".join(_tidy(t) for t in texts if _tidy(t))
    col_suburb, col_state, col_postcode = (_tidy(group.get(k)) for k in ('suburb', 'state', 'postcode'))
    if not raw and not (col_suburb or col_postcode):
        return None
    from_address = group.get('store_from_address')  # the store name was itself taken from this address
    store = '' if from_address else _tidy(group.get('store_name'))
    p = parse_address_text("\n".join(t for t in texts if t), names=not store)

    people = [n for n in ([p['head']] if p['head'] else []) + p['names']]
    company = next((n for n in people if _is_company(n)), '')
    person = next((n for n in people if n != company), '')
    if p['head'] and not company and len(people) > 1:
        company = people[1]
    receiver = company or person or store
    contact = ", ".join(x for x in (person if company else '', p['attn']) if x)

    line1, line2 = _split_lines(p['street'])
    # As always: what the address cell spells out wins over the Suburb / State / Postcode columns
    state = p['state'] or col_state
    state = _state_code(state) or _tidy(state).upper()
    country = _country(_tidy(group.get('country')) or p['country'], state)
    postcode = normalize_postcode(p['postcode'] or col_postcode, country)  # 3149.0 / 803 from numeric cells -> 3149 / 0803
    return {
        'receiver': receiver, 'contact': contact,
        'line1': line1, 'line2': line2,
        'suburb': _tidy(p['suburb']) or col_suburb, 'state': state, 'postcode': postcode, 'country': country,
        'authority_to_leave': p['atl'], 'raw': raw,
    }


def one_line(dest):
    return ", ".join(x for x in (dest['line1'], dest['line2'], dest['suburb'], dest['state'], dest['postcode']) if x)


# ---------- cartons ----------

def carton_weight(override, default):
    """A weight typed in the preview, if it's a usable number, else the spec's weight."""
    try:
        v = float(str(override).strip())
    except (TypeError, ValueError):
        return default
    return _round_to(v, 2) if 0 < v <= MAX_KG else default


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


EDITABLE_FIELDS = ('receiver', 'contact', 'line1', 'line2', 'suburb', 'state', 'postcode')


def _apply_edit(dest, edit):
    """The user's corrections from the preview table, on top of what was read from Excel."""
    if not any(k in edit for k in EDITABLE_FIELDS + ('authority_to_leave',)):
        return dest
    dest = {**dest, **{k: _tidy(edit[k]) for k in EDITABLE_FIELDS if k in edit}}
    dest['state'] = dest['state'].upper()
    if 'state' in edit and dest['country'] in ('AU', 'NZ'):
        dest['country'] = _country('', dest['state'])
    dest['postcode'] = normalize_postcode(dest['postcode'], dest['country'])
    if 'authority_to_leave' in edit:
        dest['authority_to_leave'] = bool(edit['authority_to_leave'])
    return dest


def pack_destination(group):
    """(consignment id, address as read from Excel) for a pack, or None when it has neither an address nor a
    store / receiver name. A pack with a name but no postcode (a file with only receiver names) still gets a
    consignment, so its address can be filled in from the address book or typed in the preview. The id is the
    Excel address, or 'name:' + the receiver when there's no postcode, so it's the same on every refresh."""
    dest = resolve_destination(group)
    store = _tidy(group.get('store_name'))
    if dest is None:
        if not store:
            return None
        dest = {'receiver': store, 'contact': '', 'line1': '', 'line2': '', 'suburb': '', 'state': '', 'postcode': '',
                'country': 'AU', 'authority_to_leave': False, 'raw': ''}
    if dest['postcode']:
        return _address_key(one_line(dest)), dest
    if not dest['receiver']:
        return None
    return f"name:{_norm(dest['receiver'])}|{_address_key(one_line(dest))}", dest


def source_destinations(pack_groups):
    """{consignment id: address as read from Excel} for every pack that can be sent (for the address book)."""
    found = {}
    for g in pack_groups.values():
        source = pack_destination(g)
        if source:
            found.setdefault(*source)
    return found


def build_consignments(pack_groups, edits=None, default_service='', resolve_spec=None, weights=None):
    """Groups packs by delivery address. Returns (consignments, cartons, warnings).

    edits: {consignment id: {address fields..., 'service_code': ...}} from the preview table. The id is the
    address as read from Excel, so it stays the same however often the preview is refreshed. Edits apply before
    grouping, so changing an address to match another consignment's merges them, as the courier would.
    resolve_spec: Packing Spec text -> size/weight/item type (SpecStore.resolver(); built-in defaults if None).
    weights: {pack key: kg} typed in the preview's weight boxes."""
    edits, weights = edits or {}, weights or {}
    resolve_spec = resolve_spec or SpecStore(None).resolver()
    # A pack named after its street (no name in its address cell) takes the receiver the address book or an
    # edit gave it, before labels are numbered, so its label and item reference carry that name
    for g in pack_groups.values():
        if g.get('receiver_unknown'):
            source = pack_destination(g)
            final = _apply_edit(source[1], edits.get(source[0], {})) if source else None
            if final and final['receiver']:
                g['store_name'], g['receiver_unknown'] = final['receiver'], False
    assign_label_numbers(pack_groups)
    assign_barcodes(pack_groups)
    consignments, cartons, warnings = {}, [], []
    skipped, unsized, unread = [], [], []

    for key, g in pack_groups.items():
        store = _tidy(g['store_name'])
        source = pack_destination(g)
        rows = g.get('excel_rows', [])
        if source is None:
            skipped.append(f"(no store) ({_tidy(g['pack_spec_name'])}, row {g.get('excel_rows', ['?'])[0]})")
            continue
        source_id, dest = source
        spec = resolve_spec(g['pack_spec_name'])
        size = spec['size_cm']
        if size is None:
            (unread if spec['kind'] in ('OB', 'CS', 'PALLET', 'FP') else unsized).append(_tidy(g['pack_spec_name']))
        weight = carton_weight(weights.get(key), spec['weight_kg'])
        edit = edits.get(source_id, {})
        final = _apply_edit(dest, edit)
        # Packs still without a postcode stay apart (one consignment per receiver) until an address is given
        addr_key = _address_key(one_line(final)) if final['postcode'] else source_id
        if addr_key not in consignments:
            consignments[addr_key] = {
                'number': len(consignments) + 1, 'id': source_id,
                'destination': final, 'original': dest, 'edited': final is not dest,
                'edit_source': edit.get('source', '') if final is not dest else '',
                'service_code': _tidy(edit.get('service_code')),
                'no_address': not final['postcode'],
                'cartons': [],
            }
        con = consignments[addr_key]
        carton = {
            'pack_key': key, 'source_id': source_id, 'store': store, 'packing_spec': _tidy(g['pack_spec_name']),
            'label_no': g['label_no'], 'label_total': g['label_total'],
            'item_reference': f"Label {g['label_no']} - {store}",
            'install': bool(g.get('install')), 'excel_rows': rows,
            'job_numbers': [i['job_no'] for i in g['items']],
            'barcodes': [i['barcode'] for i in g['items']],
            'size_cm': size, 'weight_kg': weight, 'weight_default': spec['weight_kg'],
            'item_type': spec['item_type'], 'spec_kind': spec['kind'],
            'consignment': con['number'], 'first_in_consignment': not con['cartons'],
        }
        con['cartons'].append(carton)
        cartons.append(carton)

    cons = list(consignments.values())
    for c in cons:
        c['service_code_used'] = c['service_code'] or _tidy(default_service)
    if skipped:
        warnings.append(f"No store / receiver name and no address, left out of the courier CSV: {'; '.join(skipped)}")
    missing = [c for c in cons if c['no_address']]
    if missing:
        warnings.append(f"No address for {len(missing)} consignment{'s' if len(missing) != 1 else ''} "
                        f"({', '.join(f'#{c['number']} {c['destination']['receiver']}' for c in missing)}): pick a saved address "
                        f"or type one in with ✏️. Without one they're left out of the courier CSV.")
    if unsized:
        warnings.append(f"No size saved for packing spec {', '.join(sorted(set(unsized)))}: add it on the Packing Specs page "
                        f"and Update Previews. Until then its dimensions are left blank.")
    if unread:
        warnings.append(f"Couldn't read the size from {', '.join(sorted(set(unread)))} (expected e.g. OB1370170170, "
                        f"FP120170). Dimensions are left blank.")
    nameless = [c for c in cons if not c['destination']['receiver']]
    if nameless:
        warnings.append(f"No receiver name for {', '.join(f'#{c['number']} {one_line(c['destination'])}' for c in nameless)}: "
                        f"add it with ✏️ (the courier needs one).")
    for c in cons:
        d = c['destination']
        if d['postcode'] and not (d['suburb'] and d['state']):
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


def sendable(consignments, cartons):
    """The consignments (and their cartons) that have an address: those still without one are left out of
    the courier CSV, the label map and the address book."""
    keep = [c for c in consignments if not c.get('no_address')]
    numbers = {c['number'] for c in keep}
    return keep, [x for x in cartons if x['consignment'] in numbers]


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
            row[16] = by_number[carton['consignment']]['service_code_used']
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
            'barcodes': carton['barcodes'],
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
            'service_code': c['service_code_used'],
            'edited_in_preview': c['edited'],
            'source_address': c['destination']['raw'],
            'cartons': [x['item_reference'] for x in c['cartons']],
        } for c in consignments],
        'stores': stores,
    }
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
