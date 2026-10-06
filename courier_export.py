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
        'suburb': (_tidy(p['suburb']) or col_suburb).upper(), 'state': state, 'postcode': postcode, 'country': country,
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

def job_series(job_no):
    """The job a Job Number belongs to, used as the Consignment Reference: strictly J + 6 digits at the start
    ('J477161-17' -> 'J477161'). Anything else ('477161-17', 'J47716-1', 'J4771612') -> ''."""
    m = re.match(r'J\d{6}(?!\d)', str(job_no or '').strip(), re.I)
    return m.group(0).upper() if m else ''


def detect_series(pack_groups):
    """(the Consignment Reference: the job most Job Numbers in the distribution file belong to, every job seen)."""
    found = Counter()
    for g in pack_groups.values():
        for item in g['items']:
            job = job_series(item.get('job_no'))
            if job:
                found[job] += 1
    return found.most_common(1)[0][0] if found else '', sorted(found)


EDITABLE_FIELDS = ('receiver', 'contact', 'line1', 'line2', 'suburb', 'state', 'postcode')


def _apply_edit(dest, edit):
    """The user's corrections from the preview table, on top of what was read from Excel."""
    if not any(k in edit for k in EDITABLE_FIELDS + ('authority_to_leave',)):
        return dest
    dest = {**dest, **{k: _tidy(edit[k]) for k in EDITABLE_FIELDS if k in edit}}
    dest['state'], dest['suburb'] = dest['state'].upper(), dest['suburb'].upper()  # suburbs always in capitals
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


SENDER_COLUMNS = {  # CSV column -> sender field: one sender for the whole file, on every row
    'Sender Name': 'name', 'Sender Address 1': 'line1', 'Sender Address 2': 'line2', 'Sender Town': 'suburb',
    'Sender State': 'state', 'Sender Postcode': 'postcode', 'Sender Country': 'country', 'Sender Contact': 'contact',
    'Sender Phone': 'phone', 'Sender Email': 'email',
}


def write_courier_csv(path, cartons, consignments, reference, fixed, sender=None, item_refs=None):
    """The OpenFreight courier CSV (CSV_HEADER): one row per packing label. item_refs: {item reference: the Open360
    Item Reference ('01T1QWXK Store')}, used as this file's Reference too, so courier labels booked from either
    file carry the job's code for Stitch Labels. A pack of several boxes ('2 x OB170170170') is No Items 2, with
    the weight and cubic of both boxes."""
    from packing_specs import split_count
    by_number = {c['number']: c for c in consignments}
    item_refs = item_refs or {}
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
            boxes = split_count(carton['packing_spec'])[0]
            row[17] = boxes
            row[18] = round(boxes * float(carton['weight_kg'] or 0), 2) if boxes > 1 else carton['weight_kg']
            row[20] = carton['item_type']
            if size:
                row[19] = round(boxes * size[0] * size[1] * size[2] / 1_000_000, 3)
                row[21:24] = [_num(v) for v in size]
            if sender and sender.get('name'):
                for column, field in SENDER_COLUMNS.items():
                    row[CSV_HEADER.index(column)] = sender.get(field, '')
            row[-1] = item_refs.get(carton['item_reference'], carton['item_reference'])
            writer.writerow(row)


# ---------- TIG Open360 bulk upload ----------

# TIG Open360's standard bulk-upload header, exactly as its template has it (including the spaces in front of a few
# names), so the portal recognises every column.
OPEN360_HEADER = (
    "Receiver Name, Receiver Address Line1,Receiver Address Line2,Receiver Suburb,Receiver State,Receiver Postcode,"
    "Receiver Country,Receiver Contact Email,Receiver Contact Phone,Receiver Contact Name,Despatch Date,"
    "Special Instructions,Service,Shipment Reference,Is Receiver Residential,Internal Reference,Cost Centre,"
    "Item Quantity,Item Weight,Item Length,Item Width,Item Height,Item Type,Item Description,Item Reference,"
    "Total Items,Total Weight,Total Volume,Sender Address Line1,Sender Suburb,Sender State,Sender Postcode,"
    " Is Sender Residential,Authority To Leave,Is Dangerous Goods,DG Reference,DG Weight,DG Quantity,DG Type,"
    "DG Package Type,Special Service, Sender Name, Sender Address Line 2, Sender Country, Sender Contact Phone,"
    " Sender Contact Name"
).split(',')


def open360_items(cartons, consignments, tab_of=None, serials=None):
    """The cartons as Open360 rows, in the distribution file's order: [{consignment, destination, service,
    item_reference, packing_spec, weight_kg, size_cm, item_type, tab}]. tab_of: {pack key: tab number}, 1 for the
    first selected tab, counting to the right (every row is tab 1 when it's not given). serials: {pack key: '01T1'},
    the serial printed on its packing label, which starts its Item Reference."""
    by_number = {c['number']: c for c in consignments}
    tab_of, serials = tab_of or {}, serials or {}
    return [{'consignment': x['consignment'], 'destination': by_number[x['consignment']]['destination'],
             'service': by_number[x['consignment']]['service_code_used'], 'item_reference': x['item_reference'],
             'packing_spec': x['packing_spec'], 'weight_kg': x['weight_kg'], 'size_cm': x['size_cm'], 'item_type': x['item_type'],
             'label_no': x['label_no'], 'store': x['store'], 'tab': tab_of.get(x['pack_key'], 1),
             'serial': serials.get(x['pack_key'])}
            for x in cartons]


def unique_codes(count, length=4):
    """count different random codes of capital letters (QWXK). 4 letters, longer when there are so many rows that
    a repeat would get likely (26^4 = 456,976 codes; a code length is kept at least 1,000 times the rows)."""
    import secrets
    import string
    while 26 ** length < 1000 * max(count, 1):
        length += 1
    codes = set()
    while len(codes) < count:
        codes.add("".join(secrets.choice(string.ascii_uppercase) for _ in range(length)))
    return list(codes)


def open360_item_references(items, numbering='row'):
    """The Open360 Item Reference of each row and its unique code: '01T1QWXK Provision Clayton'
    = the label's serial in its tab ('01T1': the count in its tab, 'T' + the tab, T1 being the first selected tab;
    the same serial is printed on the packing label), the job's code, a space, the store. The serial is zero-padded
    within the tab (01..42, 001..150). The job's code is one random code for the whole Generate, so every label of
    the job reads the same code: serial + code ('01T1QWXK') is the label's unique code, what Stitch Labels reads back
    off the courier label to find its packing label (label_stitcher.find_code). A label of another job, or of an
    earlier Generate of the same job, has another code, so it can't land on a packing label of this one.
      numbering='row'       01T1, 02T1 ... 01T2: the count of the label in its tab
      numbering='shipment'  01.1, 01.2 ... 02.1 (the older layout '01.1-QWXK-L-1-Store'): shipment, then its cartons
    Returns [(item reference, code)]."""
    if numbering == 'shipment':
        first, seen = {}, {}
        for x in items:
            first.setdefault(x['consignment'], len(first) + 1)
        most = max([sum(1 for y in items if y['consignment'] == c) for c in first] or [1])
        w_s, w_i = len(str(len(first))), len(str(most))
        prefixes = []
        for x in items:
            seen[x['consignment']] = seen.get(x['consignment'], 0) + 1
            prefixes.append(f"{first[x['consignment']]:0{w_s}d}.{seen[x['consignment']]:0{w_i}d}")
        codes = unique_codes(len(items))
        return [(f"{p}-{code}-L-{x['label_no']}-{x['store']}", code) for p, code, x in zip(prefixes, codes, items)]
    per_tab = Counter(x.get('tab', 1) for x in items)
    seen = Counter()
    job = unique_codes(1)[0]  # one code for the whole Generate
    out = []
    for x in items:
        tab = x.get('tab', 1)
        seen[tab] += 1
        width = max(2, len(str(per_tab[tab])))  # 42 labels in a tab -> 01..42, 150 -> 001..150
        serial = x.get('serial') or f"{seen[tab]:0{width}d}T{tab}"  # the packing label's own serial when given
        out.append((f"{serial}{job} {x['store']}", f"{serial}{job}"))
    return out


def write_open360_csv(path, items, reference, despatch=None, numbering='row', item_refs=None):
    """TIG Open360's standard bulk-upload file, in the format the portal accepted (J477161 - Lux Test 30 - Open360.csv):
    one row per packing label, in the distribution file's order.

      Shipment Reference   the reference chosen in the preview, on every row (the portal joins a receiver's rows)
      Item Reference       '01T1QWXK Provision Clayton': the label's count in its tab, T + the tab (T1 = the first
                           selected tab), a unique code for Stitch Labels, then the store
      Internal Reference   kept but empty
      Item Quantity        boxes on that row: 2 for a Packing Spec '2 X OB170170170', else 1; Total Items the same
      Total Weight/Volume  the whole consignment's (all its rows, counting each box)
      Sender columns       kept but empty: the portal uses the account's sender
    Dates are d/m/yyyy, as Excel saves them; UTF-8 without a BOM, CRLF line endings.
    Returns [(Item Reference, code)] per row."""
    from packing_specs import split_count
    item_refs = item_refs or open360_item_references(items, numbering)  # given when the OpenFreight CSV shares them
    despatch = despatch or date.today()
    yes_no = lambda v: 'Y' if v else 'N'
    count = {id(x): split_count(x['packing_spec'])[0] for x in items}
    totals = {}
    for x in items:
        t = totals.setdefault(x['consignment'], {'weight': 0.0, 'volume': 0.0})
        t['weight'] += count[id(x)] * float(x['weight_kg'] or 0)
        if x['size_cm']:
            t['volume'] += count[id(x)] * x['size_cm'][0] * x['size_cm'][1] * x['size_cm'][2] / 1_000_000
    dim = lambda v: _num(v) if v not in ('', None) else ''
    with open(path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f, lineterminator='\r\n')
        writer.writerow(OPEN360_HEADER)
        for n, x in enumerate(items, start=1):
            d = x['destination']
            size = x['size_cm'] or ('', '', '')
            t = totals[x['consignment']]
            row = {
                'Receiver Name': d['receiver'], 'Receiver Address Line1': d['line1'], 'Receiver Address Line2': d['line2'],
                'Receiver Suburb': d['suburb'], 'Receiver State': d['state'], 'Receiver Postcode': d['postcode'],
                'Receiver Country': d.get('country') or 'AU', 'Receiver Contact Name': d.get('contact', ''),
                'Despatch Date': f"{despatch.day}/{despatch.month}/{despatch.year}", 'Service': x['service'],
                'Shipment Reference': reference, 'Is Receiver Residential': 'N',
                'Item Quantity': count[id(x)], 'Item Weight': x['weight_kg'],
                'Item Length': dim(size[0]), 'Item Width': dim(size[1]), 'Item Height': dim(size[2]),
                'Item Type': x['item_type'], 'Item Description': x['packing_spec'], 'Item Reference': item_refs[n - 1][0],
                'Total Items': count[id(x)], 'Total Weight': _num(round(t['weight'], 2)),
                'Total Volume': round(t['volume'], 3) if t['volume'] else '',
                'Authority To Leave': yes_no(d.get('authority_to_leave')), 'Is Dangerous Goods': 'N',
            }
            writer.writerow([row.get(h.strip(), '') for h in OPEN360_HEADER])
    return item_refs


def project_open360_items(project_dir):
    """Rebuilds an already generated project's Open360 rows from its courier CSV as last saved (address-book fills,
    edits, weights, service codes, a changed reference) and its label map, in the courier CSV's (= the distribution
    file's) order. Boxes whose size wasn't known then are sized now from the Packing Specs page.
    Returns (items, reference)."""
    import glob
    from packing_specs import SpecStore
    map_path = next(iter(glob.glob(os.path.join(project_dir, '*.labelmap.json'))), None)
    if not map_path:
        raise ValueError("This project has no label map (it was made before courier CSVs were added).")
    with open(map_path, encoding='utf-8') as f:
        label_map = json.load(f)
    with open(os.path.join(project_dir, label_map['files']['courier_csv']), encoding='utf-8-sig', newline='') as f:
        rows = list(csv.reader(f))
    if rows and [h.strip() for h in rows[0]] == [h.strip() for h in OPEN360_HEADER]:
        raise ValueError("This project's courier CSV is already in the Open360 format.")
    col = {name: i for i, name in enumerate(rows[0]) if name != 'Reference'}
    ref_col = len(rows[0]) - 1  # the last 'Reference' column holds the item reference
    carton_of = {pk['item_reference']: {**pk, 'store': store} for store, packs in label_map['stores'].items() for pk in packs}
    con_of = {c['number']: c for c in label_map['consignments']}
    resolve = SpecStore(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'packing_specs.json')).resolver()
    items = []
    for r in rows[1:]:
        pk = carton_of.get(r[ref_col])
        if not pk:
            continue
        get = lambda name: r[col[name]] if name in col else ''
        size = tuple(float(get(n)) for n in ('Item Length', 'Item Width', 'Item Height')) if get('Item Length') else None
        if size is None:
            size = resolve(pk.get('packing_spec', ''))['size_cm']
        items.append({'consignment': pk['consignment'], 'destination': dict(con_of[pk['consignment']]['destination']),
                      'service': get('Service Code'), 'item_reference': r[ref_col], 'packing_spec': pk.get('packing_spec', ''),
                      'weight_kg': get('Total Weight'), 'size_cm': size, 'item_type': get('Item Type') or 'Carton',
                      'label_no': pk['label'], 'store': pk['store']})
    # The CSV as last saved wins (it may have been edited after Generate, e.g. a changed reference)
    reference = next((r[1] for r in rows[1:] if len(r) > 1 and r[1]), label_map.get('consignment_reference', ''))
    return items, reference


def write_label_map(path, cartons, consignments, reference, pdf_name, csv_name, page_info, left_out=(), sender=None,
                    open360=None, project=None):
    """Everything needed to match courier labels to packing-label pages before stitching. left_out: cartons
    that have packing labels but aren't in the courier CSV (no complete address), listed so nothing goes missing.
    open360: {'file', 'shipment_reference', 'item_references': {item reference: (Open360 Item Reference, code)}}.
    csv_name: the courier file (the Open360 CSV); csv_row is each label's row in it."""
    open360 = open360 or {}
    o360_refs = open360.get('item_references', {})
    # A pack of several boxes ('2 x OB170170170') has a set of pages per box; each box takes its own courier label,
    # on the first page of its set (all boxes share the one CSV row and code: Item Quantity 2)
    box_pages = lambda key: [box[0] + 1 for box in (page_info.get('copies') or {}).get(key, []) if box]
    stores = {}
    for csv_row, carton in enumerate(cartons, start=2):  # row 1 is the header
        pages = page_info['pages'].get(carton['pack_key'], [])
        dest = next(c['destination'] for c in consignments if c['number'] == carton['consignment'])
        stores.setdefault(carton['store'], []).append({
            'item_reference': carton['item_reference'],
            'open360_item_reference': (o360_refs.get(carton['item_reference']) or (None, None))[0],
            'open360_code': (o360_refs.get(carton['item_reference']) or (None, None))[1],
            'label': carton['label_no'], 'of': carton['label_total'],
            'packing_spec': carton['packing_spec'],
            'pdf_pages': [p + 1 for p in pages],
            'courier_label_page': pages[0] + 1 if pages else None,
            'boxes': max(1, len(box_pages(carton['pack_key']))),
            'box_pages': box_pages(carton['pack_key']) or ([pages[0] + 1] if pages else []),
            'consignment': carton['consignment'],
            'headed_to': {'receiver': dest['receiver'], 'address': one_line(dest), 'installer': carton['install']},
            'csv_row': csv_row,
            'excel_rows': carton['excel_rows'],
            'job_numbers': carton['job_numbers'],
            'barcodes': carton['barcodes'],
        })
    data = {
        'consignment_reference': reference,
        'project': project,
        'sender': sender if sender and sender.get('name') else None,
        'dispatch_date': date.today().isoformat(),
        'files': {'packing_labels_pdf': pdf_name, 'courier_csv': csv_name, 'open360_csv': open360.get('file'),
                  'courier_csvs': open360.get('csv_files') or [csv_name]},
        'open360_shipment_reference': open360.get('shipment_reference'),
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
        'not_in_courier_csv': [{
            'item_reference': x['item_reference'], 'store': x['store'], 'packing_spec': x['packing_spec'],
            'label': x['label_no'], 'of': x['label_total'],
            'pdf_pages': [p + 1 for p in page_info['pages'].get(x['pack_key'], [])],
            'boxes': max(1, len(box_pages(x['pack_key']))),
            'box_pages': box_pages(x['pack_key']) or [p + 1 for p in page_info['pages'].get(x['pack_key'], [])][:1],
            'excel_rows': x['excel_rows'], 'job_numbers': x['job_numbers'],
            'reason': 'no complete delivery address',
        } for x in left_out],
    }
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
