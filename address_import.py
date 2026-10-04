"""Reads the courier portal's CSV of verified addresses and works out which address book entries it corrects.

The portal confirms addresses only after the consignment CSV has been uploaded; it then offers a CSV of the
corrected addresses. This module matches each row of that file back to the entry the app saved at Generate
and proposes the changes, which the user reviews before anything is written.

Matching, most reliable first:
  1. item reference   any cell in the row equal to an item reference in any project's label map
                      ("Label 2 - Brighton Eyewear"), whatever column the portal puts it in. The same
                      reference can be in several jobs: if they point to different addresses, the row's
                      consignment reference (J number) picks the job
  2. receiver name    the address book entry with the same receiver (if only one)
Rows that match nothing are offered as new addresses (unticked by default).
"""
import csv
import glob
import io
import json
import os

from packing_label_generator import _norm, normalize_postcode
from address_match import best_match, best_by_address, address_score

# Only the columns the address book needs are read; everything else in the file is ignored.
# Each field is recognised by keywords in its header (ignoring capitals and punctuation), so other
# layouts work too: "Company", "Customer Name", "Ship To", "Street", "Town/Suburb", "Post Code/Zip"...
# A header scores 3 for an exact name, 2 when it contains a keyword, +1 when it says it's the
# receiver's (receiver, delivery, ship to, consignee). The best-scoring column wins each field.
FIELD_KEYWORDS = {
    'receiver': (['receiver name', 'receiver', 'consignee', 'consignee name', 'company', 'company name', 'business name',
                  'customer', 'customer name', 'delivery name', 'ship to', 'ship to name', 'recipient', 'name'],
                 ['receiver name', 'consignee', 'company', 'business name', 'customer name', 'customer', 'delivery name',
                  'ship to name', 'recipient', 'store name', 'store', 'retailer', 'name']),
    'contact': (['receiver contact name', 'contact', 'contact name', 'attention', 'attn'],
                ['contact', 'attention', 'attn']),
    'line1': (['receiver address line 1', 'address line 1', 'address 1', 'address1', 'address', 'street', 'street address'],
              ['address line 1', 'address 1', 'address1', 'street', 'address']),
    'line2': (['receiver address line 2', 'address line 2', 'address 2', 'address2'],
              ['address line 2', 'address 2', 'address2']),
    'suburb': (['receiver suburb', 'suburb', 'town', 'city', 'locality', 'town suburb', 'suburb town'],
               ['suburb', 'town', 'city', 'locality']),
    'state': (['receiver state', 'state', 'region', 'province'],
              ['state', 'region', 'province']),
    'postcode': (['receiver postcode', 'postcode', 'post code', 'postal code', 'zip', 'zip code', 'postcode zip'],
                 ['postcode', 'post code', 'postal', 'zip']),
    'country': (['receiver country', 'country'], ['country']),
    'authority_to_leave': (['authority to leave flag', 'authority to leave', 'atl'], ['authority to leave', 'atl']),
}
# Headers about someone other than the receiver, or about phone/email/reference, never count
NOT_RECEIVER = ('sender', 'pickup', 'pick up', 'from', 'return', 'shipper', 'origin')
NOT_ADDRESS = ('phone', 'mobile', 'email', 'e mail', 'fax', 'number', 'reference', 'ref', 'code', 'date', 'instruction',
               'value', 'weight', 'qty', 'quantity', 'type')
RECEIVER_HINTS = ('receiver', 'delivery', 'deliver to', 'ship to', 'consignee', 'recipient')


def _score(header, field):
    """How well a normalised header fits a field (0 = not at all)."""
    words = set(header.split())
    if not header or any(w in words or (' ' in w and w in header) for w in NOT_RECEIVER):
        return 0
    exact, keywords = FIELD_KEYWORDS[field]
    blocked = [w for w in NOT_ADDRESS if w in words]
    if field == 'postcode' and blocked == ['code']:
        blocked = []  # "Post Code" is fine
    if blocked and field != 'authority_to_leave':
        return 0
    if field == 'line1' and any(t in header for t in ('2', 'line 2')):
        return 0
    score = 3 if header in exact else 2 if any((' ' in k and k in header) or k in words for k in keywords) else 0
    if score and any(h in header for h in RECEIVER_HINTS):
        score += 1
    return score


ADDRESS_FIELDS = ('receiver', 'contact', 'line1', 'line2', 'suburb', 'state', 'postcode', 'country')
# What a verified address changes on an entry it matched: the address itself, never the receiver or Attn
# (those identify the entry; one address can be saved for several receivers / Attn names)
PLACE_FIELDS = ('line1', 'line2', 'suburb', 'state', 'postcode', 'country')
UPDATE_FIELDS = PLACE_FIELDS + ('authority_to_leave',)
SAME_PLACE = 0.85  # address_score() above which another copy for the same receiver + Attn is this place
YES = {'y', 'yes', 'true', '1'}


class ImportProblem(ValueError):
    """The file can't be used (not a CSV, no address columns...)."""


def _rows_from_csv(data):
    for encoding in ('utf-8-sig', 'cp1252', 'latin-1'):
        try:
            text = data.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    try:
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t")
    except csv.Error:
        dialect = csv.excel
    return [[c.strip() for c in row] for row in csv.reader(io.StringIO(text), dialect)]


def _cell_text(v):
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        return str(int(v))  # 3186.0 -> 3186
    return str(v).strip()


def _sheets_from_xlsx(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        return [[[_cell_text(v) for v in row] for row in ws.iter_rows(values_only=True)] for ws in wb.worksheets]
    finally:
        wb.close()


def read_sheets(data, filename):
    """The uploaded file as a list of sheets, each a list of rows of cell text. CSV is one sheet;
    Excel (.xlsx, or .xls converted through Excel) gives every sheet."""
    ext = os.path.splitext(filename or '')[1].lower()
    if ext not in ('.xlsx', '.xlsm', '.xls'):
        return [_rows_from_csv(data)]
    import tempfile
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, 'upload' + ext)
        with open(path, 'wb') as f:
            f.write(data)
        if ext == '.xls':
            from matrix_engine import convert_legacy_excel_to_xlsx
            path = convert_legacy_excel_to_xlsx(path)
        try:
            return _sheets_from_xlsx(path)
        except Exception as e:
            raise ImportProblem(f"Couldn't read that Excel file: {e}")


def map_columns(rows):
    """Finds the header row (within the first 15) and the column holding each field the address book needs.
    Returns (header row index, {field: column index}); other columns are ignored."""
    best = (None, {}, 0)
    for idx, row in enumerate(rows[:15]):
        headers = [_norm(c) for c in row]
        candidates = sorted(((_score(h, f), f, col) for col, h in enumerate(headers) for f in FIELD_KEYWORDS), reverse=True)
        found, used = {}, set()
        for score, field, col in candidates:
            if score and field not in found and col not in used:
                found[field] = col
                used.add(col)
        total = sum(_score(headers[c], f) for f, c in found.items())
        if 'postcode' in found and ({'receiver', 'line1'} & found.keys()) and total > best[2]:
            best = (idx, found, total)
    if best[0] is None:
        raise ImportProblem("Couldn't find the address columns. The file needs a header row with at least "
                            "a receiver or address column and a postcode column.")
    return best[0], best[1]


def pick_sheet(sheets):
    """The sheet whose headers fit best (most address columns recognised)."""
    best = None
    for rows in sheets:
        try:
            header, columns = map_columns(rows)
        except ImportProblem:
            continue
        if best is None or len(columns) > len(best[2]):
            best = (rows, header, columns)
    if best is None:
        raise ImportProblem("Couldn't find the address columns. The file needs a header row with at least "
                            "a receiver or address column and a postcode column.")
    return best


def load_label_maps(paths):
    """Every project's label map, indexed for matching:
    {item reference: [(consignment reference, destination), ...]}. The same item reference
    ("Label 1 - Brighton Eyewear") can appear in several jobs, so each keeps every destination it had."""
    by_reference = {}
    for path in paths:
        try:
            with open(path, encoding='utf-8') as f:
                data = json.load(f)
        except (OSError, ValueError):
            continue  # an unreadable label map just can't help with matching
        reference = str(data.get('consignment_reference') or '')
        destinations = {c['number']: c['destination'] for c in data.get('consignments', [])}
        for labels in data.get('stores', {}).values():
            for label in labels:
                dest = destinations.get(label.get('consignment'))
                if dest:
                    by_reference.setdefault(label['item_reference'], []).append((reference, dest))
    return by_reference


def all_label_maps(projects_folder):
    return glob.glob(os.path.join(projects_folder, '*', '*.labelmap.json'))


def _row_address(row, columns):
    get = lambda f: row[columns[f]].strip() if f in columns and columns[f] < len(row) else None
    address = {f: get(f) for f in ADDRESS_FIELDS if get(f) is not None}
    if 'authority_to_leave' in columns:
        address['authority_to_leave'] = (get('authority_to_leave') or '').lower() in YES
    if 'postcode' in address:
        address['postcode'] = normalize_postcode(address['postcode'], address.get('country'))
    return address


def _changes(entry, new):
    """Address details the portal changes on a matched entry (receiver and Attn are never changed)."""
    changed = [f for f in PLACE_FIELDS if f in new and new[f] != (entry.get(f) or '')]
    if 'authority_to_leave' in new and bool(new['authority_to_leave']) != bool(entry.get('authority_to_leave')):
        changed.append('authority_to_leave')
    return changed


def build_review(data, filename, label_map_paths, book):
    """Proposed changes for the review screen.

    Each proposal: {key, rows (CSV line numbers), status ('update' | 'same' | 'conflict' | 'new'),
    matched_by, entry (current book entry or None), new (address from the CSV), changed (field names),
    merges (older copies of this place for the same receiver + Attn, folded into the entry),
    siblings (other receivers / Attn names saved at the old address, which get the same correction)}."""
    rows, header, columns = pick_sheet(read_sheets(data, filename))
    by_reference = load_label_maps(label_map_paths)
    proposals, nearby = {}, {}  # nearby: per-postcode short-lists, shared by every row of this file

    for line, row in enumerate(rows[header + 1:], start=header + 2):
        if not any(row):
            continue
        new = _row_address(row, columns)
        if not (new.get('postcode') or new.get('line1')):
            continue

        entry, matched_by, ambiguous, likely = None, '', False, False

        # What the project label maps say this row is (item reference, narrowed by J number if needed)
        candidates = next((by_reference[c] for c in row if c in by_reference), [])
        cells = set(row)
        if len({id(d) for _, d in candidates}) > 1:
            narrowed = [(ref, d) for ref, d in candidates if ref and ref in cells]
            candidates = narrowed or candidates
        by_job = {}
        for _, dest in candidates:
            found = book.find(dest)
            if found:
                by_job[found['id']] = found

        if new.get('receiver'):
            # Receiver name first: the same address can serve many jobs, the receiver is what identifies it
            same_name = book.find_by_receiver(new['receiver'])
            if len(same_name) == 1:
                entry, matched_by = same_name[0], 'receiver name'
            elif same_name:
                exact = book.find({**same_name[0], **new})
                in_job = [e for e in same_name if e['id'] in by_job]
                if exact:
                    entry, matched_by = exact, 'receiver name + address'
                elif len(in_job) == 1:
                    entry, matched_by = in_job[0], 'receiver name + job'
                else:
                    by_address = best_by_address(new, same_name)
                    if by_address:
                        entry, likely = by_address[0], True
                        matched_by = f"receiver name + similar address ({by_address[1]:.0%})"
                    else:
                        ambiguous = True  # several saved addresses for this receiver; don't guess
            else:
                # No entry with this exact name: look for an obvious match among same-postcode entries
                found = best_match(new, book.near(new, cache=nearby))
                if found:
                    entry, likely = found[0], True
                    matched_by = f"likely match: name {found[2]:.0%}, address {found[3]:.0%}"
        elif len(by_job) == 1:
            entry, matched_by = next(iter(by_job.values())), 'item reference'

        key = f"id:{entry['id']}" if entry else f"new:{_norm(new.get('receiver'))}|{_norm(new.get('line1'))}|{new.get('postcode', '')}"
        if key in proposals:
            p = proposals[key]
            p['rows'].append(line)
            same = lambda a, b: {f: a.get(f) for f in ADDRESS_FIELDS} == {f: b.get(f) for f in ADDRESS_FIELDS}
            variant = next((v for v in p['variants'] if same(v['new'], new)), None)
            if variant:
                variant['rows'].append(line)
            else:
                p['variants'].append({'rows': [line], 'new': new})
                p['status'] = 'conflict'  # rows for the same entry disagree: the user must fix the file
            continue
        merges, siblings = [], []
        if entry:
            changed = _changes(entry, new)
            status = 'update' if changed else 'same'
            # Older copies of this place for the same receiver + Attn (St Kilda saved once with state D and
            # once with VIC): folded into the verified entry so the book keeps one
            verified = {**entry, **{f: new[f] for f in PLACE_FIELDS if f in new}}
            merges = [{**e, 'differs': _changes(e, new)} for e in book.same_identity(entry)
                      if address_score(verified, e) >= SAME_PLACE]
            if merges:
                status = 'update'
            if any(f in PLACE_FIELDS for f in changed) or merges:
                # Other receivers / Attn names saved at the old address get the corrected address too
                skip = {entry['id']} | {m['id'] for m in merges}
                siblings = list({e['id']: e for place in [entry] + merges for e in book.at_address(place)
                                 if e['id'] not in skip and _changes(e, new)}.values())
        else:
            changed, status = list(new.keys()), 'new'
            matched_by = 'several saved addresses for this receiver' if ambiguous else 'not in the address book'
        proposals[key] = {'key': key, 'rows': [line], 'status': status, 'matched_by': matched_by, 'ambiguous': ambiguous, 'likely': likely,
                          'entry': entry, 'new': new, 'changed': changed, 'variants': [{'rows': [line], 'new': new}],
                          'merges': merges, 'siblings': siblings}

    for p in proposals.values():  # which fields the conflicting rows disagree on, for highlighting
        p['variant_diff'] = [f for f in ADDRESS_FIELDS + ('authority_to_leave',)
                             if len({str(v['new'].get(f, '')) for v in p['variants']}) > 1]

    order = {'conflict': 0, 'update': 1, 'new': 2, 'same': 3}
    return sorted(proposals.values(), key=lambda p: (order[p['status']], p['rows'][0])), columns


def apply_review(proposals, chosen_keys, book):
    """Writes the ticked proposals, and marks every address the portal confirmed unchanged as verified
    (those need no tick: the portal already agrees with the address book).

    A ticked update changes only the entry's address details (never its receiver or Attn), folds the
    older copies of that place for the same receiver + Attn into it, and gives the corrected address to the
    other receivers / Attn names saved at the old address. Returns (applied, confirmed, added, merged, problems)."""
    applied = confirmed = added = merged = 0
    problems = []
    for p in proposals:
        if p['status'] == 'same' and p['entry']:
            try:
                confirmed += book.mark_verified(p['entry']['id'])
            except Exception as e:
                problems.append(f"CSV row {', '.join(map(str, p['rows']))}: {e}")
            continue
        if p['key'] not in chosen_keys or p['status'] == 'conflict':
            continue
        try:
            if p['entry']:
                fields = {f: p['new'][f] for f in UPDATE_FIELDS if f in p['new']}
                place = {f: fields[f] for f in PLACE_FIELDS if f in fields}
                stale = [m for m in p.get('merges') or [] if m['id'] != p['entry']['id']]
                # Who else is saved at the old address, read before anything changes
                skip = {p['entry']['id']} | {m['id'] for m in stale}
                others = {e['id'] for old in [p['entry']] + stale for e in book.at_address(old) if e['id'] not in skip}
                target = book.apply_verified(p['entry']['id'], fields)
                applied += 1
                for m in stale:
                    merged += book.merge_into(target['id'], m['id'])
                for other_id in others:
                    if place and book.get(other_id):
                        book.apply_verified(other_id, place)
            else:
                created = book.create(p['new'])
                book.apply_verified(created['id'], {})
                added += 1
        except Exception as e:  # one bad row shouldn't stop the rest
            problems.append(f"CSV row {', '.join(map(str, p['rows']))}: {e}")
    return applied, confirmed, added, merged, problems
