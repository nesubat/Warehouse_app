"""Packing specs: what each Packing Spec in the Excel file means for the courier (size, weight, item type).

Two kinds:
  formula  OB, CS, Pallet and FP codes spell out their size in millimetres (OB1370170170 -> 137 x 17 x 17 cm).
           FP codes usually give only two sizes (FP120170 -> 12 x 17), the third comes from the page (5 cm);
           a third number in the code is used when there is one. Only their default weight and item type
           (and FP's third size) are stored.
  named    everything else (A4 box, P7 jiffy bag, SRA3...): length, width, height and weight entered by
           the user on the Packing Specs page.

Stored in data/packing_specs.json. Specs seen in a file that the page doesn't know yet are added with no
values, so the page shows what still needs filling in.
"""
import json
import os
import re
import threading
import time

from packing_label_generator import _norm

ITEM_TYPES = ('Carton', 'Pallet')
FORMULA_PREFIXES = ('OB', 'CS', 'PALLET', 'FP')
FORMULA_LABELS = {'OB': 'OB', 'CS': 'CS', 'PALLET': 'Pallet', 'FP': 'FP'}
FORMULA_EXAMPLES = {'OB': 'OB1370170170 → 137 × 17 × 17 cm', 'CS': 'CS1701701200 → 17 × 17 × 120 cm',
                    'PALLET': 'Pallet 1200800500 → 120 × 80 × 50 cm', 'FP': 'FP120170 → 12 × 17 × (3rd size) cm'}
DEFAULT_FORMULA = {
    'OB': {'weight': 2, 'item_type': 'Carton'},
    'CS': {'weight': 2, 'item_type': 'Carton'},
    'PALLET': {'weight': 200, 'item_type': 'Pallet'},
    'FP': {'weight': 2, 'item_type': 'Carton', 'third_size': 5},
}
# The specs the business uses; the ones without numbers yet are for the user to fill in
DEFAULT_NAMED = [
    {'name': 'A4 box', 'length': 31, 'width': 22, 'height': 18, 'weight': 2},
    {'name': 'Half A4 box', 'length': None, 'width': None, 'height': None, 'weight': 2},
    {'name': 'A3 box', 'length': None, 'width': None, 'height': None, 'weight': 2},
    {'name': 'Half A3 box', 'length': None, 'width': None, 'height': None, 'weight': 2},
    {'name': 'SRA3', 'length': None, 'width': None, 'height': None, 'weight': 2},
    {'name': 'P7 jiffy bag', 'length': 48, 'width': 36, 'height': 3, 'weight': 1},
    {'name': 'P5 jiffy bag', 'length': None, 'width': None, 'height': None, 'weight': 1},
    {'name': 'P1 jiffy bag', 'length': None, 'width': None, 'height': None, 'weight': 1},
]
FALLBACK_WEIGHT = 2  # a spec with no weight anywhere
MAX_CM, MAX_KG = 1000, 5000

_FORMULA = re.compile(r'^\s*(OB|CS|PALLET|FP)\s*[-:#]?\s*(\d[\d\s.x×*/-]*?)\s*$', re.I)


class SpecError(ValueError):
    """A value entered on the Packing Specs page can't be saved."""


def _splits(digits, count):
    """Every way to cut digits into `count` numbers of 2-4 digits with no leading zero."""
    if count == 1:
        return [(digits,)] if 2 <= len(digits) <= 4 and digits[0] != '0' else []
    out = []
    for size in range(2, 5):
        head, tail = digits[:size], digits[size:]
        if len(head) == size and head[0] != '0':
            out += [(head,) + rest for rest in _splits(tail, count - 1)]
    return out


def _whole_cm(splits):
    return [s for s in splits if all(p.endswith('0') for p in s)]


def split_dimensions(digits, count=3):
    """'1701701200' -> (170, 170, 1200) millimetres; with count=2, '120170' -> (120, 170).

    The digits run together, so the split is worked out: each size is 2-4 digits with no leading zero,
    and sizes are whole centimetres, so every millimetre value ends in 0 (170 170 1200, not 1701 701 200).
    If that doesn't settle it, the way OB codes were always read: the later sizes are 3 digits and the
    length takes the rest (1375 170 170)."""
    splits = _splits(digits, count)
    choice = _whole_cm(splits) or splits
    if not choice:
        return None
    best = max(choice, key=lambda s: (all(len(p) == 3 for p in s[1:]), sum(len(p) == 3 for p in s), len(s[0])))
    return tuple(int(p) for p in best)


_COUNT_X = re.compile(r'^\s*(\d{1,2})\s*([x×*])(\s*)([a-z].*)$', re.I)  # '2 x OB…', '2x OB…', '2 xOB…', '2xOB…'
_COUNT_SPACE = re.compile(r'^\s*(\d{1,2})\s+([a-z].*)$', re.I)        # '2 OB…', '2 A4 Box'
_COUNT_JOINED = re.compile(r'^\s*(\d{1,2})([a-z].*)$', re.I)          # '2OB170170170'


def split_count(spec):
    """(boxes, spec): a box count written in front of the Packing Spec means that many boxes, each with its own
    packing label. '2 X OB170170170', '2x OB170170170', '2 xOB170170170', '2 OB170170170' -> (2, 'OB170170170').
    A count run straight into the spec ('2OB170170170') counts only when the rest is a formula code (OB, CS,
    PALLET, FP), so a name like '3D Sign' stays one box. With an x, the spec after it must start a new word or be a
    formula code, so '2 XL Box' is 2 of 'XL Box', not 2 of 'L Box'. The spec must start with a letter, so
    '420 x 296' stays a size. Anything else -> (1, spec)."""
    text = str(spec or '').strip()
    m = _COUNT_X.match(text)
    if m and int(m.group(1)) > 0 and (m.group(3) or m.group(2) in '×*' or parse_formula(m.group(4))):
        return int(m.group(1)), m.group(4).strip()
    m = _COUNT_SPACE.match(text)
    if m and int(m.group(1)) > 0:
        return int(m.group(1)), m.group(2).strip()
    m = _COUNT_JOINED.match(text)
    if m and int(m.group(1)) > 0 and parse_formula(m.group(2)):
        return int(m.group(1)), m.group(2).strip()
    return 1, text


def parse_formula(spec, fp_third_size=None):
    """('OB' | 'CS' | 'PALLET' | 'FP', (length, width, height) in cm or None) for formula codes, else None.
    Numbers can run together (OB1701701200) or be separated (OB 170 x 170 x 1200). An FP code with only
    two sizes (FP120170) gets fp_third_size (cm) as its height."""
    m = _FORMULA.match(str(spec or ''))
    if not m:
        return None
    prefix, rest = m.group(1).upper(), m.group(2)
    groups = re.findall(r'\d+', rest)
    mm = None
    if len(groups) in (2, 3):
        mm = tuple(int(g) for g in groups)
    elif len(groups) == 1:
        count = 3
        if prefix == 'FP':
            # Usually just two sizes; three only when only a three-way split gives whole centimetres
            two, three = _whole_cm(_splits(groups[0], 2)), _whole_cm(_splits(groups[0], 3))
            count = 3 if three and not two else 2
        mm = split_dimensions(groups[0], count)
    if mm and len(mm) == 2:
        if prefix != 'FP' or not fp_third_size:
            return prefix, None
        return prefix, tuple(_num(v / 10) for v in mm) + (_num(float(fp_third_size)),)
    if not mm or len(mm) != 3 or not all(mm):
        return prefix, None
    return prefix, tuple(_num(v / 10) for v in mm)


def _num(v, places=1):
    """Sizes to 0.1 cm, weights (places=2) to 0.01 kg; whole numbers without a decimal point."""
    v = round(float(v), places)
    return int(v) if v.is_integer() else v


def _number(value, label, limit, places=1):
    text = str(value if value is not None else '').strip()
    if text == '':
        return None
    try:
        v = float(text)
    except ValueError:
        raise SpecError(f"{label} must be a number.")
    if not 0 < v <= limit:
        raise SpecError(f"{label} must be more than 0 and at most {limit}.")
    return _num(v, places)


class SpecStore:
    def __init__(self, path):
        self.path = path  # None: built-in defaults only, nothing saved
        self._lock = threading.Lock()
        if path:
            os.makedirs(os.path.dirname(path) or '.', exist_ok=True)

    # ---------- storage ----------

    def load(self):
        try:
            with open(self.path, encoding='utf-8') as f:
                data = json.load(f)
        except (OSError, TypeError, ValueError):
            data = {}
        formula = {p: {**DEFAULT_FORMULA[p], **(data.get('formula') or {}).get(p, {})} for p in FORMULA_PREFIXES}
        named = data.get('named')
        if not isinstance(named, list):
            named = [dict(s) for s in DEFAULT_NAMED]
        return {'formula': formula, 'named': named}

    def _write(self, data):
        tmp = f"{self.path}.tmp"
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        os.replace(tmp, self.path)  # never a half-written file

    # ---------- the page ----------

    def save_page(self, formula_rows, named_rows):
        """Replaces everything with what the page sent. Returns {row id: error} and saves nothing on errors.
        formula_rows: {prefix: {weight, item_type}}; named_rows: [{name, length, width, height, weight,
        item_type, remove}]."""
        errors, formula, named, seen = {}, {}, [], {}
        for prefix in FORMULA_PREFIXES:
            row = formula_rows.get(prefix, {})
            try:
                weight = _number(row.get('weight'), 'Weight', MAX_KG, places=2)
                if weight is None:
                    raise SpecError("Weight is needed.")
                formula[prefix] = {'weight': weight, 'item_type': row.get('item_type') if row.get('item_type') in ITEM_TYPES
                                   else DEFAULT_FORMULA[prefix]['item_type']}
                if prefix == 'FP':
                    third = _number(row.get('third_size'), '3rd size', MAX_CM)
                    if third is None:
                        raise SpecError("3rd size is needed (used when an FP code gives only two sizes).")
                    formula[prefix]['third_size'] = third
            except SpecError as e:
                errors[f"formula-{prefix}"] = str(e)
        for i, row in enumerate(named_rows):
            name = " ".join(str(row.get('name') or '').split())
            if row.get('remove') or not (name or any(str(row.get(k) or '').strip() for k in ('length', 'width', 'height', 'weight'))):
                continue
            try:
                if not name:
                    raise SpecError("Give the spec a name, exactly as it's written in the Packing Spec column.")
                if parse_formula(name):
                    raise SpecError(f"{name} is an {parse_formula(name)[0]} code: its size comes from the numbers in it.")
                key = _norm(name)
                if key in seen:
                    raise SpecError(f"{name} is listed twice.")
                dims = [_number(row.get(k), k.title(), MAX_CM) for k in ('length', 'width', 'height')]
                if any(d is not None for d in dims) and not all(d is not None for d in dims):
                    raise SpecError("Fill in all of length, width and height, or none of them.")
                seen[key] = True
                named.append({'name': name, 'length': dims[0], 'width': dims[1], 'height': dims[2],
                              'weight': _number(row.get('weight'), 'Weight', MAX_KG, places=2),
                              'item_type': row.get('item_type') if row.get('item_type') in ITEM_TYPES else 'Carton'})
            except SpecError as e:
                errors[f"named-{i}"] = str(e)
        if not errors:
            with self._lock:
                self._write({'formula': formula, 'named': named, 'updated_at': time.time()})
        return errors

    def note_unknown(self, names):
        """Adds specs seen in a file that the page doesn't know, with no values, for the user to fill in."""
        with self._lock:
            data = self.load()
            known = {_norm(s['name']) for s in data['named']}
            new = []
            for name in names:
                name = " ".join(str(name or '').split())
                if name and _norm(name) not in known and not parse_formula(name):
                    known.add(_norm(name))
                    new.append({'name': name, 'length': None, 'width': None, 'height': None, 'weight': None,
                                'item_type': 'Carton', 'seen_at': time.time()})
            if new:
                data['named'] = new + data['named']
                self._write(data)
            return [s['name'] for s in new]

    # ---------- resolving a Packing Spec ----------

    def resolver(self):
        """A function spec text -> {'size_cm', 'weight_kg', 'item_type', 'kind'}, using one snapshot of the store."""
        data = self.load()
        named = {_norm(s['name']): s for s in data['named']}

        def resolve(spec):
            count, spec = split_count(spec)  # '2 X OB170170170': the size and weight are those of one OB170170170
            return {**_resolve(spec), 'count': count}

        def _resolve(spec):
            formula = parse_formula(spec, data['formula']['FP']['third_size'])
            if formula:
                prefix, size = formula
                f = data['formula'][prefix]
                return {'size_cm': size, 'weight_kg': f['weight'], 'item_type': f['item_type'], 'kind': prefix}
            s = named.get(_norm(spec))
            if s:
                size = (s['length'], s['width'], s['height']) if s.get('length') else None
                return {'size_cm': size, 'weight_kg': s.get('weight') or FALLBACK_WEIGHT,
                        'item_type': s.get('item_type') or 'Carton', 'kind': 'named' if size else 'incomplete'}
            return {'size_cm': None, 'weight_kg': FALLBACK_WEIGHT, 'item_type': 'Carton', 'kind': 'unknown'}
        return resolve
