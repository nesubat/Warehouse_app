"""Finds the address book entry an address *obviously* refers to when the receiver name isn't an exact match.

Two steps, both cheap:
  1. Short-list: entries with the same postcode (indexed lookup, typically a handful even at 50k+ entries).
  2. Score each candidate on receiver name and address, and accept the best only when both agree
     and it clearly beats the runner-up. Anything less is "no match" - never a guess.

Name: shared words, one name containing the other ("Rebel Mt Gravatt" / "Rebel Mt Gravatt (Garden City)"),
      and character similarity for typos.
Address: the numbers must agree (shop/unit/street numbers like 3047B, 4-10), street words match even when
      the courier portal has cut them short ("Jamie" / "Jamieson"), Street = St etc., and the same suburb.
"""
from difflib import SequenceMatcher

from packing_label_generator import _address_key, _norm

NAME_NOISE = {'the', 'pty', 'ltd', 'limited', 'co', 'and', 'inc', 'store', 'shop'}
ADDRESS_NOISE = {'shop', 'level', 'lvl', 'unit', 'suite', 'tenancy', 'cnr', 'corner', 'of', 'and', 'the', 'sc', 's', 'c'}

# Thresholds for an "obvious" match
MIN_NAME = 0.45      # names must share real words (or be near-identical spellings)
MIN_ADDRESS = 0.55   # street numbers and words must largely agree
MIN_TOTAL = 0.70
MIN_LEAD = 0.10      # the best candidate must beat the next by this much


def _words(text, noise):
    return [w for w in _norm(text).split() if w not in noise]


def name_score(a, b):
    """0..1 similarity of two receiver names."""
    ta, tb = set(_words(a, NAME_NOISE)), set(_words(b, NAME_NOISE))
    if not ta or not tb:
        return 0.0
    shared = len(ta & tb)
    jaccard = shared / len(ta | tb)
    smaller = min(len(ta), len(tb))
    # One name containing the other counts, but a single shared word ("Rebel") is weak evidence
    contains = (shared / smaller) * min(1.0, smaller / 2)
    best = max(jaccard, 0.9 * contains)
    if best >= 0.9:
        return best
    # Typos: character similarity, only worked out when the cheap upper bounds say it could pass 0.8
    seq = SequenceMatcher(None, " ".join(sorted(ta)), " ".join(sorted(tb)))
    if seq.real_quick_ratio() > 0.8 and seq.quick_ratio() > 0.8:
        spelling = seq.ratio()
        if spelling > 0.8:
            best = max(best, 0.9 * spelling)
    return best


def _street_parts(addr):
    words = _address_key(" ".join(str(addr.get(k) or '') for k in ('line1', 'line2'))).split()
    numbers = {w for w in words if any(ch.isdigit() for ch in w)}
    text = [w for w in words if w not in numbers and w not in ADDRESS_NOISE and len(w) > 1]
    return numbers, text


def _word_overlap(a, b):
    """Shared street words, where a word cut short by the portal still matches ('jamie' ~ 'jamieson')."""
    if not a or not b:
        return 0.0
    hits = sum(1 for w in a if any(w == x or (len(w) >= 3 and len(x) >= 3 and (w.startswith(x) or x.startswith(w))) for x in b))
    return hits / max(len(a), len(b))


def address_score(a, b):
    """0..1 similarity of two addresses' street lines and suburb (postcode is already equal)."""
    na, wa = _street_parts(a)
    nb, wb = _street_parts(b)
    if na and nb:
        if (na - nb) and (nb - na):
            # Each side has a number the other lacks (Shop 1040 vs Shop 2210 at the same 619 Doncaster Rd):
            # a different place, however similar the words. A number on one side only is fine.
            return 0.0
        numbers = len(na & nb) / len(na | nb)
    else:
        numbers = 0.5  # one side has no numbers: neither confirms nor rules out
    words = _word_overlap(wa, wb)
    suburb = 1.0 if _norm(a.get('suburb')) == _norm(b.get('suburb')) else 0.0
    return 0.45 * numbers + 0.35 * words + 0.20 * suburb


def _attn_score(a, b):
    """0..1 similarity of two Attn names; a short form counts ('Sam' ~ 'Samantha')."""
    a, b = _norm(a), _norm(b)
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    if a.startswith(b) or b.startswith(a):
        return 0.8
    return SequenceMatcher(None, a, b).ratio() * 0.7


def score(address, entry):
    n = name_score(address.get('receiver'), entry.get('receiver'))
    s = address_score(address, entry)
    return 0.5 * n + 0.5 * s, n, s


def best_match(address, candidates):
    """(entry, total, name, address) for the one candidate that obviously matches, else None."""
    ranked = sorted(((score(address, e), e) for e in candidates), key=lambda x: x[0][0], reverse=True)
    ranked = [(s, e) for s, e in ranked if s[1] >= MIN_NAME and s[2] >= MIN_ADDRESS and s[0] >= MIN_TOTAL]
    if not ranked:
        return None
    # One address can be saved for several receivers / Attn names. Those aren't rival places (a verified
    # address reaches all of them), so among entries at the same address take the best name, then the
    # closest Attn; only a candidate at another address has to be clearly beaten.
    place = lambda e: _address_key(" ".join(str(e.get(k) or '') for k in ('line1', 'line2', 'suburb', 'postcode')))
    best = {}
    for s, e in ranked:
        key = (round(s[0], 6), _attn_score(address.get('contact'), e.get('contact')))
        if place(e) not in best or key > best[place(e)][0]:
            best[place(e)] = (key, s, e)
    ranked = sorted(((s, e) for _, s, e in best.values()), key=lambda x: x[0][0], reverse=True)
    if len(ranked) > 1 and ranked[0][0][0] - ranked[1][0][0] < MIN_LEAD:
        return None  # two near-equal candidates at different addresses: don't guess
    (total, n, s), entry = ranked[0]
    return entry, total, n, s


MIN_SUGGEST = 0.35   # weakest saved address still worth showing as a suggestion


def _why(n, s, same_postcode, has_address=True, typed_name=True):
    if not typed_name:
        return "Same address" if s >= 0.9 and same_postcode else "Similar address"
    if not has_address:
        return "Saved for this receiver" if n >= 0.9 else "Similar receiver"
    if s >= 0.9 and same_postcode:
        return "Same receiver and address" if n >= 0.9 else "Same address, another receiver"
    if n >= 0.9:
        return "Same receiver, different address" if s < MIN_ADDRESS else "Same receiver, similar address"
    if s >= MIN_ADDRESS and n >= MIN_NAME:
        return "Similar receiver and address"
    return "Similar address" if s >= MIN_ADDRESS else "Similar receiver"


def _typed_name_score(typed, saved):
    """Credit for a name still being typed: 'eyed' already points at 'Eyedentity (Elsternwick)'."""
    ta, tb = _words(typed, NAME_NOISE), _words(saved, NAME_NOISE)
    if not ta or not tb:
        return 0.0
    hits = sum(1 for w in ta if any(x == w or (len(w) >= 2 and x.startswith(w)) for x in tb))
    return 0.85 * hits / len(ta)


def suggestions(address, candidates, limit=5):
    """Saved addresses most likely to be this one, most promising first, for someone to choose from.
    Nothing is applied automatically, so this is looser than best_match(). Weighted towards the address,
    because one address is often saved under several receivers / Attn names. Works on half-typed text too."""
    postcode = _norm(address.get('postcode'))
    has_address = bool(postcode or address.get('line1') or address.get('suburb'))
    typed_name = bool(_norm(address.get('receiver')))
    typed_street = bool(_norm(address.get('line1')) or _norm(address.get('line2')))
    seen, ranked = set(), []
    for e in candidates:
        if e['id'] in seen:
            continue
        seen.add(e['id'])
        n = max(name_score(address.get('receiver'), e.get('receiver')),
                _typed_name_score(address.get('receiver'), e.get('receiver')))
        same_postcode = _norm(e.get('postcode')) == postcode
        s = address_score(address, e)
        if postcode:  # a postcode not typed yet isn't a different one
            s *= 1.0 if same_postcode else 0.6
            if same_postcode and not typed_street:
                s = max(s, 0.6)  # only the postcode (and maybe suburb) typed so far: everything there is a candidate
        # Nothing typed for the name yet: rank on the address alone
        total = 0.6 * s + 0.4 * n if typed_name else s
        if total >= MIN_SUGGEST and (n >= MIN_NAME or s >= MIN_ADDRESS or not typed_name):
            ranked.append((total, n, s, same_postcode, e))
    ranked.sort(key=lambda x: (x[0], x[4].get('use_count') or 0), reverse=True)
    return [{**e, 'score': round(total, 2), 'why': _why(n, s, same, has_address, typed_name)} for total, n, s, same, e in ranked[:limit]]


def best_by_address(address, candidates):
    """Among entries that already share the receiver name, the one whose address obviously matches."""
    ranked = sorted(((address_score(address, e), e) for e in candidates), key=lambda x: x[0], reverse=True)
    ranked = [(s, e) for s, e in ranked if s >= MIN_ADDRESS + 0.15]
    if not ranked or (len(ranked) > 1 and ranked[0][0] - ranked[1][0] < MIN_LEAD):
        return None
    return ranked[0][1], ranked[0][0]
