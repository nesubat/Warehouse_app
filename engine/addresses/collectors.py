"""Who collects packs from the warehouse, and the colour their labels' courier region is filled with.

A name keeps its colour for good, across projects: the packers learn "blue is Gessie's", so it must never change under
them. Three colours are the same on every computer and belong to no one else (STANDARD): Kyle red on yellow (the
first collect labels), Gessie blue, and any name with 'client' in it green. Everyone else gets the next colour of
PALETTE that no saved name has yet, on this computer (names and colours aren't shared between computers).

The names and colours are kept in the address book (AddressBook.collectors()); this file only decides colours, so the
preview's script (script.js section 30) can show the same colour for a name typed before it's saved: it gets this
palette and the saved names and follows assign() exactly.
"""
import colorsys
import re

from engine.packing_labels.packing_label_generator import _norm

KYLE = 'kyle'
KYLE_COLOUR = {'fill': '#FFEB40', 'ink': '#CC0000'}  # red on yellow, as collect labels have always been
GESSIE = 'gessie'
GESSIE_COLOUR = {'fill': '#8FD3FF', 'ink': '#003A63'}  # blue
CLIENT_COLOUR = {'fill': '#9BE89B', 'ink': '#0B4D0B'}  # green: the client collects ('Client', 'Client - Bob')
STANDARD = {KYLE: KYLE_COLOUR, GESSIE: GESSIE_COLOUR}
RESERVED = {c['fill'] for c in (KYLE_COLOUR, GESSIE_COLOUR, CLIENT_COLOUR)}
MAX_NAME = 40
HEX = re.compile(r'^#[0-9A-F]{6}$')

# Light fills a packer tells apart at a glance across the floor, each with a dark ink of its own hue for the text.
# None is close to the standard yellow, blue or green.
_BASE = (
    ('#FFB3D1', '#7A0036'),  # pink
    ('#FFC285', '#6B2E00'),  # orange
    ('#C9B3FF', '#2E1170'),  # lavender
    ('#7FE0D4', '#004D45'),  # teal
    ('#FF9C9C', '#6E0000'),  # salmon
    ('#B0BEC5', '#1C2B33'),  # grey
    ('#E0B080', '#4A2C00'),  # tan
    ('#C9A0FF', '#3A0A66'),  # violet
    ('#F4A6F0', '#5E0059'),  # orchid
    ('#FFD6A5', '#5C3300'),  # peach
)


def _more(count):
    """Colours after the named ones: hues spread by the golden angle, skipping the yellows (Kyle's)."""
    out, n = [], 0
    while len(out) < count:
        n += 1
        hue = (n * 137.508) % 360
        if 35 <= hue <= 75 or 85 <= hue <= 150 or 185 <= hue <= 215:  # not near the standard yellow, green, blue
            continue
        fill = colorsys.hls_to_rgb(hue / 360, 0.80, 0.70)
        ink = colorsys.hls_to_rgb(hue / 360, 0.22, 0.80)
        to_hex = lambda rgb: '#' + ''.join(f"{round(c * 255):02X}" for c in rgb)  # noqa: E731
        out.append((to_hex(fill), to_hex(ink)))
    return out


PALETTE = [{'fill': f, 'ink': i} for f, i in _BASE + tuple(_more(36))]


def clean_name(name):
    """A collector's name as it's shown: spaces tidied, a word all in capitals or all in small letters with a capital
    first ('KYLE', 'kyle' -> 'Kyle'; 'McKenzie' stays), at most 40 characters. 'Kyle to collect' -> 'Kyle'."""
    text = " ".join(str(name or '').split())
    text = re.sub(r'\s+(?:to\s+)?collect(?:s|ing|ion)?\b.*$', '', text, flags=re.I).strip(' -:,.')
    word = lambda w: w[:1].upper() + (w[1:].lower() if w in (w.upper(), w.lower()) else w[1:])  # noqa: E731
    return " ".join(word(w) for w in text.split())[:MAX_NAME]


def name_key(name):
    """Comparison key: capitals, spaces and punctuation don't make another person ('KYLE' == 'kyle.')."""
    return _norm(clean_name(name))


def standard_colour(name_or_key):
    """The colour a name has on every computer, or None: Kyle yellow, Gessie blue, any name with 'client' green."""
    key = name_key(name_or_key)
    if key in STANDARD:
        return dict(STANDARD[key])
    if re.search(r'\bclients?\b', key):
        return dict(CLIENT_COLOUR)
    return None


def label_text(name):
    """What the courier region says: 'KYLE TO COLLECT', or 'TO BE COLLECTED' when no one is named yet."""
    name = clean_name(name)
    return f"{name.upper()} TO COLLECT" if name else "TO BE COLLECTED"


def assign(names, saved):
    """{name key: colour} for every name in names. saved: {name key: colour} already given. A standard name (Kyle,
    Gessie, 'client') always gets its STANDARD colour. A new name gets the first colour of PALETTE that no saved name
    (and no new name before it) has; new names take colours in alphabetical order, so the preview and Generate give
    the same colours however the consignments are listed. When every colour is taken they're used again, in order."""
    out, taken = {}, {c['fill'] for c in saved.values()} | RESERVED
    keys = {name_key(n) for n in names if name_key(n)}
    new = sorted(k for k in keys - set(saved) if not standard_colour(k))
    free = [c for c in PALETTE if c['fill'] not in taken]
    for n, key in enumerate(new):
        out[key] = dict(free[n] if n < len(free) else PALETTE[(n - len(free)) % len(PALETTE)])
    for key in keys:
        std = standard_colour(key)
        if std:
            out[key] = std
        elif key in saved:
            out[key] = dict(saved[key])
    return out


def as_rgb(hex_colour):
    """'#FFEB40' -> (1.0, 0.92, 0.25) for drawing in the PDF."""
    h = str(hex_colour or '').lstrip('#')
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4)) if len(h) == 6 else (0, 0, 0)
