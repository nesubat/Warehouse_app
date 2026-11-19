"""Collect from warehouse (a box on every consignment, ticked by itself only for 'Kyle to collect'),
each collector's own colour in the courier region, postcodes from verified addresses, and the header words that mean
the same thing (Size = Dimensions, Artwork = Thumbnail) in vertical and horizontal files."""
import html
import json
import os
import re
import sys

import fitz
import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from engine.addresses.address_book import AddressBook  # noqa: E402
from engine.addresses.collectors import (KYLE_COLOUR, GESSIE_COLOUR, CLIENT_COLOUR, PALETTE, assign, clean_name,  # noqa: E402
                                         label_text)
from engine.courier.courier_export import build_consignments, collection, sendable  # noqa: E402
from engine.packing_labels.horizontal_distribution import _row_field  # noqa: E402
from engine.packing_labels.packing_label_generator import _header_field, generate_packing_labels  # noqa: E402
from tests.test_horizontal import client, files  # noqa: E402,F401  (fixtures: the page end to end)


# ---------------------------------------------------------------- header words

def test_size_and_dimensions_artwork_and_thumbnail_are_the_same_column_in_both_layouts():
    for word in ("Size", "Sizes", "Item Size (mm)", "Dimensions", "Dimension", "Dims"):
        assert _header_field(word) == "dim_combined", word  # vertical header
        assert _row_field(word) == "dimension", word  # horizontal row label
    for word in ("Artwork", "Art work", "Thumbnail", "Image", "Picture", "Visual"):
        assert _header_field(word) == "thumbnail", word
        assert _row_field(word) == "thumbnail", word
    # A size of something else isn't the item's; a part of a size is that part
    for word in ("Store Size", "Carton Size", "Size Code"):
        assert _header_field(word) != "dim_combined", word
    assert _header_field("Size W (mm)") == "dim_w" and _header_field("Dimension H") == "dim_h"
    assert _header_field("Description") == "desc"


# ---------------------------------------------------------------- who collects

def test_the_box_ticks_itself_only_for_kyle():
    blank = {'receiver': 'Medibank Sydney', 'line1': '', 'line2': '', 'suburb': '', 'state': '', 'postcode': '', 'raw': ''}
    kyle = {**blank, 'line1': 'Kyle collect from warehouse', 'raw': 'Kyle collect from warehouse'}
    john = {**blank, 'line1': 'John Smith will collect', 'raw': 'John Smith will collect'}
    real = {**blank, 'line1': '12 Main St', 'suburb': 'RICHMOND', 'state': 'VIC', 'postcode': '3121', 'raw': '12 Main St'}
    assert collection(kyle) == {'on': True, 'name': 'Kyle', 'auto': True, 'suggested': 'Kyle'}
    # Anyone else, or a blank address: ticked by hand (a name the file gives is filled in for the user)
    assert collection(john) == {'on': False, 'name': '', 'auto': False, 'suggested': 'John Smith'}
    assert collection(blank)['on'] is False and collection(real)['on'] is False
    # By hand: on for any address, with a name; off even when it ticked itself
    assert collection(real, {'collect': True, 'collect_by': 'sam lee'})['name'] == 'Sam Lee'
    assert collection(john, {'collect': True})['name'] == 'John Smith' 
    assert collection(kyle, {'collect': False})['on'] is False
    assert collection(kyle, {'collect': True})['name'] == 'Kyle'  # ticked again: the file's name stays


def _group(store, spec, address='', suburb='', state='', postcode=''):
    return {'pack_spec_name': spec, 'store_name': store, 'receiver_name': '', 'contact': '', 'install': False,
            'address_1': address, 'address_2': '', 'suburb': suburb, 'state': state, 'postcode': postcode, 'country': '',
            'excel_rows': [2], 'items': [{'job_no': 'J476699-01', 'qty': '1', 'desc': '', 'dimension': '', 'material': '',
                                          'install': '', 'notes': '', 'thumbnails': []}]}


def test_collected_consignments_are_left_out_of_the_courier_csv_and_count_their_boxes():
    groups = {'A': _group('Chadstone', '2 x OB600100100', '1341 Dandenong Rd', 'MALVERN EAST', 'VIC', '3145'),
              'B': _group('Medibank Sydney', 'OB600100100', 'Kyle collect from warehouse'),  # Kyle: collected
              'C': _group('Corio', 'OB600100100', '1 Bacchus Marsh Rd', 'CORIO', 'VIC', '3214')}
    cons, cartons, warnings = build_consignments(groups)
    by_store = {c['destination']['receiver']: c for c in cons}
    assert by_store['Medibank Sydney']['collection']['on'] and not by_store['Chadstone']['collection']['on']
    assert by_store['Chadstone']['packages'] == 2 and by_store['Corio']['packages'] == 1  # '2 x' is two boxes
    kept, kept_cartons = sendable(cons, cartons)
    assert {c['destination']['receiver'] for c in kept} == {'Chadstone', 'Corio'}
    assert any('Collected from the warehouse' in w for w in warnings)
    assert not any(w.startswith('No address for') for w in warnings)  # collected isn't "missing"
    # A blank address isn't collected by itself: it's "No address", for the user to fill in or tick
    blank = build_consignments({'B': _group('Medibank Sydney', 'OB600100100'), 'C': groups['C']})
    assert not any(c['collection']['on'] for c in blank[0]) and any(w.startswith('No address for') for w in blank[2])
    # Ticked by hand on an addressed store: out of the CSV too
    edits = {by_store['Corio']['id']: {'collect': True, 'collect_by': 'Sam'}}
    cons, cartons, _ = build_consignments(groups, edits)
    assert {c['destination']['receiver'] for c in sendable(cons, cartons)[0]} == {'Chadstone'}
    # A file with only store names: nothing is collected by itself
    names_only = {'A': _group('Chadstone', 'OB600100100'), 'B': _group('Corio', 'OB600100100')}
    assert not any(c['collection']['on'] for c in build_consignments(names_only)[0])


def test_each_collector_keeps_a_colour_and_kyles_yellow_is_his_alone(tmp_path):
    book = AddressBook(str(tmp_path / "book.db"))
    assert [(c['name'], c['fill'], c['ink']) for c in book.collectors()] == [('Kyle', KYLE_COLOUR['fill'], KYLE_COLOUR['ink'])]
    # New names take the free colours in alphabetical order, whatever order they're typed in (preview = Generate)
    assert book.collector_colours(['zoe', 'Amy']) == book.collector_colours(['Amy', 'zoe'])
    colours = book.remember_collectors(['Sam', 'amy', 'Kyle to collect'])
    assert colours['kyle'] == KYLE_COLOUR and colours['amy'] == PALETTE[0] and colours['sam'] == PALETTE[1]
    # Saved: the same colour in the next job, a new name gets the next free one, no one gets Kyle's yellow
    assert book.collector_colours(['SAM'])['sam'] == PALETTE[1]
    assert book.collector_colours(['Bob'])['bob'] == PALETTE[2]
    fills = [assign([f"Person {n}"], {})[f"person {n}"]['fill'] for n in range(60)]
    assert KYLE_COLOUR['fill'] not in fills and KYLE_COLOUR['fill'] not in {c['fill'] for c in PALETTE}
    # The standard colours, the same on every computer and no one else's: Gessie blue, any 'client' name green
    std = book.collector_colours(['Gessie', 'Client - Bob', 'The client'])
    assert std['gessie'] == GESSIE_COLOUR and std['client bob'] == CLIENT_COLOUR == std['the client']
    assert not {GESSIE_COLOUR['fill'], CLIENT_COLOUR['fill']} & set(fills)
    # Removing a name frees its colour; Kyle can't be removed
    assert book.delete_collector('amy') and not book.delete_collector('kyle')
    assert book.collector_colours(['Bob'])['bob'] == PALETTE[0]
    assert clean_name('  kyle   to collect ') == 'Kyle' and label_text('') == 'TO BE COLLECTED' and label_text('sam lee') == 'SAM LEE TO COLLECT'


def test_the_courier_region_is_filled_with_the_collectors_colour(tmp_path):
    groups = {'G1': _group('Medibank Sydney', 'OB600100100')}
    sam = PALETTE[1]
    out = tmp_path / "l.pdf"
    generate_packing_labels(groups, str(out), ['desc', 'job_no', 'qty'],
                            {'G1': {'address': '', 'receiver': 'Medibank Sydney', 'contact': '', 'sent': False,
                                    'collect': 'SAM TO COLLECT', 'collect_colour': sam}})
    page = fitz.open(out)[0]
    rgb = lambda h: tuple(int(h[i:i + 2], 16) / 255 for i in (1, 3, 5))  # noqa: E731
    assert any(d.get('fill') and all(abs(a - b) < 0.01 for a, b in zip(d['fill'], rgb(sam['fill']))) for d in page.get_drawings())
    span = next(s for b in page.get_text('dict')['blocks'] for ln in b.get('lines', []) for s in ln['spans'] if 'SAM TO COLLECT' in s['text'])
    assert span['color'] == int(sam['ink'][1:], 16)


# ---------------------------------------------------------------- postcodes from verified addresses

def _verified(book, receiver, line1, suburb, state, postcode):
    entry = book.create({'receiver': receiver, 'line1': line1, 'suburb': suburb, 'state': state, 'postcode': postcode})
    book.mark_verified(entry['id']) if hasattr(book, 'mark_verified') else book._connect().execute(
        "UPDATE addresses SET verified_at = 1 WHERE id = ?", (entry['id'],))
    book._connect().commit()


def test_a_verified_suburbs_postcode_and_state_are_put_in_straight_away(tmp_path, monkeypatch):
    book = AddressBook(str(tmp_path / "book.db"))
    _verified(book, 'Myer Brisbane', '91 Queen St', 'BRISBANE', 'QLD', '4001')
    groups = {'A': _group('David Jones', 'OB600100100', '149 Queen St', 'Brisbane', 'QLD', '4000'),
              'B': _group('Target', 'OB600100100', '20 Adelaide St', 'Brisbane', '', ''),
              'C': _group('Kmart', 'OB600100100', '', 'Brisbane', 'QLD', '')}  # no street: not an address yet
    cons, _, _ = build_consignments(groups, verified_place=book.verified_place)
    by_store = {c['destination']['receiver']: c for c in cons}
    assert by_store['David Jones']['destination']['postcode'] == '4001'
    assert by_store['David Jones']['postcode_from'] == ('4000', '4001') and not by_store['David Jones']['edited']
    assert by_store['David Jones']['state_from'] is None
    # No state in the file: the verified state goes in too
    target = by_store['Target']['destination']
    assert (target['state'], target['postcode']) == ('QLD', '4001') and by_store['Target']['state_from'] == 'QLD'
    assert by_store['Kmart']['postcode_from'] is None
    # A postcode typed in ✏️ is the user's: left as typed
    edits = {by_store['David Jones']['id']: {'postcode': '4000'}}
    cons, _, _ = build_consignments(groups, edits, verified_place=book.verified_place)
    assert next(c for c in cons if c['destination']['receiver'] == 'David Jones')['destination']['postcode'] == '4000'
    # Verified addresses with different postcodes settle nothing: each is a one-click choice (state and postcode)
    _verified(book, 'GPO Brisbane', '261 Queen St', 'BRISBANE', 'QLD', '4000')
    assert book.verified_place('BRISBANE', 'QLD') is None and book.verified_place('BRISBANE', 'NSW') is None
    import app as app_module
    monkeypatch.setattr(app_module, "ADDRESS_BOOK", book)
    choices = app_module._verified_choices({'suburb': 'BRISBANE', 'state': '', 'postcode': '', 'country': 'AU'})
    assert sorted(c['fields']['postcode'] for c in choices) == ['4000', '4001']
    assert all(c['fields']['state'] == 'QLD' for c in choices)


def test_the_same_stores_saved_address_comes_before_the_official_street_suggestions(tmp_path, monkeypatch):
    import app as app_module
    book = AddressBook(str(tmp_path / "book.db"))
    book.create({'receiver': 'Rouse Hill', 'line1': 'Shop GR009 Windsor Rd', 'suburb': 'ROUSE HILL', 'state': 'NSW', 'postcode': '2155'})
    monkeypatch.setattr(app_module, "ADDRESS_BOOK", book)
    street_note = {'level': 'note', 'message': 'Street not on record in ROUSE HILL', 'fields': {}, 'options': [{'label': 'Windsor Road'}]}
    monkeypatch.setattr(app_module, "_reference_for", lambda *a, **k: dict(street_note))
    groups = {'G': _group('ROUSE HILL', 'OB600100100', 'S GR009 Rouse Hill Town Ctr, Winsor Road', 'ROUSE HILL', 'NSW', '2155')}
    cons, _, _ = build_consignments(groups)
    c = app_module._check_against_book(cons, {}, app_module._book_options(groups, {}))[0]
    assert c['book_options'] and c['ref']['level'] == 'none' and c['ref_held']['level'] == 'note'
    # No same-store address in the book: the official street suggestions show
    groups['G']['store_name'] = 'Somewhere Else'
    cons, _, _ = build_consignments(groups)
    c = app_module._check_against_book(cons, {}, app_module._book_options(groups, {}))[0]
    assert not c['book_options'] and c['ref']['level'] == 'note' and c['ref_held'] is None


# ---------------------------------------------------------------- the page

def test_preview_has_excel_filters_and_collect_boxes_and_generate_leaves_collected_out(client, files):  # noqa: F811
    final, sig, _ = files
    c, app_module, projects = client
    with open(final, "rb") as f, open(sig, "rb") as s:
        c.post("/packing-labels/horizontal", data={"file": (f, final.name), "signature_file": (s, sig.name)},
               content_type="multipart/form-data", follow_redirects=True)
    form = {"filename": app_module.secure_filename(final.name), "all_tabs": ["Spring Sale", "Easter"],
            "selected_tabs": ["Spring Sale", "Easter"], "header_row_Spring Sale": "8", "header_row_Easter": "8"}
    page = c.post("/packing-labels/horizontal", data={**form, "preview": "1"}, follow_redirects=True).data.decode()
    for key in ("state", "country", "packages"):
        assert f'data-filter="{key}"' in page
    assert page.count('class="pl-collect-check"') == page.count('class="pl-con-row')  # one box per consignment
    rows = re.findall(r'<tr class="pl-con-row[^"]*" data-id="([^"]+)"[^>]*>(.*?)</tr>', page, re.S)
    chadstone = html.unescape(next(i for i, body in rows if 'Chadstone' in body))
    edits = {chadstone: {'collect': True, 'collect_by': 'sam lee'}}
    gen = {**form, "generate": "1", "who_pays": "S", "service_code": app_module.SERVICE_STORE.services()[0]['code'],
           "reference": "J476699", "project_name": "Collect test", "courier_csv": ["open360"],
           "attribute_order": "desc,job_no,barcode,qty", "consignment_edits": json.dumps(edits)}
    r = c.post("/packing-labels/horizontal", data=gen, follow_redirects=True)
    assert "Labels Successfully Generated" in r.data.decode(), r.data.decode()[:1500]
    folder = next(p for p in projects.iterdir() if p.is_dir())
    csv_text = open(folder / next(n for n in os.listdir(folder) if n.endswith("Open360.csv")), encoding="utf-8").read()
    assert "Chadstone" not in csv_text and "Corio" in csv_text
    pdf = fitz.open(folder / next(n for n in os.listdir(folder) if n.startswith("Packing Labels Only")))
    assert any("SAM LEE TO COLLECT" in " ".join(p.get_text().split()) for p in pdf)  # wrapped to fit the region
    saved = {x['key']: x for x in app_module.ADDRESS_BOOK.collectors()}
    assert saved['sam lee']['use_count'] == 1 and saved['sam lee']['fill'] != KYLE_COLOUR['fill']


@pytest.mark.parametrize("key", ["kyle", ""])
def test_kyle_and_nothing_cannot_be_removed(client, key):  # noqa: F811
    c, app_module, _ = client
    assert c.post("/api/collectors/delete", json={"key": key}).status_code == 400


def test_back_up_and_move_brings_the_collectors_and_keeps_colours_here(tmp_path):
    from engine.core import data_transfer
    a, b = tmp_path / "a", tmp_path / "b"
    a.mkdir(), b.mkdir()
    book_a, book_b = AddressBook(str(a / "address_book.db")), AddressBook(str(b / "address_book.db"))
    book_a.remember_collectors(['Amy', 'Sam'])  # Amy: PALETTE[0], Sam: PALETTE[1] there
    book_b.remember_collectors(['Zoe'])  # Zoe already has PALETTE[0] here
    data, _ = data_transfer.export_zip(str(a), book_a, ['address_book'], None)
    data_transfer.import_zip(data, {'address_book': 'merge'}, str(b), book_b, None)
    here = {c['key']: c['fill'] for c in book_b.collectors()}
    assert here['zoe'] == PALETTE[0]['fill']  # a name here keeps its colour
    assert here['sam'] == PALETTE[1]['fill']  # brought over with its own
    assert here['amy'] not in (PALETTE[0]['fill'], PALETTE[1]['fill'], KYLE_COLOUR['fill'])  # its colour was taken here: a free one
    assert len(set(here.values())) == len(here)


def test_address_book_streets_are_offered_only_for_the_same_store(tmp_path, monkeypatch):
    import app as app_module
    book = AddressBook(str(tmp_path / "book.db"))
    monkeypatch.setattr(app_module, "ADDRESS_BOOK", book)
    other = book.create({'receiver': 'Macpac Rouse Hill', 'line1': 'Shop AGR042 Windsor Rd', 'suburb': 'ROUSE HILL',
                         'state': 'NSW', 'postcode': '2155'})
    book.mark_verified(other['id'])
    groups = {'G': _group('ROUSE HILL', 'OB600100100', 'S GR009 Rouse Hill Town Ctr, Winsor Road', 'ROUSE HILL', 'NSW', '2155')}
    assert app_module._book_options(groups, {}) == {}  # a similar name isn't this store
    book.create({'receiver': 'Rouse Hill', 'line1': 'Shop GR009 Windsor Rd', 'suburb': 'ROUSE HILL', 'state': 'NSW', 'postcode': '2155'})
    offered = list(app_module._book_options(groups, {}).values())[0]
    assert [e['receiver'] for e, _ in offered] == ['Rouse Hill']
