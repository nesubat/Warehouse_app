"""Packs collected rather than sent ('Kyle collect from warehouse' in the address cell): no courier label, a red on
yellow 'KYLE TO COLLECT' / 'PACK x OF n' note in the courier region, n counted over that store's boxes only."""
import os
import sys

import fitz

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from engine.courier.courier_export import collect_note  # noqa: E402
from engine.packing_labels.packing_label_generator import generate_packing_labels  # noqa: E402


def test_collect_wording_names_who_collects():
    assert collect_note('Kyle collect from warehouse') == 'KYLE TO COLLECT'
    assert collect_note('KYLE TO COLLECT FROM WAREHOUSE') == 'KYLE TO COLLECT'
    assert collect_note('Kyle - collection from warehouse') == 'KYLE TO COLLECT'
    assert collect_note('John Smith will collect') == 'JOHN SMITH TO COLLECT'
    assert collect_note('Please pick up from warehouse') == 'TO BE COLLECTED'
    # Real addresses and store counters aren't collections
    assert collect_note('Shop 4, Collection Point Centre, 12 Main St, Richmond VIC 3121') == ''
    assert collect_note('Click & Collect desk, Westfield Knox') == ''
    assert collect_note('M. Hazelton, 34 Robinson St North, Wiley Park NSW') == ''
    assert collect_note('') == ''


def _group(store, spec, row):
    return {'pack_spec_name': spec, 'store_name': store, 'receiver_name': '', 'contact': '', 'install': False,
            'address_1': '', 'address_2': '', 'suburb': '', 'state': '', 'postcode': '', 'country': '',
            'excel_rows': [row], 'items': [{'job_no': 'J476699-01', 'qty': '2', 'desc': 'A1 poster', 'dimension': '',
                                            'material': '', 'install': '', 'notes': '', 'thumbnails': []}]}


def test_collect_note_counts_the_stores_own_boxes(tmp_path):
    # Kyle's store: a 2-box pack and a 1-box pack (3 boxes); another collected store with 1; one sent store
    groups = {'G1': _group('Medibank Sydney', '2 x OB600100100', 2), 'G2': _group('Medibank Sydney', 'OB600100100', 3),
              'G3': _group('Medibank Perth', 'OB600100100', 4), 'G4': _group('Medibank Hobart', 'OB600100100', 5)}
    kyle = {'address': '', 'receiver': 'Medibank Sydney', 'contact': '', 'sent': False, 'collect': 'KYLE TO COLLECT'}
    addresses = {'G1': kyle, 'G2': kyle,
                 'G3': {'address': '', 'receiver': 'Medibank Perth', 'contact': '', 'sent': False,
                        'collect': 'TO BE COLLECTED'},
                 'G4': {'address': '1 Elizabeth St, Hobart TAS 7000', 'receiver': 'Medibank Hobart', 'contact': '',
                        'sent': True, 'collect': ''}}
    out = tmp_path / "l.pdf"
    info = generate_packing_labels(groups, str(out), ['desc', 'job_no', 'barcode', 'qty'], addresses)
    doc = fitz.open(out)
    pages_of = lambda k: [p for box in info['copies'][k] for p in box]  # noqa: E731
    text = lambda k: " ".join(doc[p].get_text() for p in pages_of(k))  # noqa: E731
    assert [doc[p].get_text().count('KYLE TO COLLECT') for p in pages_of('G1')] == [1, 1]
    assert 'PACK 1 OF 3' in doc[info['copies']['G1'][0][0]].get_text()
    assert 'PACK 2 OF 3' in doc[info['copies']['G1'][1][0]].get_text()
    # The store the packs are for, in the panel too (not only under it)
    region = fitz.Rect(0, 0, 320, 430)
    assert 'Medibank Sydney' in doc[info['copies']['G1'][0][0]].get_textbox(region)
    assert 'PACK 3 OF 3' in text('G2') and 'KYLE TO COLLECT' in text('G2')
    assert 'TO BE COLLECTED' in text('G3') and 'PACK 1 OF 1' in text('G3')  # its own store, not Kyle's count
    assert 'COLLECT' not in text('G4') and 'PACK ' not in text('G4').replace('PACK AUDIT', '')
    # Red on yellow, in the courier region
    page = doc[pages_of('G1')[0]]
    assert any(d.get('fill') and abs(d['fill'][0] - 1) < 0.01 and d['fill'][2] < 0.5 for d in page.get_drawings())
    span = next(s for b in page.get_text('dict')['blocks'] for ln in b.get('lines', []) for s in ln['spans']
                if 'KYLE TO COLLECT' in s['text'])
    assert span['color'] == 0xCC0000
    assert 'NOT IN COURIER CSV' not in text('G1')



def test_an_installer_pack_goes_to_the_installer_named_in_the_address():
    """Install = Y and no Receiver Name: the name at the start of the address cell is the receiver (out of Address
    Line 1); Install N or blank keeps the store, as before; a Receiver Name column always wins."""
    from engine.courier.courier_export import resolve_destination
    who = lambda a, install, **g: (lambda d: (d['receiver'], d['line1']))(  # noqa: E731
        resolve_destination({'address_1': a, 'store_name': 'MACQUARIE ST #1177', 'install': install, **g}))
    assert who('M. Hazelton, 34 Robinson St North, Wiley Park NSW', True) == ('M. Hazelton', '34 Robinson St North')
    assert who('Aurora Signs, 1/39 Grimwade St, Mitchell ACT', True) == ('Aurora Signs', '1/39 Grimwade St')
    assert who('M. Hazelton, 34 Robinson St North, Wiley Park NSW', False) == ('MACQUARIE ST #1177', 'M. Hazelton')
    assert who('12 Main St, Richmond VIC 3121', True) == ('MACQUARIE ST #1177', '12 Main St')  # no name: the store
    assert who('M. Hazelton, 34 Robinson St North, Wiley Park NSW', True, receiver_name='Wilson Storage')[0] == 'Wilson Storage'
    # Collected by the installer: 'Installer: Kyle to collect' on the label
    assert who('Kyle collect from warehouse', True)[0] == 'Kyle to collect'
    assert who('Kyle collect from warehouse', False)[0] == 'MACQUARIE ST #1177'


def test_a_spec_with_a_zero_typed_twice_gets_a_hint_not_a_size():
    from engine.courier.packing_specs import likely_meant, parse_formula
    assert parse_formula('OB11800170170') == ('OB', None)  # never guessed for the CSV
    assert likely_meant('OB11800170170') == 'OB1180170170'
    assert likely_meant('OB1100100100100') == ''  # several ways to fix it: no hint
    # No side of 1000 cm or more, however it's written: refused, with the code it most likely is
    assert parse_formula('OB18500170170') == parse_formula('OB 18500 x 170 x 170') == ('OB', None)
    assert likely_meant('OB18500170170') == 'OB1850170170'
    assert likely_meant('OB 18500 x 170 x 170') == 'OB 1850 x 170 x 170'
    assert parse_formula('OB9990170170') == ('OB', (999, 17, 17))
    assert likely_meant('OB1370170170') == likely_meant('A4 Box') == ''
