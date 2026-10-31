"""One-click fixes the official address data offers: a street spelt wrong swapped for its official name, the number
and anything before it kept. No official data needed (pure functions)."""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from engine.addresses.address_reference import _street_option, _street_title  # noqa: E402

TYPES = {'RD': 'ROAD', 'ROAD': 'ROAD', 'ST': 'STREET', 'STREET': 'STREET', 'LANE': 'LANE', 'CENTRE': 'CENTRE'}


def fix(line1, street, line2=''):
    o = _street_option({'line1': line1, 'line2': line2}, street, TYPES)
    return o and (o['label'], o['fields'])


def test_street_words_swapped_number_and_shop_kept():
    assert fix("250 Centre Dander", "CENTRE DANDENONG ROAD") == ("250 Centre Dandenong Road", {'line1': "250 Centre Dandenong Road"})
    assert fix("Shop 2, 250 Centre Dander", "CENTRE DANDENONG ROAD")[0] == "Shop 2, 250 Centre Dandenong Road"
    assert fix("68 Rickets Rd", "RICKETTS ROAD")[0] == "68 Ricketts Road"


def test_street_written_as_one_word():
    assert fix("Unit 4/12 Glenhuntley Road", "GLEN HUNTLY ROAD")[0] == "Unit 4/12 Glen Huntly Road"


def test_street_in_line_2():
    assert fix("Westfield Doncaster", "DONCASTER ROAD", line2="619 Doncastr Rd") == ("619 Doncaster Road", {'line2': "619 Doncaster Road"})


def test_street_type_ending_the_line_is_not_a_name():
    # 'Lane' looks like LADNER, but it's the street's type: no fix rather than a wrong one
    assert fix("1 Nowhere Lane", "LADNER COURT") is None


def test_suffix_written_out():
    assert _street_title("HIGH STREET S") == "High Street South"
    assert fix("12 Hight St", "HIGH STREET S")[0] == "12 High Street South"


def _places(*places):
    """A small official-data database: (id, name, state, [postcodes])."""
    import sqlite3
    from engine.addresses.address_reference import SCHEMA, _key
    db = sqlite3.connect(":memory:")
    db.executescript(SCHEMA)
    for i, name, state, pcs in places:
        db.execute("INSERT INTO locality VALUES (?, ?, ?, ?, '')", (i, name, _key(name), state))
        db.execute("INSERT INTO locality_name VALUES (?, ?)", (_key(name), i))
        for pc in pcs:
            db.execute("INSERT INTO locality_postcode VALUES (?, ?)", (pc, i))
    return db


def test_a_suburb_written_in_an_address_line_is_found_to_move_into_suburb():
    """The suburb written in a street line instead of Suburb, from the official data: at the end of the line (with a
    state and postcode after it, or nothing else filled in), or as a whole part between commas; a street that shares
    a suburb's name isn't one."""
    from engine.addresses.address_reference import _suburb_in_lines
    db = _places((1, "BONDI JUNCTION", "NSW", ["2022"]), (2, "BONDI", "NSW", ["2026"]), (3, "SYDNEY", "NSW", ["2000"]),
                 (4, "RICHMOND", "VIC", ["3121"]), (5, "RICHMOND", "NSW", ["2753"]))
    find = lambda **a: (lambda r: r and (r[0]["name"], r[0]["state"], r[1], r[2], r[3]))(_suburb_in_lines(db, a))
    assert find(line1="Shop 5012/500 Oxford St, Bondi Junction NSW 2022") == \
        ("BONDI JUNCTION", "NSW", "line1", "Shop 5012/500 Oxford St", "2022")
    assert find(line1="Shop 5012/500 Oxford St, Bondi Junction") == ("BONDI JUNCTION", "NSW", "line1", "Shop 5012/500 Oxford St", "")
    assert find(line1="Shop 5012, Bondi Junction, 500 Oxford St") == ("BONDI JUNCTION", "NSW", "line1", "Shop 5012, 500 Oxford St", "")
    assert find(line1="Level 2, 340 Kent St", line2="Sydney NSW 2000") == ("SYDNEY", "NSW", "line2", "", "2000")
    assert find(line1="12 Bondi Rd") is None  # a street, not the suburb
    assert find(line1="5 Swan St, Richmond 2753")[1] == "NSW"  # the postcode written with it picks the state
    assert find(line1="5 Swan St, Richmond", state="VIC")[1] == "VIC"


def test_a_city_with_no_postcode_of_its_own_offers_its_suburbs_postcode():
    """'Hervey Bay' is kept by the official data with no postcode: its suburbs (which carry the name too) give it."""
    from engine.addresses.address_reference import _assess_in
    db = _places((1, "HERVEY BAY", "QLD", []), (2, "PIALBA", "QLD", ["4655"]), (3, "URANGAN", "QLD", ["4655"]))
    db.executemany("INSERT INTO locality_name VALUES ('HERVEY BAY', ?)", [(2,), (3,)])
    a = {'line1': '12 Main St', 'line2': '', 'suburb': 'Hervey Bay', 'state': 'QLD', 'postcode': ''}
    r = _assess_in('AU', db, a)
    assert r['message'].startswith('No postcode: PIALBA QLD is 4655')
    assert r['options'][0]['fields'] == {'suburb': 'PIALBA', 'state': 'QLD', 'postcode': '4655'}
    assert _assess_in('AU', db, dict(a, postcode='4655'))['fields'].get('postcode') is None  # written in: no problem


def test_a_country_written_after_the_suburb_is_read_as_the_country():
    from engine.courier.courier_export import parse_address_text
    read = lambda t: (lambda r: (r['suburb'], r['state'], r['postcode'], r['country']))(parse_address_text(t))  # noqa: E731
    assert read('Billboard Solutions, 3F Miro Place, Albany NZ') == ('Albany', '', '', 'NZ')
    assert read('5 High St, Lower Hutt NZ 5010') == ('Lower Hutt', '', '5010', 'NZ')
    assert read('5 High St, Lower Hutt 5010 New Zealand') == ('Lower Hutt', '', '5010', 'New Zealand')
    # Australian addresses read as before
    assert read('34 Robinson St North, Wiley Park NSW 2195') == ('Wiley Park', 'NSW', '2195', '')
    assert read('Shop 4, 12 Main St, Richmond VIC 3121, Australia') == ('Richmond', 'VIC', '3121', 'Australia')


def test_a_postcode_is_filled_only_when_the_suburb_leaves_no_doubt(monkeypatch):
    """postcode_for: a suburb that is one place with one postcode gets it (and its state); a name in several states
    needs the state; one New Zealand has too needs the state; several postcodes, an NZ address or an unknown suburb
    are left to the user (the why says what to do, or is blank when the usual check says it)."""
    from engine.addresses import address_reference as ref
    au = _places((1, "TOOWOOMBA", "QLD", ["4350"]), (2, "RICHMOND", "VIC", ["3121"]), (3, "RICHMOND", "NSW", ["2753"]),
                 (4, "ALBANY", "WA", ["6330"]), (5, "BLACKTOWN", "NSW", ["2148", "2147"]))
    nz = _places((9, "ALBANY", "", []))
    monkeypatch.setattr(ref, "_db", lambda country: {"AU": au, "NZ": nz}.get(country))
    fill = lambda *a: ref.postcode_for(*a)  # noqa: E731
    assert fill("Toowoomba", "", "") == ({'postcode': '4350', 'state': 'QLD', 'place': 'TOOWOOMBA QLD 4350'}, '')
    assert fill("TOOWOOMBA", "QLD", "AU")[0]['postcode'] == '4350'
    none, why = fill("Richmond", "", "")
    assert none is None and why.startswith("There's a RICHMOND in NSW, VIC")
    assert fill("Richmond", "VIC", "")[0]['postcode'] == '3121'  # the state decides
    none, why = fill("Albany", "", "AU")
    assert none is None and "Australia and in New Zealand" in why
    assert fill("Albany", "WA", "")[0]['postcode'] == '6330'  # an Australian state says it's the Australian one
    none, why = fill("Blacktown", "NSW", "")
    assert none is None and "more than one postcode (2147, 2148)" in why
    assert fill("Albany", "", "NZ") == (None, '')  # LINZ has no postcodes
    assert fill("Toowooba", "QLD", "") == (None, '')  # not on record: the official data check offers the fixes
    assert fill("Toowoomba", "NSW", "") == (None, '')  # not in that state: the check says so
