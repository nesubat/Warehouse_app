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
