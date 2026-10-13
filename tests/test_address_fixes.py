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
