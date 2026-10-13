"""Deleting several Address Book entries at once: the store, and the page's route (which another website can't call).
Every store is on temp files of the test's own."""
import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from engine.addresses.address_book import AddressBook  # noqa: E402


def entry(n):
    return {'receiver': f"Store {n}", 'contact': '', 'line1': f"{n} Example St", 'line2': '', 'suburb': 'CURLEWIS',
            'state': 'VIC', 'postcode': '3222', 'country': 'AU'}


@pytest.fixture()
def book(tmp_path):
    return AddressBook(str(tmp_path / "address_book.db"))


def test_delete_many_removes_only_those(book):
    ids = [book.create(entry(n))['id'] for n in range(1, 6)]
    assert book.delete_many(ids[:3] + [99999]) == 3  # one that isn't there is skipped
    assert book.count() == 2
    assert {book.get(i)['receiver'] for i in ids[3:]} == {"Store 4", "Store 5"}


@pytest.fixture()
def client(tmp_path, monkeypatch, book):
    import app as app_module
    monkeypatch.setattr(app_module, "ADDRESS_BOOK", book)
    app_module.app.config["TESTING"] = True
    return app_module.app.test_client()


def test_page_deletes_the_picked_addresses(client, book):
    ids = [book.create(entry(n))['id'] for n in range(1, 4)]
    r = client.post("/api/addresses/delete", json={'ids': ids[:2]}, headers={'Sec-Fetch-Site': 'same-origin'})
    assert r.status_code == 200 and r.get_json() == {'deleted': 2}
    assert book.count() == 1
    assert client.post("/api/addresses/delete", json={'ids': []}).status_code == 400
    assert client.post("/api/addresses/delete", json={'ids': ['x']}).status_code == 400


def test_other_websites_cant_delete(client, book):
    ids = [book.create(entry(1))['id']]
    r = client.post("/api/addresses/delete", json={'ids': ids}, headers={'Sec-Fetch-Site': 'cross-site'})
    assert r.status_code == 403 and book.count() == 1
