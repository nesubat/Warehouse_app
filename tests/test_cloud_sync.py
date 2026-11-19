"""The shared address book: two computers' address books syncing through the real Apps Script (cloud/apps_script/
Code.gs, run in node by tests/cloud_server.js with stand-ins for Google's services). Skipped without node."""
import json
import os
import shutil
import subprocess
import sys
import time
import urllib.request

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from engine.addresses.address_book import AddressBook  # noqa: E402
from engine.addresses.cloud_sync import CloudSync, SyncError  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
pytestmark = pytest.mark.skipif(not shutil.which('node'), reason="node isn't installed (it runs the Apps Script here)")


@pytest.fixture()
def cloud():
    proc = subprocess.Popen(['node', os.path.join(HERE, 'cloud_server.js'), 'sekret'], stdout=subprocess.PIPE, text=True)
    port = int(proc.stdout.readline())
    url = f"http://127.0.0.1:{port}/"
    yield url
    proc.terminate()
    proc.wait(5)


def state(url):
    with urllib.request.urlopen(url + "__state") as r:
        return json.load(r)


def sheet(url, name):
    main = state(url)['books'][0]
    rows = next(s['rows'] for s in main['sheets'] if s['name'] == name)
    head = rows[0]
    return [dict(zip(head, r)) for r in rows[1:]]


def log_rows(url):
    books = state(url)['books']
    return [r for b in books if b['title'].startswith('Warehouse address book log') for r in b['sheets'][0]['rows'][1:]]


def pc(tmp_path, url, name, user='Sam'):
    folder = tmp_path / name
    folder.mkdir()
    book = AddressBook(str(folder / "address_book.db"))
    sync = CloudSync(book, url, 'sekret', str(folder / "sync_settings.json"), pc=name)
    if user:
        sync.set_user_name(user)
    return book, sync


def add(book, receiver, line1, suburb='RICHMOND', state_='VIC', postcode='0803'):
    return book.create({'receiver': receiver, 'line1': line1, 'suburb': suburb, 'state': state_, 'postcode': postcode})


def entry(book, receiver):
    return next((dict(r) for r in book._connect().execute("SELECT * FROM addresses WHERE receiver = ?", (receiver,))), None)


def test_an_address_saved_on_one_computer_reaches_the_other_with_who_did_it(tmp_path, cloud):
    a, sa = pc(tmp_path, cloud, 'PC-A', 'Sam Lee')
    b, sb = pc(tmp_path, cloud, 'PC-B', 'Gessie')
    add(a, 'Myer Chadstone', '1341 Dandenong Rd')
    a.record_used([({'receiver': 'Myer Chadstone', 'contact': '', 'line1': '1341 Dandenong Rd', 'line2': '', 'suburb': 'RICHMOND',
                     'state': 'VIC', 'postcode': '0803', 'country': 'AU', 'authority_to_leave': False}, ['myer chadstone raw'])])
    assert sa.sync_now()['pushed'] == 1
    shared = sheet(cloud, 'Addresses')
    assert shared[0]['postcode'] == '0803'  # kept as text: not 803
    assert shared[0]['changed_by'] == 'Sam Lee' and shared[0]['changed_pc'] == 'PC-A'
    assert sb.sync_now()['pulled'] == 1
    got = entry(b, 'Myer Chadstone')
    assert got['postcode'] == '0803' and got['gid'] == entry(a, 'Myer Chadstone')['gid'] and not got['dirty']
    assert b._connect().execute("SELECT COUNT(*) FROM aliases WHERE source_key = 'myer chadstone raw'").fetchone()[0] == 1
    assert b.search('chadstone')[0]  # the search index has it
    assert any(r[1] == 'Sam Lee' and r[2] == 'PC-A' and r[3] == 'added' for r in log_rows(cloud))
    # Nothing new: nothing goes up again, nothing comes back
    assert sa.sync_now() == {'pushed': 0, 'pulled': 0} and sb.sync_now() == {'pushed': 0, 'pulled': 0}


def test_verified_beats_unverified_and_otherwise_the_latest_wins(tmp_path, cloud):
    a, sa = pc(tmp_path, cloud, 'PC-A')
    b, sb = pc(tmp_path, cloud, 'PC-B')
    e = add(a, 'Target Corio', '1 Bacchus Marsh Rd', 'CORIO', 'VIC', '3214')
    sa.sync_now()
    sb.sync_now()
    eb = entry(b, 'Target Corio')
    # A verifies it; B then changes the street (later, but not verified): A's verified copy stays everywhere
    a.mark_verified(e['id'])
    time.sleep(0.01)
    b.update(eb['id'], {'line1': '9 Wrong St'})
    sa.sync_now()
    sb.sync_now()
    assert entry(b, 'Target Corio')['line1'] == '1 Bacchus Marsh Rd' and entry(b, 'Target Corio')['verified_at']
    assert any(r[3] == 'kept the shared copy' and 'verified' in r[6] for r in log_rows(cloud))
    # Both verified now: the latest change wins
    time.sleep(0.01)
    b.update(entry(b, 'Target Corio')['id'], {'line2': 'Shop 12'})
    sb.sync_now()
    sa.sync_now()
    assert entry(a, 'Target Corio')['line2'] == 'Shop 12'


def test_the_same_address_saved_on_two_computers_becomes_one_entry(tmp_path, cloud):
    a, sa = pc(tmp_path, cloud, 'PC-A')
    b, sb = pc(tmp_path, cloud, 'PC-B')
    add(a, 'Kmart Knox', '425 Burwood Hwy', 'WANTIRNA SOUTH', 'VIC', '3152')
    add(b, 'Kmart Knox', '425 Burwood Hwy', 'WANTIRNA SOUTH', 'VIC', '3152')
    sa.sync_now()
    sb.sync_now()
    sa.sync_now()
    assert len(sheet(cloud, 'Addresses')) == 1
    assert entry(a, 'Kmart Knox')['gid'] == entry(b, 'Kmart Knox')['gid']


def test_deleting_is_only_on_this_computer_and_stays_deleted_here(tmp_path, cloud):
    a, sa = pc(tmp_path, cloud, 'PC-A')
    b, sb = pc(tmp_path, cloud, 'PC-B')
    e = add(a, 'Big W Highpoint', '120 Rosamond Rd', 'MARIBYRNONG', 'VIC', '3032')
    sa.sync_now()
    sb.sync_now()
    a.delete(e['id'])
    sa.sync_now()
    assert len(sheet(cloud, 'Addresses')) == 1 and entry(b, 'Big W Highpoint')  # not deleted anywhere else
    sb.sync_now()
    sa.sync_now()
    assert entry(a, 'Big W Highpoint') is None  # a pull doesn't bring it back
    # Changed in the shared book after the delete (here: B corrects it): it comes back, as the newer address
    time.sleep(0.01)
    b.update(entry(b, 'Big W Highpoint')['id'], {'line1': '120-200 Rosamond Rd'})
    sb.sync_now()
    sa.sync_now()
    assert entry(a, 'Big W Highpoint')['line1'] == '120-200 Rosamond Rd'


def test_a_computer_offline_for_a_while_catches_up_both_ways(tmp_path, cloud):
    a, sa = pc(tmp_path, cloud, 'PC-A')
    b, sb = pc(tmp_path, cloud, 'PC-B')
    for n in range(30):
        add(a, f'Store {n}', f'{n} Main St', 'RICHMOND', 'VIC', '3121')
    add(b, 'Offline Store', '7 Side St', 'CORIO', 'VIC', '3214')  # B saved this while offline
    sa.sync_now()
    import engine.addresses.cloud_sync as cs
    old = cs.PULL_PAGE
    cs.PULL_PAGE = 7  # pages of 7: the pull keeps going until it has everything
    try:
        sb.sync_now()
    finally:
        cs.PULL_PAGE = old
    sa.sync_now()
    assert b._connect().execute("SELECT COUNT(*) FROM addresses").fetchone()[0] == 31
    assert entry(a, 'Offline Store')


def test_nothing_is_pushed_without_a_name_or_with_a_wrong_key(tmp_path, cloud):
    a, sa = pc(tmp_path, cloud, 'PC-A', user=None)
    add(a, 'No Name Store', '1 High St', 'RICHMOND', 'VIC', '3121')
    sa.sync_now()  # pulls, pushes nothing
    assert sheet(cloud, 'Addresses') == [] and 'name' in sa.status()['error']
    sa.set_user_name('  Kyle  ')
    assert sa.user_name() == 'Kyle' and sa.sync_now()['pushed'] == 1
    bad = CloudSync(a, cloud, 'wrong', str(tmp_path / 'x.json'), pc='PC-X')
    with pytest.raises(SyncError, match='key'):
        bad.sync_now()


def test_a_generate_that_resaves_an_address_unchanged_pushes_nothing(tmp_path, cloud):
    a, sa = pc(tmp_path, cloud, 'PC-A')
    add(a, 'Myer Bourke St', '314 Bourke St', 'MELBOURNE', 'VIC', '3000')
    sa.sync_now()
    a.record_used([({'receiver': 'Myer Bourke St', 'contact': '', 'line1': '314 Bourke St', 'line2': '', 'suburb': 'MELBOURNE',
                     'state': 'VIC', 'postcode': '3000', 'country': 'AU', 'authority_to_leave': False}, [])])
    assert sa.status()['waiting'] == 0


def test_the_address_book_page_asks_for_the_name_and_syncs(tmp_path, cloud, monkeypatch):
    import app as app_module
    book = AddressBook(str(tmp_path / "book.db"))
    monkeypatch.setattr(app_module, "ADDRESS_BOOK", book)
    monkeypatch.setattr(app_module, "SYNC", CloudSync(lambda: app_module.ADDRESS_BOOK, cloud, 'sekret',
                                                      str(tmp_path / "sync_settings.json"), pc='PC-T'))
    app_module.app.config["TESTING"] = True
    c = app_module.app.test_client()
    page = c.get("/address-book").data.decode()
    assert 'id="shared-book"' in page and 'id="sb-name"' in page and 'id="sb-name" hidden' not in page
    add(book, 'Page Store', '5 Page St', 'RICHMOND', 'VIC', '3121')
    assert c.post("/api/shared-book/name", json={"name": ""}).status_code == 400
    assert c.post("/api/shared-book/name", json={"name": "Sam"}).get_json()['user_name'] == 'Sam'
    done = c.post("/api/shared-book/sync").get_json()
    assert done['pushed'] == 1 and done['waiting'] == 0 and done['user_name'] == 'Sam'
    assert c.get("/api/shared-book/status").get_json()['last_push']
    assert 'id="sb-name" hidden' in c.get("/address-book").data.decode()
