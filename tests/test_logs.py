"""The Logs page: the app's log file read back as runs people can follow (engine/core/log_view.py)."""
import os
import sys
from datetime import datetime

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from engine.core import log_view  # noqa: E402

NOW = datetime(2026, 10, 18, 16, 30, 0)
LOG = """\
2026-10-17 09:00:00 INFO    wh.book    MainThread | Address book: 3 Excel addresses, 0 with an address-book option
2026-10-18 16:23:43 INFO    wh.packing MainThread | Preview Final.xlsx: tabs ['Spring Sale']
2026-10-18 16:23:43 INFO    wh.book    MainThread | Address book: 8 Excel addresses, 0 with an address-book option
2026-10-18 16:23:43 INFO    wh.packing MainThread |   18 cartons in 8 consignments | job J476699 | service -
2026-10-18 16:23:43 INFO    wh.packing MainThread | Preview ready in 0.18s
2026-10-18 16:23:44 INFO    wh.packing MainThread | Generate Final.xlsx: tabs ['Spring Sale'], project 'Spring'
2026-10-18 16:23:44 DEBUG   wh.packing MainThread |     Label 1 - Corio -> ('01T1P1 Corio', '01T1P1')
2026-10-18 16:23:44 WARNING wh.packing MainThread |   tab Spring Sale: Row 10 has no store name
2026-10-18 16:23:45 INFO    wh.book    Thread-11 (work) | Address book checked against the official data (generate): 8 entries
2026-10-18 16:23:45 INFO    wh.packing MainThread | Generate done in 1.46s: Spring_261018_1623
2026-10-18 16:24:00 ERROR   wh.stitch  MainThread | Stitch FAILED after 0.10s
Traceback (most recent call last):
  File "app.py", line 1, in stitch
ValueError: no labels
2026-10-18 16:29:30 DEBUG   wh.packing Thread-12 (render) |   packing labels PDF big.pdf ...
2026-10-18 16:29:40 DEBUG   wh.packing MainThread |   Open360 CSV a.csv ...
2026-10-18 16:29:41 INFO    wh.packing MainThread |   Open360 CSV a.csv done in 0.01s
"""


@pytest.fixture()
def log_dir(tmp_path):
    (tmp_path / "warehouse.log").write_text(LOG, encoding='utf-8')
    (tmp_path / "warehouse.log.1").write_text(
        "2026-10-16 08:00:00 INFO    wh.specs   MainThread | Packing specs saved\n", encoding='utf-8')
    return str(tmp_path)


def test_lines_are_read_whatever_is_in_the_thread_name_or_message():
    entries = log_view.parse_lines(LOG.splitlines())
    bg = next(e for e in entries if e['thread'] == 'Thread-11 (work)')
    assert bg['area'] == 'book' and bg['text'].startswith('Address book checked')
    cartons = next(e for e in entries if 'cartons' in e['text'])
    assert cartons['depth'] == 1 and cartons['text'].endswith('| service -')  # a " | " in the message stays in it
    failed = next(e for e in entries if e['level'] == 'ERROR')
    assert failed['more'][-1] == 'ValueError: no labels'  # the traceback belongs to the line before


def test_a_run_and_everything_it_logged_is_one_activity(log_dir):
    entries = log_view.read_entries(log_dir)
    assert entries[0]['text'] == 'Packing specs saved'  # the older, rotated file first
    acts = log_view.activities(entries)
    gen = next(a for a in acts if a['title'].startswith('Generate Final.xlsx'))
    assert gen['result'] == 'Generate done in 1.46s: Spring_261018_1623'
    assert gen['level'] == 'WARNING'  # what it found
    assert len(gen['lines']) == 3  # the label line, the warning and the end; not the background check
    preview = next(a for a in acts if a['title'].startswith('Preview'))
    assert [x['area'] for x in preview['lines']] == ['book', 'packing', 'packing']  # the address book's line inside
    assert next(a for a in acts if a['title'].startswith('Address book checked'))['background']


def test_running_is_a_step_started_with_no_end_yet(log_dir):
    running = log_view.running(log_view.read_entries(log_dir), now=NOW)
    assert [(r['what'], r['background']) for r in running] == [('packing labels PDF big.pdf', True)]
    assert log_view.running(log_view.read_entries(log_dir), now=datetime(2026, 10, 18, 18, 0)) == []  # cut short


def test_filters_counts_and_older(log_dir):
    d = log_view.view(log_dir, now=NOW)
    titles = [a['title'] for a in d['activities']]
    assert titles[:3] == ['Open360 CSV a.csv ...', 'packing labels PDF big.pdf ...', 'Stitch FAILED after 0.10s']
    assert 'Packing specs saved' not in titles  # newest first, today only
    assert d['activities'][0]['details'][0]['text'] == 'Open360 CSV a.csv done in 0.01s'  # not part of the Stitch
    assert d['today'] == {'activities': 6, 'issues': 1, 'problems': 1}
    assert [a['title'] for a in log_view.view(log_dir, show='problems', now=NOW)['activities']] == ['Stitch FAILED after 0.10s']
    assert len(log_view.view(log_dir, q='corio', day='all', now=NOW)['activities']) == 1  # found in a run's details
    assert [a['area'] for a in log_view.view(log_dir, area='book', day='all', now=NOW)['activities']] == ['book', 'book']
    page = log_view.view(log_dir, day='all', limit=2, now=NOW)
    assert page['more'] and len(page['activities']) == 2
    rest = log_view.view(log_dir, day='all', before=page['oldest_shown'], now=NOW)
    assert rest['activities'][-1]['title'] == 'Packing specs saved' and not rest['more']


def test_page_and_api(log_dir, monkeypatch):
    import app as app_module
    monkeypatch.setattr(app_module, 'LOG_DIR', log_dir)
    client = app_module.app.test_client()
    page = client.get('/logs')
    assert page.status_code == 200 and b'Logs' in page.data
    assert b'href="/logs"' in client.get('/packing-specs').data  # in the nav on every page
    d = client.get('/api/logs?day=all&limit=abc').get_json()
    assert d['activities'] and d['version']
    assert client.get(f"/api/logs?since={d['version']}").get_json() == {'unchanged': True}


def test_a_failure_says_where_it_stopped_and_what_came_before(log_dir):
    d = log_view.view(log_dir, show='problems', now=NOW)
    stitch = d['activities'][0]
    assert stitch['failure'] == {'error': 'ValueError: no labels', 'where': 'app.py, line 1, in stitch', 'code': '',
                                 'inside': ''}
    gen = next(a for a in log_view.view(log_dir, now=NOW)['activities'] if a['title'].startswith('Generate'))
    assert [x['text'] for x in gen['before']][-2:] == ['18 cartons in 8 consignments | job J476699 | service -',
                                                       'Preview ready in 0.18s']  # what led up to it


def test_a_library_trace_names_the_apps_own_line_and_the_library():
    trace = ['Traceback (most recent call last):',
             r'  File "C:\w\Warehouse_app\app.py", line 3110, in create_horizontal_labels',
             '    page_info = generate_packing_labels(groups, out)',
             r'  File "C:\w\Warehouse_app\engine\packing_labels\packing_label_generator.py", line 1520, in generate_packing_labels',
             '    doc.save(output_pdf_path)',
             r'  File "C:\w\Warehouse_app\venv\Lib\site-packages\pymupdf\__init__.py", line 9, in save',
             '    raise FileDataError("cannot save")',
             'pymupdf.FileDataError: cannot save']
    assert log_view.stopped_at(trace) == {
        'error': 'pymupdf.FileDataError: cannot save', 'code': 'doc.save(output_pdf_path)',
        'where': 'engine/packing_labels/packing_label_generator.py, line 1520, in generate_packing_labels',
        'inside': 'pymupdf/__init__.py, line 9, in save'}
    assert log_view.stopped_at(['a line']) is None


def test_a_page_request_thread_is_not_background_work():
    assert not log_view.is_background('MainThread')
    assert not log_view.is_background('Thread-19 (process_request_thread)')  # the running app's page requests
    assert log_view.is_background('Thread-14 (work)')


def test_a_failed_page_is_pointed_at_its_run(log_dir):
    d = log_view.view(log_dir, focus_at='2026-10-18 16:24:00', focus_text='Stitch FAILED after', now=NOW)
    assert d['activities'][0]['title'] == 'Stitch FAILED after 0.10s' and d['focus'] == d['activities'][0]['id']
    gen = log_view.view(log_dir, focus_at='2026-10-18 16:23:45', focus_text='Generate done', now=NOW)
    assert gen['activities'][0]['result'].startswith('Generate done')  # found by a line inside the run
    assert log_view.view(log_dir, focus_at='2020-01-01 00:00:00', now=NOW)['focus'] is None
