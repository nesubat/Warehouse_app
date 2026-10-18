"""Live progress for the loading overlay (engine/core/progress.py)."""
import logging
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from engine.core import progress  # noqa: E402


def test_nothing_happens_without_a_run():
    progress.finish()
    progress.stage("Drawing", 10, 90)  # outside a followed request: no error, nothing kept
    progress.tick(1, 2)
    assert progress.current() is None
    assert progress.start("bad id!") is None and progress.start("") is None


def test_a_step_fills_its_share_of_the_bar_and_never_goes_back():
    run = progress.start("unit-run-0001")
    try:
        progress.stage("Reading", 2, 15)
        progress.tick(1, 2)
        assert progress.get(run)['percent'] == 8.5
        progress.stage("Drawing", 20, 90)
        progress.tick(3, 4, "Drawing 4 of 4")
        assert progress.get(run)['percent'] == 72.5 and progress.get(run)['step'] == "Drawing 4 of 4"
        progress.tick(1, 4)  # a late, smaller tick doesn't move it back
        assert progress.get(run)['percent'] == 72.5
        log = logging.getLogger('unit_progress_test')  # not under 'wh': nothing reaches logs/warehouse.log
        log.propagate = False
        log.setLevel(logging.DEBUG)
        log.addHandler(progress.Handler())
        log.warning("  tab VIC: 1 problem")
        line = progress.get(run)['lines'][-1]
        assert line == {**line, 'level': 'WARNING', 'depth': 1, 'text': 'tab VIC: 1 problem'}
    finally:
        progress.finish()
    assert progress.get(run)['finished'] and progress.get(run)['percent'] == 100


def test_work_handed_to_a_background_thread_keeps_its_run():
    """The matching page's Finalise: the request detaches, the stitch thread attaches and ends the run."""
    import threading
    run = progress.start("handoff-run-0001")
    assert progress.detach() == run and progress.current() is None
    progress.finish()  # the request ending: nothing happens to the run
    assert not progress.get(run)['finished']

    def work():
        progress.attach(run)
        progress.stage("Rendering the courier labels", 10, 80)
        progress.tick(1, 2)
        progress.finish()
    t = threading.Thread(target=work)
    t.start()
    t.join()
    d = progress.get(run)
    assert d['finished'] and d['step'] == "Rendering the courier labels"


def test_a_script_sends_its_progress_id_in_a_header():
    import app as app_module
    client = app_module.app.test_client()
    client.post('/stitch/not-a-project/finalise', json={}, headers={'X-Progress-Id': 'header-run-0001'})
    assert progress.get('header-run-0001')['finished']


def test_the_label_shuffler_logs_what_it_found_and_counts_pages(tmp_path):
    import pymupdf as fitz
    from engine.label_shuffler.pdf_engine import process_and_shuffle_pdf
    src = fitz.open()
    for store in ("Corio Village", "Knox City", "Somewhere Else"):
        src.new_page(width=288, height=432).insert_text((20, 40), f"Deliver to {store}")
    src.save(tmp_path / "labels.pdf")
    run = progress.start("shuffle-run-0001")
    try:
        progress.stage("Shuffling labels.pdf", 5, 95)
        process_and_shuffle_pdf(str(tmp_path / "labels.pdf"), {"Corio Village": "A", "Knox City": "B", "Highpoint": "B"},
                                str(tmp_path / "out.pdf"), signature_header="Pack 1")
        d = progress.get(run)
    finally:
        progress.finish()
    texts = [x['text'] for x in d['lines']]
    assert any(t.startswith("shuffling labels.pdf by the codes of 'Pack 1'") for t in texts)
    assert any("2 of 3 store(s) matched" in t for t in texts)
    assert any(t == "needs a look: code B, no label for Highpoint" for t in texts)
    assert d['percent'] > 5 and "page" in d['step']


def test_a_render_worker_writes_how_many_labels_it_has_done(tmp_path):
    import pymupdf as fitz
    from engine.packing_labels.label_render import _render_pages, _count
    doc = fitz.open()
    for _ in range(2):
        doc.new_page(width=288, height=432).insert_text((20, 40), "TO: Corio Village")
    doc.save(tmp_path / "c.pdf")
    done = tmp_path / "done.txt"
    out = _render_pages({"c.pdf": str(tmp_path / "c.pdf")}, [("c.pdf", 1), ("c.pdf", 2)], (300, 400), str(done))
    assert len(out) == 2 and _count(str(done)) == 2 and _count(str(tmp_path / "none.txt")) == 0
