"""Horizontal distribution labels: reading, checks, packing-sheet order, references, the PDF, stitching codes,
and the whole Packing Labels page (upload -> preview -> generate) through Flask's test client.

    pip install -r requirements-dev.txt
    pytest -q
"""
import copy
import json
import os
import random
import re
import sys

import pymupdf as fitz
import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from fixtures import build_final_file, build_signature_links  # noqa: E402
from horizontal_distribution import read_horizontal, read_signature_file, item_references  # noqa: E402
from packing_label_generator import PackCheckError, assign_barcodes  # noqa: E402

FINAL = "Final J476699 Spring rollout.xlsx"
SIG = "Signature links_J476699.xlsx"


@pytest.fixture()
def files(tmp_path):
    codes = build_final_file(tmp_path / FINAL)
    build_signature_links(tmp_path / SIG, codes)
    return tmp_path / FINAL, tmp_path / SIG, codes


def read(path, tab="Spring Sale", sig=None, tab_no=1):
    return read_horizontal(str(path), 8, tab, None, tab_no, read_signature_file(str(sig)) if sig else None)


def test_codes_read_from_code_columns_in_packing_sheet_order(files):
    final, sig, _ = files
    groups, _, found, warnings = read(final, sig=sig)
    order = [(g["pack_no"], g["pack_code"], g["store_name"]) for g in groups.values()]
    assert order[:3] == [(1, "A", "Westfield Parramatta"), (1, "A", "Carindale"), (1, "B", "Corio Village")]
    assert [p for p, _, _ in order] == sorted(p for p, _, _ in order)  # pack by pack
    assert found["plugin"] and not warnings


def test_reference_is_store_number_tab_pack(files):
    final, sig, _ = files
    groups, *_ = read(final, sig=sig)
    refs = dict(item_references(groups))
    assert "03T1P2 Chadstone" in refs and refs["03T1P2 Chadstone"] == "03T1P2"
    easter, *_ = read(final, "Easter", sig, tab_no=2)
    assert {g["serial"] for g in easter.values()} == {"01T2P1", "02T2P1", "03T2P1", "04T2P1"}


def test_subgroup_codes_kept_and_shown(files):
    final, sig, _ = files
    groups, *_ = read(final, sig=sig)
    chad = next(g for g in groups.values() if g["serial"] == "03T1P2")
    assert chad["sub_codes"] == [("1-2", "C")] and "1-2: C" in chad["pack_tag"]


def test_without_code_columns_codes_come_from_signature_links_then_quantities(tmp_path):
    codes = build_final_file(tmp_path / FINAL, codes=False)
    build_signature_links(tmp_path / SIG, codes)
    with_sig, *_ = read(tmp_path / FINAL, sig=tmp_path / SIG)
    alone, _, found, _ = read(tmp_path / FINAL)
    assert [g["pack_code"] for g in with_sig.values()] == [g["pack_code"] for g in alone.values()]
    assert "worked out" in found["codes"][0]


def test_codes_differing_from_signature_links_stop_generate(tmp_path):
    codes = build_final_file(tmp_path / FINAL)
    build_signature_links(tmp_path / SIG, codes, tamper=("Spring Sale", "1", 0, "C"))
    with pytest.raises(PackCheckError) as e:
        read(tmp_path / FINAL, sig=tmp_path / SIG)
    assert any("Signature Links" in i["text"] for i in e.value.issues)


def test_quantities_changed_after_packing_sheets_stop_generate(tmp_path):
    build_final_file(tmp_path / FINAL, change=("Spring Sale", 0, 0, 7))
    with pytest.raises(PackCheckError) as e:
        read(tmp_path / FINAL)
    assert any("doesn't match the quantities" in i["text"] for i in e.value.issues)


def test_every_unit_lands_on_a_label(files):
    final, sig, _ = files
    groups, *_ = read(final, sig=sig)
    from openpyxl import load_workbook
    ws = load_workbook(final, data_only=True)["Spring Sale"]
    jobs = [c for c in range(1, ws.max_column + 1) if re.match(r"J\d", str(ws.cell(row=7, column=c).value or ""))]
    grid = sum(ws.cell(row=r, column=c).value or 0 for r in range(9, 17) for c in jobs)
    assert grid == sum(float(i["qty"]) for g in groups.values() for i in g["items"])


def test_barcode_kinds_follow_item_columns(files):
    final, sig, _ = files
    groups, *_ = read(final, sig=sig)
    assign_barcodes(groups)
    kinds = {i["barcode"] for g in groups.values() for i in g["items"] if i["job_no"] == "J476699-03"}
    assert kinds == {"J476699-03 K1", "J476699-03 K2"}  # two columns share the job: two kinds, on every store


def test_barcode_kinds_of_vertical_files_unchanged():
    """No source_key (vertical files): every row is its own kind, as before."""
    rnd = random.Random(4)
    for _ in range(200):
        groups = {f"g{i}": {"items": [{"job_no": rnd.choice(["J1-01", "J1-02", "j1-02 ", ""])} for _ in range(rnd.randint(0, 5))]}
                  for i in range(rnd.randint(1, 6))}
        expected = copy.deepcopy(groups)
        seen, totals = {}, {}
        for g in expected.values():
            for it in g["items"]:
                k = it["job_no"].strip().upper()
                if k:
                    totals[k] = totals.get(k, 0) + 1
        for g in expected.values():
            for it in g["items"]:
                k, job = it["job_no"].strip().upper(), it["job_no"].strip()
                if not k:
                    it["barcode"] = ""
                    continue
                seen[k] = seen.get(k, 0) + 1
                it["barcode"] = f"{job} K{seen[k]}" if totals[k] > 1 else job
        assign_barcodes(groups)
        assert groups == expected


def test_stitcher_reads_horizontal_codes():
    from label_stitcher import find_code
    known = {"03T1P2", "12T2P1"}
    assert find_code("Item Ref: 03T1P2 Chadstone", known) == "03T1P2"
    assert find_code("Item Reference 12T2P1 Rundle Mall", known) == "12T2P1"
    assert find_code("Item Ref: 103T1P2 Chadstone", known) is None
    assert find_code("Item Ref: 03T1P21 Chadstone", known) is None


def test_signature_links_tabs_skipped_by_shuffler():
    from core_math import NOT_STORE_TABS, LABEL_PLUGIN_SHEET, DIVIDER_BARCODE_SHEET
    assert set(NOT_STORE_TABS) == {LABEL_PLUGIN_SHEET, DIVIDER_BARCODE_SHEET}


# ------------------------------------------------------------------ the page, end to end

@pytest.fixture()
def client(tmp_path, monkeypatch):
    import app as app_module
    monkeypatch.setitem(app_module.app.config, "UPLOAD_FOLDER", str(tmp_path / "projects"))
    os.makedirs(tmp_path / "projects", exist_ok=True)
    for name in ("close_if_open_elsewhere", "save_if_open_elsewhere"):  # Excel-only helpers
        monkeypatch.setattr(app_module, name, lambda *a, **k: None)
    monkeypatch.setattr(app_module, "convert_legacy_excel_to_xlsx", lambda p: p)
    app_module.app.config["TESTING"] = True
    return app_module.app.test_client(), app_module, tmp_path / "projects"


def run_page(client, final, sig, dividers=True):
    c, app_module, projects = client
    with open(final, "rb") as f, open(sig, "rb") as s:
        r = c.post("/packing-labels/horizontal", data={"file": (f, final.name), "signature_file": (s, sig.name)},
                   content_type="multipart/form-data", follow_redirects=True)
    assert r.status_code == 200, r.data[:500]
    page = r.data.decode()
    assert 'value="8"' in page  # header rows filled in from the Label Plugin
    filename = app_module.secure_filename(final.name)
    form = {"filename": filename, "all_tabs": ["Spring Sale", "Easter"], "selected_tabs": ["Spring Sale", "Easter"],
            "header_row_Spring Sale": "8", "header_row_Easter": "8"}
    r = c.post("/packing-labels/horizontal", data={**form, "preview": "1"}, follow_redirects=True)
    preview = r.data.decode()
    assert "Cannot Generate" not in preview, preview[preview.find("issue"):][:800]
    assert "Signature Links loaded" in preview
    from courier_export import SERVICE_CODES
    service = SERVICE_CODES[0][0]
    gen = {**form, "generate": "1", "who_pays": "S", "service_code": service, "reference": "J476699",
           "project_name": "Spring rollout", "courier_csv": ["open360", "openfreight"],
           "attribute_order": "thumbnail,desc,dimension,job_no,barcode,qty"}
    if dividers:
        gen["code_dividers"] = "1"
    r = c.post("/packing-labels/horizontal", data=gen, follow_redirects=True)
    assert r.status_code == 200 and "Labels Successfully Generated" in r.data.decode(), r.data.decode()[:1500]
    folder = next(p for p in projects.iterdir() if p.is_dir())
    return folder


def test_page_generates_labels_csv_and_label_map(client, files):
    final, sig, _ = files
    folder = run_page(client, final, sig)
    names = os.listdir(folder)
    pdf = next(n for n in names if n.startswith("Packing Labels Only"))
    open360 = next(n for n in names if n.endswith("Open360.csv"))
    labelmap = json.load(open(folder / next(n for n in names if n.endswith(".labelmap.json"))))
    assert any(n.startswith("Used Signature Links") for n in names)
    rows = open(folder / open360, encoding="utf-8").read()
    assert "03T1P2 Chadstone" in rows and "01T2P1 Corio Village" in rows
    doc = fitz.open(folder / pdf)
    texts = [p.get_text() for p in doc]
    # report, then per pack: audit page, divider per code, labels
    assert "PACK AUDIT - TAB 1 - PACK 1" in texts[labelmap.get("report_pages", 1)]
    codes = {pk["open360_code"] for packs in labelmap["stores"].values() for pk in packs}
    assert "03T1P2" in codes
    for packs in labelmap["stores"].values():  # every label-map page really is that pack's packing label
        for pk in packs:
            for p in pk["pdf_pages"] + pk["box_pages"]:
                assert pk["open360_code"] + "." in texts[p - 1], (pk["open360_code"], texts[p - 1][:200])
    chad = next(pk for pk in labelmap["stores"]["Chadstone"] if pk["open360_code"] == "03T1P2")
    assert chad["boxes"] == 2 and len(set(chad["box_pages"])) == 2  # one courier label per box


def test_page_without_dividers_has_no_divider_sheets(client, files):
    final, sig, _ = files
    folder = run_page(client, final, sig, dividers=False)
    pdf = next(n for n in os.listdir(folder) if n.startswith("Packing Labels Only"))
    assert not any("Matched Labels" in p.get_text() for p in fitz.open(folder / pdf))


def test_cross_site_actions_refused(client):
    c, *_ = client
    assert c.post("/packing-labels/horizontal", headers={"Sec-Fetch-Site": "cross-site"}).status_code == 403
    assert c.get("/open_upload/x.xlsx", headers={"Sec-Fetch-Site": "cross-site"}).status_code == 403
    assert c.post("/packing-labels/horizontal", headers={"Origin": "http://evil.example"}).status_code == 403
    assert c.get("/packing-labels/horizontal", headers={"Sec-Fetch-Site": "cross-site"}).status_code == 200  # just a page


def test_courier_labels_stitch_onto_their_packing_labels(client, files):
    """Courier labels booked from the Open360 CSV (Item Ref '03T1P2 Chadstone'), in any order, land on their own
    packing label; a pack of two boxes takes two."""
    import csv
    from label_stitcher import stitch
    final, sig, _ = files
    folder = run_page(client, final, sig)
    names = os.listdir(folder)
    labelmap = json.load(open(folder / next(n for n in names if n.endswith(".labelmap.json"))))
    rows = list(csv.DictReader(open(folder / next(n for n in names if n.endswith("Open360.csv")), encoding="utf-8")))
    refs = [r["Item Reference"] for r in rows for _ in range(int(r["Item Quantity"] or 1))]
    assert all(re.match(r"\d{2}T\dP\d ", ref) for ref in refs)
    random.Random(7).shuffle(refs)
    courier = fitz.open()
    for ref in refs:  # a 4x6 courier label with its Item Ref
        page = courier.new_page(width=288, height=432)
        page.insert_text((20, 40), "TIG Freight - Open360", fontsize=12)
        page.insert_text((20, 300), f"Item Ref: {ref}", fontsize=10)
    out = folder / "Complete.pdf"
    result = stitch(str(folder), labelmap, [("courier.pdf", courier.tobytes())], str(out))
    assert not result["unmatched"] and len(result["placed"]) == len(refs)
    assert fitz.open(out).page_count >= len(refs)


def test_pack_tag_leaves_packing_spec_full_size(tmp_path):
    """The PACK/CODE tag must not shrink the Packing Spec printed under it (they once shared a variable)."""
    from packing_label_generator import generate_packing_labels
    group = lambda tag: {"g": {"pack_spec_name": "2 x CS450320150", "store_name": "Chadstone", "receiver_name": "", "install": False,
                               "address_1": "", "address_2": "", "suburb": "", "state": "", "postcode": "", "country": "",
                               "pack_tag": tag, "excel_rows": [9], "items": [{"job_no": "J1-01", "qty": "1", "desc": "", "dimension": "",
                                                                              "material": "", "install": "", "notes": "", "thumbnails": []}]}}
    sizes = []
    for tag in ("", "PACK 2 \u00b7 CODE C \u00b7 1-2: C"):
        out = tmp_path / f"l{len(sizes)}.pdf"
        generate_packing_labels(group(tag), str(out))
        spans = [s for b in fitz.open(out)[0].get_text("dict")["blocks"] for line in b.get("lines", []) for s in line["spans"]]
        sizes.append(max(s["size"] for s in spans if "CS450320150" in s["text"]))
    assert sizes[0] == sizes[1] >= 18


def test_labels_carry_code_boxes_and_audit_page_is_short(client, files):
    """Each label has an item box per code: the code column's header in the job number's place, the store's code
    in the quantity's, no barcode. Code boxes never count as job numbers. The audit page has no code table."""
    final, sig, _ = files
    folder = run_page(client, final, sig)
    names = os.listdir(folder)
    doc = fitz.open(folder / next(n for n in names if n.startswith("Packing Labels Only")))
    labelmap = json.load(open(folder / next(n for n in names if n.endswith(".labelmap.json"))))
    chad = next(pk for pk in labelmap["stores"]["Chadstone"] if pk["open360_code"] == "03T1P2")
    page = " ".join(doc[chad["pdf_pages"][0] - 1].get_text().split())  # a long header wraps in its box
    assert page.count("Code for Pack 2") == 1 and "Code for pack 2: 1-2" in page  # header once: no barcode caption
    words = [w[4] for w in doc[chad["pdf_pages"][0] - 1].get_text("words")]
    assert words.count("C") >= 2  # the pack code and the sub-group code, each in its quantity box
    assert all(not str(j).startswith("Code for") for packs in labelmap["stores"].values() for pk in packs
               for j in pk.get("job_numbers", []))
    audit = next(p.get_text() for p in doc if "PACK AUDIT - TAB 1 - PACK 2" in p.get_text())
    assert "J476699-05" not in audit and "Spring Sale" not in audit and "LABELS" in audit


def test_code_box_named_after_pack_when_codes_come_from_signature_links(tmp_path):
    codes = build_final_file(tmp_path / FINAL, codes=False)
    build_signature_links(tmp_path / SIG, codes)
    groups, *_ = read(tmp_path / FINAL, sig=tmp_path / SIG)
    g = next(iter(groups.values()))
    assert [c["job_no"] for c in g["code_cells"]] == ["Code for Pack 1"] and g["code_cells"][0]["qty"] == g["pack_code"]
    assert all(not i.get("is_code") for i in g["items"])


def test_found_names_the_cell_layout_blocks(files):
    """The page's cell layout offers a block only when read_horizontal() says its details were found, by the same
    names parse_packing_data() uses. Without them every slot showed "Empty Field" for horizontal files."""
    final, sig, _ = files
    _, _, found, _ = read(final, sig=sig)
    for name in ("Job Number", "Quantity", "Description", "Dimensions", "Thumbnail", "Install",
                 "Packing Spec", "Store Name", "Address Info"):
        assert found[name], name
    assert not found["Material"] and not found["Notes"]  # no Material or Notes row in the file


def test_page_offers_layout_blocks_and_previews_code_box_first(client, files):
    final, sig, _ = files
    c, app_module, _ = client
    with open(final, "rb") as f, open(sig, "rb") as s:
        c.post("/packing-labels/horizontal", data={"file": (f, final.name), "signature_file": (s, sig.name)},
               content_type="multipart/form-data", follow_redirects=True)
    form = {"filename": app_module.secure_filename(final.name), "all_tabs": ["Spring Sale"],
            "selected_tabs": ["Spring Sale"], "header_row_Spring Sale": "8", "preview": "1"}
    page = c.post("/packing-labels/horizontal", data=form, follow_redirects=True).data.decode()
    cell = page[page.index('id="attribute-list"'):page.index('id="attribute_order"')]
    for block in ("thumbnail", "desc", "dimension", "job_no", "barcode", "qty"):
        assert f'data-id="{block}"' in cell, block
    assert 'data-id="empty"' not in cell  # no "Empty Field" slots
    assert 'data-id="install"' in page[page.index('id="mapped-headers-list"'):page.index('id="attribute-list"')]
    preview = json.loads(re.search(r"data-cells='(.*?)'", page, re.DOTALL).group(1))
    first, item = preview["items"][0], preview["items"][1]
    assert first["job_no"].startswith("Code for Pack") and not first["modules"]  # the code box: no barcode
    assert item["job_no"].startswith("J476699") and item["modules"]
