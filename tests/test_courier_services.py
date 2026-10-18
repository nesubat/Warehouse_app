"""Courier services and their logos: the list (data/courier_services.json), the logo each label gets when stitched,
Back up & move with the logos, the logo sheet cut into files, and item boxes laid out the same in every box."""
import csv
import io
import json
import os
import sys

import pymupdf as fitz
import pytest
from PIL import Image, ImageDraw

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from engine.courier.courier_services import ServiceStore, ServiceError  # noqa: E402
from engine.core import data_transfer  # noqa: E402
from engine.packing_labels.label_stitcher import pack_services, draw_logos, logo_region, csv_kind  # noqa: E402
from engine.packing_labels.packing_label_generator import plan_cells, _layout_cell, CELL_W, CELL_H  # noqa: E402


def png(colour=(200, 0, 0), size=(400, 64)):
    buf = io.BytesIO()
    Image.new('RGBA', size, colour + (255,)).save(buf, 'PNG')
    return buf.getvalue()


@pytest.fixture()
def store(tmp_path):
    logos = tmp_path / "courier logos"
    logos.mkdir()
    (logos / "dhl.png").write_bytes(png((255, 204, 0)))  # DHL's logo, put in the folder by hand
    return ServiceStore(str(tmp_path / "courier_services.json"), str(logos))


def test_starting_list_has_the_groups_and_the_four_toll_priority_services(store):
    groups = {g['name']: [s['code'] for s in g['services']] for g in store.load()['groups']}
    assert groups['Toll Priority'] == ['20', '21', '22', '2']
    assert groups['Startrack Road Express'] == ['STEROAD', 'STERE2', 'STERET']
    assert groups['TOLL IPEC ROAD EXPRESS'] == ['IPECX'] and groups['TNT (NSW) ROAD EXPRESS'] == ['TNTN']
    assert {'BORDERP', 'BORDERB', 'DTRLOF', 'STNPRM', 'STNFPP', 'STNROAD', 'TNTNSWSPEC'} <= store.codes()
    assert len(store.codes()) == len(store.services())  # every code once


def test_add_edit_delete_and_codes_stay_unique(store):
    s = store.add_service('tolx', 'TOLL TEST', new_group='Toll Priority')
    assert s == {'code': 'TOLX', 'name': 'TOLL TEST', 'group': 'toll-priority', 'group_name': 'Toll Priority'}
    with pytest.raises(ServiceError):
        store.add_service('20', 'DUPLICATE')
    with pytest.raises(ServiceError):
        store.add_service('bad code!', 'X')
    store.update_service('TOLX', 'TOLY', 'TOLL TEST 2', 'dhl')  # moved to another group
    assert store.group_of('TOLY')['id'] == 'dhl'
    store.delete_service('TOLY')
    assert 'TOLY' not in store.codes()
    alone = store.add_service('NEW1', 'NEW COURIER')  # no group: a group of its own, named after it
    assert store.group_of('NEW1')['name'] == 'NEW COURIER' and alone['group_name'] == 'NEW COURIER'
    store.delete_service('NEW1')
    assert all(g['name'] != 'NEW COURIER' for g in store.load()['groups'])  # an empty group goes


def test_group_id_logo_in_the_folder_is_used_until_replaced_or_removed(store, tmp_path):
    folder = tmp_path / "courier logos"
    assert store.logo_for('DHLWPXOD') == str(folder / "dhl.png")
    assert store.logo_for('IPECX') is None  # no file for it
    store.set_logo('dhl', png((0, 0, 255)))
    assert os.path.dirname(store.logo_for('DHLWPXUD')) == str(folder)
    assert not (folder / "dhl.png").exists()  # replaced: nothing uses the old one, so it goes
    with pytest.raises(ServiceError):
        store.set_logo('dhl', b'not an image')
    store.delete_logo('dhl')
    assert store.logo_for('DHLWPXOD') is None  # removed on purpose: not even its dhl.png
    store.set_logo('dhl', png())
    assert store.logo_for('DHLWPXOD')


def test_deleting_a_group_keeps_its_services_each_on_its_own(store):
    store.delete_group('flash-couriers')
    assert store.group_of('FLASHFW')['name'] == 'FLASH COURIERS EXP WAGON'
    assert 'FLASHSW' in store.codes()


def test_back_up_and_move_carries_the_list_and_every_logo(store, tmp_path):
    from engine.addresses.address_book import AddressBook
    book = AddressBook(str(tmp_path / "book.db"))
    store.add_service('TOLX', 'TOLL TEST', new_group='Toll Priority')
    store.set_logo('toll-priority', png((0, 120, 0)))
    data, manifest = data_transfer.export_zip(str(tmp_path), book, ['courier_services'], store)
    assert 'courier_services' in manifest['items']
    other_dir = tmp_path / "other"
    other_dir.mkdir()
    other = ServiceStore(str(other_dir / "courier_services.json"), str(other_dir / "courier logos"))
    _, items = data_transfer.describe_zip(data)
    assert [i['key'] for i in items] == ['courier_services']
    result = data_transfer.import_zip(data, {'courier_services': 'replace'}, str(other_dir), book, None, other)
    assert result['done'][0][0] == 'Courier services'
    assert 'TOLX' in other.codes()
    assert other.logo_for('20') and other.logo_for('DHLWPXOD')  # the dropped logo and DHL's dhl.png both came


def test_service_of_each_label_comes_from_the_csv_as_saved(tmp_path):
    label_map = {
        'files': {'courier_csv': 'old.csv', 'courier_csvs': ['old.csv'], 'open360_csv': 'o360.csv'},
        'consignments': [{'number': 1, 'service_code': 'STEPRM'}, {'number': 2, 'service_code': 'STEPRM'}],
        'stores': {'Corio': [{'item_reference': 'Label 1 - Corio', 'open360_item_reference': '01T1P1 Corio', 'consignment': 1, 'box_pages': [2]}],
                   'Knox': [{'item_reference': 'Label 1 - Knox', 'open360_item_reference': '02T1P1 Knox', 'consignment': 2, 'box_pages': [3]}]},
    }
    with open(tmp_path / 'o360.csv', 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(['Receiver Name', 'Service', 'Item Reference'])
        w.writerow(['Corio', 'TNTN', '01T1P1 Corio'])  # changed outside the app
    assert csv_kind(['Receiver Name', 'Service', 'Item Reference']) == 'open360'
    assert pack_services(str(tmp_path), label_map) == {'Label 1 - Corio': 'TNTN', 'Label 1 - Knox': 'STEPRM'}


def test_logo_drawn_in_its_slot_on_each_box_first_page_once_in_the_file(tmp_path):
    logo = tmp_path / "logo.png"
    logo.write_bytes(png(size=(1000, 159)))
    doc = fitz.open()
    for _ in range(3):
        doc.new_page(width=842, height=595)
    label_map = {'logo_region_mm': {'x': 67.9, 'y': 199.1, 'width': 43.1, 'height': 6.84},
                 'stores': {'A': [{'item_reference': 'L1', 'box_pages': [1, 3]}], 'B': [{'item_reference': 'L2', 'box_pages': [2]}]}}
    assert draw_logos(doc, label_map, {'L1': str(logo), 'L2': None}) == 2
    region = logo_region(label_map)
    for n in (0, 2):
        (xref, *_), = doc[n].get_images()
        where = doc[n].get_image_rects(xref)[0]
        assert fitz.Rect(region.x0 - 0.01, region.y0 - 0.01, region.x1 + 0.01, region.y1 + 0.01).contains(where)
        assert abs(where.x1 - region.x1) < 0.5  # right edge of the slot
    assert not doc[1].get_images()
    assert len({img[0] for p in doc for img in p.get_images()}) == 1  # one image, reused


def test_label_layout_has_a_logo_slot_in_the_label_map(tmp_path):
    from engine.packing_labels.packing_label_generator import generate_packing_labels
    groups = {'G1': {'pack_spec_name': 'OB600100100', 'store_name': 'Corio Village', 'receiver_name': '', 'contact': '',
                     'install': False, 'address_1': 'Shop K001', 'address_2': '', 'suburb': 'Corio', 'state': 'VIC',
                     'postcode': '3214', 'country': '', 'excel_rows': [2],
                     'items': [{'job_no': 'J476699-01', 'qty': '2', 'desc': 'A1 poster', 'dimension': '', 'material': '',
                                'install': '', 'notes': '', 'thumbnails': []}]}}
    info = generate_packing_labels(groups, str(tmp_path / "l.pdf"), ['desc', 'job_no', 'barcode', 'qty'])
    r = info['logo_region_mm']
    assert abs(r['width'] - 43.1) < 0.1 and abs(r['x'] + r['width'] - 111) < 0.05  # right edge of the label block
    page = fitz.open(tmp_path / "l.pdf")[-1]
    chip = next(d for d in page.get_drawings() if d['rect'].height > 15 and d['rect'].y1 > 540)['rect']  # LABEL 1 OF 1's outline
    assert abs(chip.height - r['height'] * 72 / 25.4) < 1.6  # the logo slot is as tall as the chips
    assert abs(chip.y1 - (r['y'] + r['height']) * 72 / 25.4) < 1  # on the same line
    assert page.search_for("LABEL 1 OF 1")[0].x1 < r['x'] * 72 / 25.4  # and the chips stay clear of it


def test_item_boxes_share_one_layout_whatever_each_item_has():
    cell = fitz.Rect(0, 0, CELL_W, CELL_H)
    order = ['desc', 'job_no', 'barcode', 'qty']
    items = [{'job_no': 'J1-01', 'qty': '2', 'desc': 'A1 poster with a long description over two lines', 'barcode': 'J1-01'},
             {'job_no': 'J1-02', 'qty': '4', 'desc': '', 'barcode': 'J1-02'}]  # no description
    ts, th, slots = plan_cells(items, order, cell)
    assert slots['desc'] > 0
    tops = []
    for it in items:  # where the job number bar starts in each box
        page = fitz.open().new_page(width=CELL_W, height=CELL_H)
        _layout_cell(it, order, cell, ts, th, draw_page=page, slots=slots)
        tops.append(round(page.search_for(it['job_no'])[0].y0, 1))
    assert tops[0] == tops[1]


def test_logo_sheet_is_cut_into_one_file_per_group(tmp_path):
    from engine.courier.logo_sheet import cut, ORDER, LOGO_WIDTH
    sheet = Image.new('RGB', (1150, 470), 'white')
    d = ImageDraw.Draw(sheet)
    for gid, c, r in ORDER:
        x0, y0 = 6 + c * 572, 6 + r * 92
        fill = (255, 255, 255) if gid == 'fedex-international' else (20 * r + 40, 100, 30 * c + 60)
        d.rectangle([x0, y0, x0 + 555, y0 + 86], fill=fill, outline=(200, 200, 200))
        d.text((x0 + 200, y0 + 35), gid, fill=(0, 0, 0))
    path = tmp_path / "sheet.png"
    sheet.save(path)
    made = cut(str(path), str(tmp_path / "logos"))
    assert [m[0] for m in made] == [gid for gid, _, _ in ORDER]
    for gid, out, size, nbytes in made:
        assert size[0] == LOGO_WIDTH and abs(size[1] - LOGO_WIDTH * 87 / 556) < 4 and nbytes < 80_000  # its own proportions
    dhl = Image.open(tmp_path / "logos" / "dhl.png").convert('RGB')
    assert dhl.getpixel((LOGO_WIDTH // 5, dhl.height // 2)) == (40, 100, 60)  # its own colour, unchanged


def logo_files(store):
    return sorted(os.listdir(store.logo_dir)) if os.path.isdir(store.logo_dir) else []


def test_a_service_takes_the_logo_of_the_group_it_is_moved_to(store):
    store.set_logo('toll-priority', png((0, 90, 0)))
    store.set_logo('tnt-overnight', png((230, 100, 0)))
    toll, tnt = store.logo_for('20'), store.logo_for('TNTNONFC')
    store.update_service('20', '20', 'TOLL PRIORITY 1KG SATCHEL', 'tnt-overnight')  # moved to another group
    assert store.logo_for('20') == tnt and store.logo_for('21') == toll  # the group keeps its own
    store.update_service('20', '20', 'TOLL PRIORITY 1KG SATCHEL', '__own__')  # out on its own: no logo yet
    assert store.group_of('20')['name'] == 'TOLL PRIORITY 1KG SATCHEL' and store.logo_for('20') is None
    assert store.logo_for('TNTNONFC') == tnt


def test_a_replaced_logo_file_goes_unless_another_group_still_uses_it(store):
    store.set_logo('dhl', png((1, 2, 3)))
    first = logo_files(store)
    store.set_logo('fedex-international', png((1, 2, 3)))  # the same logo: one file for both groups
    assert logo_files(store) == first
    store.set_logo('dhl', png((9, 9, 9)))  # replaced: the old file stays, FedEx still uses it
    assert len(logo_files(store)) == 2 and first[0] in logo_files(store)
    store.delete_logo('fedex-international')  # now nothing uses it: gone
    assert first[0] not in logo_files(store) and len(logo_files(store)) == 1
    store.delete_group('dhl')  # its services stay, on their own, and its logo file goes
    assert logo_files(store) == [] and 'DHLWPXOD' in store.codes()


def test_export_holds_only_the_logos_in_use_each_once(store, tmp_path):
    store.set_logo('flash-couriers', png((5, 5, 5)))
    store.set_logo('tnt-international', png((5, 5, 5)))  # shared
    os.makedirs(store.logo_dir, exist_ok=True)
    with open(os.path.join(store.logo_dir, 'stray.txt'), 'w') as f:
        f.write('not a logo')
    files = store.export_files()
    logos = [k for k in files if k.startswith('courier_logos/')]
    data = json.loads(files['courier_services.json'])
    linked = {g['logo'] for g in data['groups'] if g['logo']}
    assert len(logos) == 2  # the shared one once, and DHL's dhl.png
    assert {k.split('/', 1)[1] for k in logos} == linked


def test_group_id_logos_go_with_their_group_and_other_files_stay(store, tmp_path):
    folder = tmp_path / "courier logos"
    (folder / "toll-priority.png").write_bytes(png((0, 90, 0)))
    (folder / "spare logo.png").write_bytes(png((1, 1, 1)))  # no group's id: never the store's to delete
    assert store.logo_for('22') == str(folder / "toll-priority.png")
    store.delete_group('toll-priority')  # its services stay, each on its own, and its logo goes
    assert not (folder / "toll-priority.png").exists() and '22' in store.codes()
    store.delete_logo('dhl')  # removed on purpose: the file goes too
    assert not (folder / "dhl.png").exists() and store.logo_for('DHLWPXOD') is None
    store.add_service('NEW1', 'NEW COURIER')
    assert (folder / "spare logo.png").exists()


def test_logos_in_the_old_folder_move_to_the_new_one_on_first_use(tmp_path):
    old, new = tmp_path / "courier_logos", tmp_path / "courier logos"
    old.mkdir()
    new.mkdir()
    (old / "logo-0123456789abcdef.png").write_bytes(png((3, 3, 3)))  # dropped before the move
    (old / "dhl.png").write_bytes(png((4, 4, 4)))
    (new / "dhl.png").write_bytes(png((5, 5, 5)))  # already moved by hand: this one is kept
    path = tmp_path / "courier_services.json"
    data = ServiceStore(str(path), str(new)).load()
    next(g for g in data['groups'] if g['id'] == 'tnt-overnight')['logo'] = 'logo-0123456789abcdef.png'
    path.write_text(json.dumps(data), encoding='utf-8')
    store = ServiceStore(str(path), str(new), legacy_dir=str(old))
    assert old.exists()  # nothing moves until it's used
    assert store.logo_for('TNTNONFC') == str(new / "logo-0123456789abcdef.png")
    assert Image.open(new / "dhl.png").convert('RGB').getpixel((0, 0)) == (5, 5, 5)
    assert sorted(os.listdir(old)) == ['dhl.png']  # the one it couldn't move stays where it was


def test_adding_a_code_from_another_group_moves_it_here(store):
    """+ Add in a group's card with a code that's in another group: moved here, its name kept; that group, emptied, goes."""
    moved = store.add_service('BORDERB', '', 'tnt-nsw-road-express', move=True)
    assert moved['moved_from'] == 'BORDER EXPRESS BULK' and moved['group'] == 'tnt-nsw-road-express'
    assert store.group_of('BORDERB')['id'] == 'tnt-nsw-road-express'
    assert next(s for s in store.services() if s['code'] == 'BORDERB')['name'] == 'BORDER EXPRESS BULK'
    assert all(g['name'] != 'BORDER EXPRESS BULK' for g in store.load()['groups'])
    with pytest.raises(ServiceError):
        store.add_service('BORDERB', '', 'tnt-nsw-road-express', move=True)  # already here
    with pytest.raises(ServiceError):
        store.add_service('TNTN', 'X', 'dhl')  # without move (the preview's + Add service): a duplicate stays an error


LABEL_MAP = {
    'files': {'courier_csvs': ['o360.csv', 'of.csv']},
    'consignments': [{'number': 1, 'service_code': 'STEPRM',
                      'destination': {'receiver': 'Corio Village', 'line1': 'Shop K001', 'line2': '', 'suburb': 'Corio',
                                      'state': 'VIC', 'postcode': '3214', 'country': ''}}],
    'stores': {'Corio Village': [{'item_reference': 'Label 1 - Corio Village', 'open360_item_reference': '01T1P1 Corio Village',
                                  'consignment': 1, 'box_pages': [2]}]},
}


def _write_csvs(folder, o360_service='STEPRM', of_service='TNTN', of_suburb='Corio'):
    with open(folder / 'o360.csv', 'w', newline='', encoding='utf-8-sig') as f:
        csv.writer(f).writerows([['Receiver Name', 'Receiver Suburb', 'Service', 'Item Reference'],
                                 ['Corio Village', 'Corio', o360_service, '01T1P1 Corio Village']])
    with open(folder / 'of.csv', 'w', newline='', encoding='cp1252') as f:  # saved by Excel
        csv.writer(f).writerows([['Reference', 'Receiver Name', 'Receiver Address Line 1', 'Receiver Suburb', 'Receiver State',
                                  'Receiver Postcode', 'Service Code', 'Reference'],
                                 ['J1', 'Corio Village', 'Shop K001', of_suburb, 'VIC', '3214', of_service, '01T1P1 Corio Village']])


def test_the_picked_csv_alone_decides_the_service_openfreight_too(tmp_path):
    """The file picked on the Stitch page decides: an OpenFreight CSV is matched by its last Reference column (the
    Open360 reference), and the Open360 one no longer overrides it."""
    _write_csvs(tmp_path)
    assert pack_services(str(tmp_path), LABEL_MAP, 'of.csv') == {'Label 1 - Corio Village': 'TNTN'}
    assert pack_services(str(tmp_path), LABEL_MAP, 'o360.csv') == {'Label 1 - Corio Village': 'STEPRM'}


def test_an_address_changed_in_the_csv_is_reported(tmp_path):
    from engine.packing_labels.label_stitcher import address_changes
    _write_csvs(tmp_path, of_suburb='Norlane')
    (m,) = address_changes(str(tmp_path), LABEL_MAP, 'of.csv')
    assert m['number'] == 1 and m['stores'] == ['Corio Village']
    assert 'Corio, VIC, 3214' in m['labels'] and 'Norlane, VIC, 3214' in m['now']
    _write_csvs(tmp_path)  # the same place, written differently (case, AU left blank): no change
    assert address_changes(str(tmp_path), LABEL_MAP, 'of.csv') == []


def _two_computers(tmp_path):
    """PC A with changes made on it, PC B with nothing saved yet (its folder has the carriers' dhl.png, as a clone has)."""
    from engine.addresses.address_book import AddressBook
    from engine.courier.packing_specs import SpecStore
    a, b = tmp_path / "pc A data", tmp_path / "pc B data"
    for d in (a, b):
        (d / "courier logos").mkdir(parents=True)
        (d / "courier logos" / "dhl.png").write_bytes(png((255, 204, 0)))
    store_a = ServiceStore(str(a / "courier_services.json"), str(a / "courier logos"))
    store_a.update_service('TNTN', 'TNTN', 'TNT (NSW) ROAD EXPRESS', 'tnt-overnight')  # moved
    store_a.rename_group('toll-priority', 'Toll Priority Express')
    store_a.set_logo('toll-priority', png((0, 90, 0)))
    store_a.add_group('Couriers Please')  # no services yet
    store_a.delete_logo('fedex-international')  # removed on purpose
    with open(a / "packing_specs.json", "w", encoding="utf-8") as f:
        json.dump({'formula': {}, 'named': [{'name': 'A4 box', 'length': 32, 'width': 23, 'height': 19, 'weight': 1.2,
                                             'item_type': 'Carton'}]}, f)
    store_b = ServiceStore(str(b / "courier_services.json"), str(b / "courier logos"))
    return (a, AddressBook(str(a / "book.db")), store_a), (b, AddressBook(str(b / "book.db")), store_b, SpecStore(str(b / "packing_specs.json")))


def test_a_backup_merged_on_a_new_computer_brings_everything(tmp_path):
    """Merge (the page's choice) on a computer with nothing saved yet takes the backup as it is: moved services, a
    renamed group and its logo, an empty group, a removed logo and the packing specs all come across."""
    (a, book_a, store_a), (b, book_b, store_b, specs_b) = _two_computers(tmp_path)
    data, _ = data_transfer.export_zip(str(a), book_a, ['courier_services', 'packing_specs'], store_a)
    result = data_transfer.import_zip(data, {'courier_services': 'merge', 'packing_specs': 'merge'}, str(b), book_b,
                                      specs_b, store_b)
    assert all('taken as' in what for _, what in result['done'])
    assert store_b.group_of('TNTN')['id'] == 'tnt-overnight'
    toll = store_b.group_of('20')
    assert (toll['id'], toll['name']) == ('toll-priority', 'Toll Priority Express') and store_b.logo_for('21')
    assert any(g['name'] == 'Couriers Please' for g in store_b.load()['groups'])
    assert store_b.logo_for('FXIEOIE') is None  # removed on A, stays removed
    assert store_b.logo_for('DHLWPXOD') == str(b / "courier logos" / "dhl.png")  # the same picture: left as it is
    assert specs_b.load()['named'][0]['length'] == 32


def test_two_imports_in_the_same_second_each_keep_a_backup(tmp_path):
    (a, book_a, store_a), (b, book_b, store_b, specs_b) = _two_computers(tmp_path)
    data, _ = data_transfer.export_zip(str(a), book_a, ['courier_services'], store_a)
    first = data_transfer.import_zip(data, {'courier_services': 'replace'}, str(b), book_b, specs_b, store_b)
    second = data_transfer.import_zip(data, {'courier_services': 'replace'}, str(b), book_b, specs_b, store_b)
    assert first['backup'] != second['backup'] and os.path.isdir(second['backup'])


def test_no_logo_on_a_packing_label_that_got_no_courier_label(tmp_path):
    """The logo names the courier of the label beside it: a box the stitch left without a courier label (Macpac Kent
    St in a real run) gets no logo, the others do."""
    logo = tmp_path / "logo.png"
    logo.write_bytes(png(size=(1000, 159)))
    doc = fitz.open()
    for _ in range(3):
        doc.new_page(width=842, height=595)
    label_map = {'logo_region_mm': {'x': 67.9, 'y': 199.1, 'width': 43.1, 'height': 6.84},
                 'stores': {'A': [{'item_reference': 'L1', 'box_pages': [1, 3]}], 'Kent St': [{'item_reference': 'L2', 'box_pages': [2]}]}}
    assert draw_logos(doc, label_map, {'L1': str(logo), 'L2': str(logo)}, pages={1, 3}) == 2
    assert doc[0].get_images() and doc[2].get_images() and not doc[1].get_images()
