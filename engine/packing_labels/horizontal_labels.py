"""The packing labels PDF for a horizontal distribution file, in packing-sheet order:

  job report (page 1)
  for each tab, for each pack:   audit page   - the pack's codes as on its packing sheet, with stores, boxes and
                                                quantities per item, where the codes came from, and what needs a look
      for each code A, B, C …:   divider      - the Label Shuffler's divider sheet (pdf_engine.build_divider_sheet),
                                                without barcodes: every item on the labels has its own
                                 labels       - the standard A4 packing labels (generate_packing_labels), one per box

The labels are drawn by generate_packing_labels() exactly as for vertical files, then put in this order. The page
numbers returned (and written to the label map) point at the final PDF, so Stitch Labels finds every label.
"""
import os
import tempfile
from collections import Counter

import pymupdf as fitz

from engine.packing_labels.packing_label_generator import generate_packing_labels, draw_report_pages, A4_W, A4_H, _latin, \
    draw_audit_page
from engine.label_shuffler.pdf_engine import build_divider_sheet
from engine.packing_labels.horizontal_distribution import pack_blocks, pack_summary

INK = (0.2, 0.2, 0.2)  # the divider caption


def generate_horizontal_labels(pack_groups, output_pdf_path, attribute_order=None, addresses=None, report=None,
                               serials=None, add_dividers=True, notes=None, group_by_code=True, on_progress=None):
    """pack_groups: every selected tab's groups, combined in tab order, each tab already in packing-sheet order
    (horizontal_distribution.read_horizontal). notes: {(tab_no, pack_no): [issue dicts]} shown on the audit page.
    Returns generate_packing_labels()'s page info, with the page numbers of the final PDF.
    on_progress(done, total, step): as generate_packing_labels; drawing is most of the work, the reordering the rest."""
    notes = notes or {}
    if not group_by_code:
        # Not grouped by allocation: each pack's labels in the sheet's store order, no dividers. The code boxes stay:
        # the 'Code for Pack …' columns are columns of the pack like its items, and none is left off a label. The
        # labels only: Item References, the CSVs and the label map are the same either way (the keys don't change)
        pack_groups = {k: pack_groups[k] for k in sorted(pack_groups, key=lambda k: (
            pack_groups[k].get('tab_no', 1), pack_groups[k]['pack_no'], pack_groups[k]['store_no']))}
        add_dividers = False
    tmp = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
    tmp.close()
    try:
        drawn = (lambda d, t, step: on_progress(0.85 * d / max(t, 1), 1, step)) if on_progress else None
        info = generate_packing_labels(pack_groups, tmp.name, attribute_order, addresses, None, serials, None,
                                       on_progress=drawn)
        labels = fitz.open(tmp.name)
        out = fitz.open()
        report_pages = draw_report_pages(out, report, at=0) if report else 0
        # The final order, page by page: an int is a page already in `out` (report, audit, divider), ('label', p) is
        # page p of the labels. The labels go in with one copy at the end (each artwork once) and select() then puts
        # every page in place: copying them a page at a time copied every image again for each page (a 200 MB PDF)
        order = list(range(len(out)))
        for (tab_no, pack_no), keys in pack_blocks(pack_groups):
            before = len(out)
            draw_pack_audit(out, pack_groups, keys, notes.get((tab_no, pack_no), []), add_dividers, group_by_code)
            order += range(before, len(out))
            first = pack_groups[keys[0]]
            codes = []
            for k in keys:
                if not codes or codes[-1][0] != pack_groups[k]['pack_code']:
                    codes.append((pack_groups[k]['pack_code'], []))
                codes[-1][1].append(k)
            for code, code_keys in codes:
                if add_dividers:
                    # No installer barcodes here (the Label Shuffler's dividers have them): each item on these labels
                    # carries its own job-number barcode
                    n = len(code_keys)
                    div = build_divider_sheet(A4_W, A4_H, code or '?', n, n, False)
                    _divider_caption(div[0], first, n)
                    before = len(out)
                    out.insert_pdf(div)
                    order += range(before, len(out))
                    div.close()
                for k in code_keys:
                    order += [('label', p) for p in info['pages'][k]]
        base = len(out)
        out.insert_pdf(labels)
        labels.close()
        new_page = {x[1]: i for i, x in enumerate(order) if isinstance(x, tuple)}
        out.select([base + x[1] if isinstance(x, tuple) else x for x in order])
        if on_progress:
            on_progress(0.95, 1, f"Saving the PDF in packing-sheet order ({len(order)} pages)")
        out.save(output_pdf_path, garbage=1, deflate=True)  # garbage=1 drops what select() left out; images are shared already
        out.close()
    finally:
        os.unlink(tmp.name)
    info['pages'] = {k: [new_page[p] for p in v] for k, v in info['pages'].items()}
    info['copies'] = {k: [[new_page[p] for p in box] for box in v] for k, v in info['copies'].items()}
    info['report_pages'] = report_pages
    return info


def _divider_caption(page, first, count):
    """The pack this divider belongs to, small at the foot, so a stack of dividers can be told apart."""
    pack = f"Pack {first['pack_no']}" + ("" if str(first['pack_name']).strip() == str(first['pack_no']) else f" - {first['pack_name']}")
    text = _latin(f"Tab {first['tab_no']} - {first['tab_name']}  |  {pack}  |  "
                  f"{count} label{'s' if count != 1 else ''}")
    w = fitz.get_text_length(text, fontname='hebo', fontsize=12)
    page.insert_text(((page.rect.width - w) / 2, page.rect.height - 52), text, fontname='hebo', fontsize=12, color=INK)


def draw_pack_audit(doc, pack_groups, keys, issues, dividers, grouped=True):
    """One pack's audit page, before its labels: the counts to check against the packing sheet, where the codes
    came from, and what needs a look. Green header when nothing does, amber when something should be checked."""
    first = pack_groups[keys[0]]
    _, rows = pack_summary(pack_groups, keys)
    boxes = sum(r['boxes'] for _, r in rows)
    columns = first.get('pack_columns') or len({i.get('job_no') for k in keys for i in pack_groups[k]['items']})
    # 'Tab 1 - Pack 1', and a pack with a name of its own says it too: 'Tab 1 - Pack 2: Window kit'
    name = str(first['pack_name']).strip()
    title = f"Tab {first['tab_no']} - Pack {first['pack_no']}" + ('' if name.isdigit() or name == str(first['pack_no']) else f": {name}")
    order = ("Order: stores in sheet order, not grouped by code. Each label still has its code boxes; no dividers." if not grouped else
             "Order: code by code (A, B, C ...), a divider before each code, stores in sheet order." if dividers else
             "Order: code by code (A, B, C ...), stores in sheet order. Dividers off.")
    specs = Counter(pack_groups[k].get('pack_spec_name', '') for k in keys)
    # Items: the pack's item columns in the sheet (every column under the pack's header), not the sum of quantities
    draw_audit_page(doc, title, [("Labels", len(keys)), ("Boxes", boxes), ("Codes", len(rows)), ("Items", columns)],
                    [f"Codes: {first.get('code_source', '')}", order], issues,
                    "Every unit in this pack is on a label, and every store has its packing-sheet code.", specs=specs)
