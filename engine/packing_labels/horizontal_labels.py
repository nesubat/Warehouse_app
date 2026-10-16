"""The packing labels PDF for a horizontal distribution file, in packing-sheet order:

  job report (page 1)
  for each tab, for each pack:   audit page   - the pack's codes as on its packing sheet, with stores, boxes and
                                                quantities per item, where the codes came from, and what needs a look
      for each code A, B, C …:   divider      - the Label Shuffler's divider sheet (pdf_engine.build_divider_sheet),
                                                with the installer barcodes when the Signature Links file has them
                                 labels       - the standard A4 packing labels (generate_packing_labels), one per box

The labels are drawn by generate_packing_labels() exactly as for vertical files, then put in this order. The page
numbers returned (and written to the label map) point at the final PDF, so Stitch Labels finds every label.
"""
import os
import tempfile
from collections import Counter

import pymupdf as fitz

from engine.packing_labels.packing_label_generator import generate_packing_labels, draw_report_pages, A4_W, A4_H, _latin, \
    draw_audit_page, spec_summary
from engine.label_shuffler.pdf_engine import build_divider_sheet
from engine.packing_labels.horizontal_distribution import pack_blocks, pack_summary

INK = (0.2, 0.2, 0.2)  # the divider caption


def generate_horizontal_labels(pack_groups, output_pdf_path, attribute_order=None, addresses=None, report=None,
                               serials=None, add_dividers=True, divider_barcodes=None, notes=None, group_by_code=True):
    """pack_groups: every selected tab's groups, combined in tab order, each tab already in packing-sheet order
    (horizontal_distribution.read_horizontal). divider_barcodes: read_divider_barcodes() of the Signature Links file,
    {(tab, 'Code for <pack>'): {code: [entries]}}. notes: {(tab_no, pack_no): [issue dicts]} shown on the audit page.
    Returns generate_packing_labels()'s page info, with the page numbers of the final PDF."""
    divider_barcodes, notes = divider_barcodes or {}, notes or {}
    if not group_by_code:
        # Not grouped by allocation: each pack's labels in the sheet's store order, and no packing-sheet code on them
        # (the Packing Spec panel says only the pack, no code boxes), no dividers. The labels only: Item References,
        # the CSVs and the label map are the same either way (the keys don't change)
        ordered = sorted(pack_groups, key=lambda k: (pack_groups[k].get('tab_no', 1), pack_groups[k]['pack_no'],
                                                     pack_groups[k]['store_no']))
        pack_groups = {k: {**pack_groups[k], 'pack_tag': f"PACK {pack_groups[k]['pack_no']}", 'code_cells': []}
                       for k in ordered}
        add_dividers = False
    tmp = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
    tmp.close()
    try:
        info = generate_packing_labels(pack_groups, tmp.name, attribute_order, addresses, None, serials, None)
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
                    entries = _barcodes_for(divider_barcodes, first['tab_name'], first['pack_name'], code)
                    n = len(code_keys)
                    div = build_divider_sheet(A4_W, A4_H, code or '?', n, n, False, barcode_entries=entries)
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
        out.save(output_pdf_path, garbage=1, deflate=True)  # garbage=1 drops what select() left out; images are shared already
        out.close()
    finally:
        os.unlink(tmp.name)
    info['pages'] = {k: [new_page[p] for p in v] for k, v in info['pages'].items()}
    info['copies'] = {k: [[new_page[p] for p in box] for box in v] for k, v in info['copies'].items()}
    info['report_pages'] = report_pages
    return info


def _barcodes_for(lookup, tab, pack, code):
    """The installer barcodes for one code's divider, keyed as the Label Shuffler keys them."""
    for (t, header), by_code in lookup.items():
        name = header[len('Code for '):] if header.startswith('Code for ') else header
        if t == tab and name.strip().lower() in (str(pack).strip().lower(), f"pack {pack}".strip().lower()):
            return by_code.get(code)
    return None


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
    units = sum(float(i['qty']) for k in keys for i in pack_groups[k]['items'])
    # 'Tab 1 - Pack 1', and a pack with a name of its own says it too: 'Tab 1 - Pack 2: Window kit'
    name = str(first['pack_name']).strip()
    title = f"Tab {first['tab_no']} - Pack {first['pack_no']}" + ('' if name.isdigit() or name == str(first['pack_no']) else f": {name}")
    order = ("Order: stores in sheet order, not grouped by code. No code on the labels, no dividers." if not grouped else
             "Order: code by code (A, B, C ...), a divider before each code, stores in sheet order." if dividers else
             "Order: code by code (A, B, C ...), stores in sheet order. Dividers off.")
    specs = Counter(pack_groups[k].get('pack_spec_name', '') for k in keys)
    draw_audit_page(doc, title, [("Labels", len(keys)), ("Boxes", boxes), ("Codes", len(rows)), ("Units", f"{units:g}")],
                    [spec_summary(specs), f"Codes: {first.get('code_source', '')}", order], issues,
                    "Every unit in this pack is on a label, and every store has its packing-sheet code.")
