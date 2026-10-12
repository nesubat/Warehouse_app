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

import pymupdf as fitz

from engine.packing_labels.packing_label_generator import generate_packing_labels, draw_report_pages, A4_W, A4_H, MM2PT, _latin
from engine.label_shuffler.pdf_engine import build_divider_sheet
from engine.packing_labels.horizontal_distribution import pack_blocks, pack_summary

GREEN, RED, AMBER = (0.18, 0.63, 0.36), (0.75, 0.22, 0.17), (0.85, 0.55, 0.0)
INK, MUTED, RULE = (0.2, 0.2, 0.2), (0.45, 0.45, 0.45), (0.82, 0.82, 0.82)
M = 12 * MM2PT


def generate_horizontal_labels(pack_groups, output_pdf_path, attribute_order=None, addresses=None, report=None,
                               serials=None, add_dividers=True, divider_barcodes=None, notes=None):
    """pack_groups: every selected tab's groups, combined in tab order, each tab already in packing-sheet order
    (horizontal_distribution.read_horizontal). divider_barcodes: read_divider_barcodes() of the Signature Links file,
    {(tab, 'Code for <pack>'): {code: [entries]}}. notes: {(tab_no, pack_no): [issue dicts]} shown on the audit page.
    Returns generate_packing_labels()'s page info, with the page numbers of the final PDF."""
    divider_barcodes, notes = divider_barcodes or {}, notes or {}
    tmp = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
    tmp.close()
    try:
        info = generate_packing_labels(pack_groups, tmp.name, attribute_order, addresses, None, serials, None)
        labels = fitz.open(tmp.name)
        out = fitz.open()
        report_pages = draw_report_pages(out, report, at=0) if report else 0
        new_page = {}
        for (tab_no, pack_no), keys in pack_blocks(pack_groups):
            draw_pack_audit(out, pack_groups, keys, notes.get((tab_no, pack_no), []), add_dividers)
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
                    out.insert_pdf(div)
                    div.close()
                for k in code_keys:
                    for p in info['pages'][k]:
                        new_page[p] = len(out)
                        out.insert_pdf(labels, from_page=p, to_page=p)
        labels.close()
        out.save(output_pdf_path, garbage=3, deflate=True)
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


def _text(page, x, y, text, size=10, bold=False, color=INK, right=None):
    font = 'hebo' if bold else 'helv'
    text = _latin(text)
    if right is not None:
        x = right - fitz.get_text_length(text, fontname=font, fontsize=size)
    page.insert_text((x, y), text, fontname=font, fontsize=size, color=color)


def _fit(text, size, width, bold=False):
    """The text cut to the width, ending in '...' when it had to be cut."""
    font = 'hebo' if bold else 'helv'
    text = _latin(text)
    if fitz.get_text_length(text, fontname=font, fontsize=size) <= width:
        return text
    while text and fitz.get_text_length(text + '...', fontname=font, fontsize=size) > width:
        text = text[:-1]
    return text + '...' if text else ''


def draw_pack_audit(doc, pack_groups, keys, issues, dividers):
    """One pack's audit page, before its labels: the counts to check against the packing sheet, where the codes
    came from, and what needs a look. Green header when nothing does, amber when something should be checked."""
    first = pack_groups[keys[0]]
    _, rows = pack_summary(pack_groups, keys)
    boxes = sum(r['boxes'] for _, r in rows)
    units = sum(float(i['qty']) for k in keys for i in pack_groups[k]['items'])
    title = f"Pack {first['pack_name']}" if str(first['pack_name']).strip().isdigit() else first['pack_name']
    colour = AMBER if issues else GREEN

    page = doc.new_page(width=A4_W, height=A4_H)
    page.draw_rect(fitz.Rect(0, 0, A4_W, 54), color=colour, fill=colour)
    _text(page, M, 24, f"PACK AUDIT - TAB {first['tab_no']} - PACK {first['pack_no']}", 9, True, (1, 1, 1))
    _text(page, M, 44, title, 18, True, (1, 1, 1))
    _text(page, 0, 44, "Check before packing" if issues else "All checks passed", 12, True, (1, 1, 1), right=A4_W - M)

    # The figures to check against the packing sheet
    x = M
    for label, value in (("Labels", f"{len(keys)}"), ("Boxes", f"{boxes}"), ("Codes", f"{len(rows)}"), ("Units", f"{units:g}")):
        w = max(70, fitz.get_text_length(value, fontname='hebo', fontsize=15) + 22)
        page.draw_rect(fitz.Rect(x, 68, x + w - 8, 112), color=RULE, fill=(0.97, 0.97, 0.97), radius=0.15, width=0.6)
        _text(page, x + 8, 84, label.upper(), 7.5, True, MUTED)
        _text(page, x + 8, 104, value, 15, True)
        x += w
    _text(page, M, 134, f"Codes: {first.get('code_source', '')}", 9, False, MUTED)
    _text(page, M, 148, ("Order: code by code (A, B, C ...), a divider before each code, stores in sheet order."
                         if dividers else "Order: code by code (A, B, C ...), stores in sheet order. Dividers off."),
          9, False, MUTED)

    # What needs a look, for this pack only
    y = 176
    if issues:
        _text(page, M, y, "Needs a look", 11, True, AMBER)
        for i in issues[:20]:
            y += 15
            where = " ".join(p for p in (f"Row {i.get('rows')}" if i.get('rows') else '', i.get('col', '')) if p)
            _text(page, M, y, _fit(f"- {where}: {i.get('text', '')}" if where else f"- {i.get('text', '')}", 9.5, A4_W - 2 * M), 9.5)
    else:
        _text(page, M, y, "Every unit in this pack is on a label, and every store has its packing-sheet code.", 9.5, False, GREEN)
