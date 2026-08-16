import datetime
import pymupdf as fitz  # `fitz` is PyMuPDF's legacy, deprecated import name - `pymupdf` is current
import os
import math
import re
import collections

# ==========================================
# UTILITY: SIZE-SAFE TEXTBOX INSERTION
# ==========================================
def _fit_textbox(page, rect, text, fontname, max_fontsize, align, color, min_fontsize=6):
    """Inserts text into rect, shrinking the font until it fits. PyMuPDF draws nothing at
    all when a textbox overflows, so on tight/short pages (e.g. split-layout half-page
    dividers) a fixed fontsize can silently render a blank box - this guarantees something
    legible is always shown."""
    fontsize = max_fontsize
    while fontsize > min_fontsize:
        if page.insert_textbox(rect, text, fontname=fontname, fontsize=fontsize, align=align, color=color, render_mode=3) >= 0:
            break
        fontsize -= 1
    page.insert_textbox(rect, text, fontname=fontname, fontsize=fontsize, align=align, color=color)
    return fontsize

# ==========================================
# UTILITY: MISSING STORE GROUPING
# ==========================================
def group_missing_stores_by_code(store_mapping, found_stores):
    """Groups the stores that never got matched by their signature code, keeping each
    store in its original spreadsheet row order within its code group. Code groups
    themselves are ordered with the same (length, alphabetical) key used everywhere
    else in this file to decide the divider/label stitching order, so the audit
    report's list always lines up with the order things actually appear later in
    the PDF."""
    missing_by_code = {}
    for store in store_mapping:  # dict iteration order == original spreadsheet row order
        if store not in found_stores:
            missing_by_code.setdefault(store_mapping[store], []).append(store)

    ordered_codes = sorted(missing_by_code.keys(), key=lambda c: (len(c), c))
    return [(code, missing_by_code[code]) for code in ordered_codes]


# ==========================================
# UTILITY: PROFESSIONAL AUDIT REPORT GENERATOR
# ==========================================
def build_audit_report(pdf_filename, page_width, report_height, missing_stores_grouped, unmatched_count, blank_count, item_name="PAGES"):
    """Generates a highly professional, formatted warning page."""
    audit_doc = fitz.open()
    audit_page = audit_doc.new_page(width=page_width, height=report_height)

    # 1. Determine Status & Colors (RGB normalized 0 to 1)
    has_issues = bool(missing_stores_grouped or unmatched_count or blank_count)
    bg_color = (0.75, 0.22, 0.17) if has_issues else (0.18, 0.63, 0.36)  # Deep Red vs Professional Green
    text_dark = (0.2, 0.2, 0.2)
    line_color = (0.8, 0.8, 0.8)
    
    # 2. Draw Solid Header Bar
    header_rect = fitz.Rect(0, 0, page_width, 50)
    audit_page.draw_rect(header_rect, color=bg_color, fill=bg_color)
    
    # 3. Header Text
    title_text = "REPORT: ATTENTION REQUIRED" if has_issues else "REPORT: SUCCESS"
    audit_page.insert_text((30, 32), title_text, fontname="hebo", fontsize=16, color=(1, 1, 1))
    
    # 4. Report Metadata
    y_pos = 80
    audit_page.insert_text((30, y_pos), f"File Processed: {pdf_filename}", fontname="hebo", fontsize=12, color=text_dark)
    
    y_pos += 15
    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    audit_page.insert_text((30, y_pos), f"Generated On: {current_time}", fontname="helv", fontsize=10, color=(0.4, 0.4, 0.4))
    
    SECTION_GAP = 25  # consistent gap used after EVERY separator line in this report

    y_pos += 20
    audit_page.draw_line(fitz.Point(30, y_pos), fitz.Point(page_width - 30, y_pos), color=line_color, width=1.5)
    y_pos += SECTION_GAP

    # 5. Success Scenario
    if not has_issues:
        audit_page.insert_text((30, y_pos), "No issues detected. All labels matched the distribution list flawlessly.", fontname="helv", fontsize=12, color=text_dark)
        return audit_doc

    # --- HELPER: Smart Page Breaks (used by every section EXCEPT Missing Stores,
    # which flows into its own multi-column layout further below) ---
    def check_page_break(current_y, required_space):
        nonlocal audit_page, audit_doc
        if current_y + required_space > report_height - 40:
            audit_page = audit_doc.new_page(width=page_width, height=report_height)
            # Add a slim continuation header
            audit_page.draw_rect(fitz.Rect(0, 0, page_width, 15), color=bg_color, fill=bg_color)
            return 40
        return current_y

    # 7. Unmatched & Blank Summaries
    if unmatched_count or blank_count:
        y_pos = check_page_break(y_pos, 50)

        if unmatched_count:
            y_pos = check_page_break(y_pos, 20)
            audit_page.insert_text((30, y_pos), f"Unmatched {item_name.title()}:", fontname="hebo", fontsize=11, color=text_dark)
            audit_page.insert_text((170, y_pos), str(unmatched_count), fontname="hebo", fontsize=11, color=bg_color)
            y_pos += 20

        if blank_count:
            y_pos = check_page_break(y_pos, 20)
            audit_page.insert_text((30, y_pos), f"Blank {item_name.title()}:", fontname="hebo", fontsize=11, color=text_dark)
            audit_page.insert_text((170, y_pos), str(blank_count), fontname="hebo", fontsize=11, color=(0.6, 0.6, 0.6))
            y_pos += 20

        y_pos += SECTION_GAP - 20  # top up to the same SECTION_GAP used after every other separator

    # 8. Missing Stores Section - grouped by code, laid out in auto-width columns.
    # Rendered LAST (after every other message above) so it's free to spill into as
    # many side-by-side columns and follow-on pages as it needs.
    if missing_stores_grouped:
        y_pos = check_page_break(y_pos, 40)
        audit_page.draw_line(fitz.Point(30, y_pos), fitz.Point(page_width - 30, y_pos), color=line_color, width=1.5)
        y_pos += SECTION_GAP
        y_pos = check_page_break(y_pos, 40)
        audit_page.insert_text((30, y_pos), "MISSING STORES:", fontname="hebo", fontsize=12, color=bg_color)
        y_pos += 30

        LEFT_MARGIN = 30
        RIGHT_MARGIN = 30
        COLUMN_GAP = 24
        HEADER_ROW_H = 22
        STORE_ROW_H = 18
        MIN_COL_WIDTH = 90
        bottom_limit = report_height - 40

        # Flatten into a single ordered stream of drawable rows, so header rows and
        # store rows can be measured/flowed identically by the column-fill loop below.
        rows = []
        for code, store_names in missing_stores_grouped:
            rows.append({"text": f"CODE {code} - {len(store_names)} missing", "header": True, "code": code})
            for name in store_names:
                rows.append({"text": name, "header": False, "code": code})

        def row_draw_width(row):
            fontname = "hebo" if row["header"] else "helv"
            fontsize = 11 if row["header"] else 10
            w = fitz.get_text_length(row["text"], fontname=fontname, fontsize=fontsize)
            return w if row["header"] else w + 16  # +16 accounts for the bullet + indent

        page_top_y = y_pos  # top-of-column y for whatever page the section is currently on
        col_x = LEFT_MARGIN
        idx = 0
        n = len(rows)

        while idx < n:
            # rows[idx] is always a code header here (every group starts with one,
            # and a mid-group break re-inserts a "(continued)" header). If the
            # current page doesn't even have room for that header PLUS one store
            # row, don't start a column here at all - it would draw a store-less
            # header and immediately need another "(continued)" header right after
            # it, over and over, until the page runs out of width. Jump straight to
            # a fresh page instead (unless we're already on one, to guarantee this
            # can't loop forever on a pathologically short report_height).
            if page_top_y != 40 and page_top_y + HEADER_ROW_H + STORE_ROW_H > bottom_limit:
                audit_page = audit_doc.new_page(width=page_width, height=report_height)
                audit_page.draw_rect(fitz.Rect(0, 0, page_width, 15), color=bg_color, fill=bg_color)
                page_top_y = 40
                col_x = LEFT_MARGIN
                continue

            # --- Fill one column vertically, starting at (col_x, page_top_y) ---
            col_start_idx = idx
            cursor_y = page_top_y
            col_rows = []
            while idx < n:
                row_h = HEADER_ROW_H if rows[idx]["header"] else STORE_ROW_H
                if cursor_y + row_h > bottom_limit:
                    break
                col_rows.append(rows[idx])
                cursor_y += row_h
                idx += 1

            # Safety net: if not even one row fits (a pathologically short page), force
            # one through anyway so this can never loop forever.
            if not col_rows and idx < n:
                col_rows.append(rows[idx])
                idx += 1

            # --- Size this column from ONLY its own content ---
            col_width = max(MIN_COL_WIDTH, max(row_draw_width(r) for r in col_rows) + 20)

            # --- Does this column fit on the current page? ---
            if col_x > LEFT_MARGIN and col_x + col_width > page_width - RIGHT_MARGIN:
                audit_page = audit_doc.new_page(width=page_width, height=report_height)
                audit_page.draw_rect(fitz.Rect(0, 0, page_width, 15), color=bg_color, fill=bg_color)
                page_top_y = 40
                col_x = LEFT_MARGIN
                idx = col_start_idx  # redo this column's row selection at the new page's top
                continue

            # --- Separator to the left of every column after the first on a page ---
            if col_x > LEFT_MARGIN:
                sep_x = col_x - COLUMN_GAP / 2
                audit_page.draw_line(fitz.Point(sep_x, page_top_y - 10), fitz.Point(sep_x, cursor_y),
                                      color=line_color, width=1)

            # --- Draw the column ---
            draw_y = page_top_y
            for row in col_rows:
                if row["header"]:
                    audit_page.insert_text((col_x, draw_y), row["text"], fontname="hebo", fontsize=11, color=bg_color)
                    draw_y += HEADER_ROW_H
                else:
                    audit_page.draw_circle(fitz.Point(col_x + 6, draw_y - 4), 2, color=text_dark, fill=text_dark)
                    audit_page.insert_text((col_x + 16, draw_y), row["text"], fontname="helv", fontsize=10, color=text_dark)
                    draw_y += STORE_ROW_H

            # --- If a code group got cut off mid-list, resume it with a "(continued)"
            # header at the top of the very next column ---
            if idx < n and not rows[idx]["header"]:
                cont_code = rows[idx]["code"]
                rows.insert(idx, {"text": f"CODE {cont_code} (continued)", "header": True, "code": cont_code})
                n += 1

            col_x += col_width + COLUMN_GAP

    return audit_doc

# ==========================================
# UTILITY: STORE-NAME COLLISION DETECTION
# ==========================================
def find_name_collisions(store_names):
    """Finds store name pairs where one name is fully contained inside another
    (case-insensitive), e.g. 'Northlands' vs 'Northlands NZ'. Such pairs are a risk
    for the substring-based label matcher silently grouping both labels under one name."""
    names = list(store_names)
    collisions = []
    for i, name_a in enumerate(names):
        low_a = name_a.lower()
        for j, name_b in enumerate(names):
            if i == j:
                continue
            low_b = name_b.lower()
            if low_a != low_b and low_a in low_b:
                collisions.append((name_a, name_b))  # name_a is a substring of name_b
    return collisions


def analyze_matches(store_mapping, sorted_stores, found_stores):
    """Cross-checks matched vs. expected stores per signature code, and flags
    ambiguous name collisions where one of a colliding pair was matched while the
    other was not - the signature of a label being grouped under the wrong name.

    The collision signal only ever surfaces as a plain red-border flip + generic
    "Possible mismatch" disclaimer on that code's own divider sheet (never a named
    explanation) - the missing store already shows up on the Missing Stores list
    and the mis-attributed labels (if any) already show up in Unmatched, so a
    separate audit-report callout would just repeat information that's already
    there in more detail."""
    expected_by_code = {}
    for store in sorted_stores:
        expected_by_code.setdefault(store_mapping[store], []).append(store)

    mismatched_codes = set()
    for name_a, name_b in find_name_collisions(store_mapping.keys()):
        a_found = name_a in found_stores
        b_found = name_b in found_stores
        if a_found == b_found:
            continue  # both found or both missing - not the ambiguous signal we care about

        found_name = name_a if a_found else name_b
        missing_name = name_b if a_found else name_a
        mismatched_codes.add(store_mapping[found_name])
        mismatched_codes.add(store_mapping[missing_name])

    divider_info_by_code = {}
    for code, stores_for_code in expected_by_code.items():
        found_for_code = [s for s in stores_for_code if s in found_stores]
        divider_info_by_code[code] = {
            'matched_count': len(found_for_code),
            'expected_count': len(stores_for_code),
            'collision_note': code in mismatched_codes,
        }

    return divider_info_by_code


# ==========================================
# UTILITY: PROFESSIONAL DIVIDER SHEET
# ==========================================
def build_divider_sheet(page_width, page_height, signature_code, matched_count, expected_count, collision_note=False):
    """Generates a professional divider page: green border if the number of distinct
    stores matched into this code equals the expected count from the signature links,
    red border otherwise (with a shortfall/excess disclaimer). The full missing-store
    list lives only on the audit report page, not here - so it can't silently run out
    of room and get dropped the way a fixed-height divider sheet would if a code had
    a lot of missing stores.

    "Matched Labels" shows just matched_count when it equals expected_count (the
    common, all-clear case), or "matched/expected" the moment it doesn't - one
    line either way, instead of separate always-shown counters.

    collision_note is just a bool here - it only flips the border red and adds the
    generic "Possible mismatch" disclaimer below; the specific reasoning behind it
    is for the tool operator to read on the audit report, never on a divider a
    warehouse floor worker might see.
    Layout scales down on short pages (e.g. split-layout half-page dividers) so nothing
    overflows or gets skipped past the bottom edge."""
    is_match = (matched_count == expected_count) and not collision_note
    bottom_limit = page_height - 20

    doc = fitz.open()
    page = doc.new_page(width=page_width, height=page_height)

    # 1. Draw Border - Green when verified, Red when it needs attention
    border_color = (0.18, 0.63, 0.36) if is_match else (0.75, 0.22, 0.17)
    page.draw_rect(fitz.Rect(15, 15, page_width - 15, page_height - 15), color=border_color, width=40)

    # 2. Header Text - shrinks to fit short/narrow pages so it never renders blank
    header_height = min(200, page_height * 0.6)
    header_top = max(20, page_height / 2 - header_height / 2)
    header_bottom = min(bottom_limit, header_top + header_height)
    header_rect = fitz.Rect(40, header_top, page_width - 40, header_bottom)
    label_count = str(matched_count) if matched_count == expected_count else f"{matched_count}/{expected_count}"
    text = f"CODE {signature_code}\n\nMatched Labels: {label_count}"
    _fit_textbox(page, header_rect, text, "hebo", 40, fitz.TEXT_ALIGN_CENTER, (0, 0, 0))

    # 3. Disclaimer + collision note when it doesn't check out - each block only
    # draws if there is still room left, and is clipped to the page so it never crashes.
    if not is_match:
        y = header_bottom + 10

        diff = expected_count - matched_count
        if diff > 0:
            disclaimer = f"Short by {diff} store(s) (expected {expected_count})"
        elif diff < 0:
            disclaimer = f"{-diff} extra store(s) matched (expected {expected_count})"
        else:
            disclaimer = "Possible mismatch - please verify"

        if y + 15 < bottom_limit:
            disclaimer_rect = fitz.Rect(40, y, page_width - 40, min(y + 25, bottom_limit))
            _fit_textbox(page, disclaimer_rect, disclaimer, "hebo", 14, fitz.TEXT_ALIGN_CENTER, border_color)
            y += 30

    return doc


def build_unmatched_divider_sheet(page_width, page_height, count):
    """Generates a divider page (red border) marking the start of the Unmatched Pages group."""
    doc = fitz.open()
    page = doc.new_page(width=page_width, height=page_height)

    border_color = (0.75, 0.22, 0.17)  # Deep Red - unmatched pages always need attention
    page.draw_rect(fitz.Rect(15, 15, page_width - 15, page_height - 15), color=border_color, width=40)

    box_height = min(200, page_height * 0.6)
    top = max(20, page_height / 2 - box_height / 2)
    text_rect = fitz.Rect(40, top, page_width - 40, min(page_height - 20, top + box_height))
    text = f"UNMATCHED PAGES\n\nTotal Pages: {count}"
    _fit_textbox(page, text_rect, text, "hebo", 40, fitz.TEXT_ALIGN_CENTER, (0, 0, 0))

    return doc

# ==========================================
# ENGINE: SPLIT LAYOUT (CUT & STACK)
# ==========================================
def process_split_layout(doc, sorted_stores, store_mapping, page_width, page_height, pdf_filename, add_dividers):
    all_extracted_halves = []
    found_stores = set()
    unmatched_halves = []
    blank_halves = []

    half_height = page_height / 2
    top_rect = fitz.Rect(0, 0, page_width, half_height)
    bottom_rect = fitz.Rect(0, half_height, page_width, page_height)

    top_target = fitz.Rect(0, 0, 289, half_height)
    bottom_target = fitz.Rect(0, half_height, 289, page_height)

    for page_num in range(len(doc)):
        page = doc[page_num]

        top_text = page.get_text("text", clip=top_target).lower()
        bottom_text = page.get_text("text", clip=bottom_target).lower()

        top_match = next((store for store in sorted_stores if store.lower() in top_text), None)
        bottom_match = next((store for store in sorted_stores if store.lower() in bottom_text), None)

        if top_match:
            found_stores.add(top_match)
            all_extracted_halves.append({'doc_type': 'original', 'page_num': page_num, 'half': 'top', 'code': store_mapping[top_match]})
        elif not top_text.strip():
            blank_halves.append({'doc_type': 'original', 'page_num': page_num, 'half': 'top'})
            print(f"[DEBUG] BLANK Top Half detected on Page {page_num + 1}")
        else:
            unmatched_halves.append({'doc_type': 'original', 'page_num': page_num, 'half': 'top'})
            print(f"[DEBUG] Unmatched Top Half on Page {page_num + 1}")

        if bottom_match:
            found_stores.add(bottom_match)
            all_extracted_halves.append({'doc_type': 'original', 'page_num': page_num, 'half': 'bottom', 'code': store_mapping[bottom_match]})
        elif not bottom_text.strip():
            blank_halves.append({'doc_type': 'original', 'page_num': page_num, 'half': 'bottom'})
            print(f"[DEBUG] BLANK Bottom Half detected on Page {page_num + 1}")
        else:
            unmatched_halves.append({'doc_type': 'original', 'page_num': page_num, 'half': 'bottom'})
            print(f"[DEBUG] Unmatched Bottom Half on Page {page_num + 1}")

    # Generate Universal Audit Report
    missing_stores_grouped = group_missing_stores_by_code(store_mapping, found_stores)
    divider_info_by_code = analyze_matches(store_mapping, sorted_stores, found_stores)
    audit_doc = build_audit_report(pdf_filename, page_width, half_height, missing_stores_grouped, len(unmatched_halves), len(blank_halves), "HALVES")

    audit_halves = [{'doc_type': 'audit', 'page_num': i, 'half': 'full'} for i in range(len(audit_doc))]

    # Stitching
    final_shuffled_doc = fitz.open()
    # NEW: Sort by length first, then alphabetically (A, B, C... Z, AA, AB)
    ordered_matched = sorted(all_extracted_halves,key=lambda x: (len(x['code']), x['code']))
    # NEW: Generate and inject Divider halves
    if add_dividers:
        grouped_halves = []
        current_code = None
        divider_docs = [] # Keep open until save
        for item in ordered_matched:
            if item['code'] != current_code:
                current_code = item['code']
                info = divider_info_by_code.get(current_code, {'matched_count': 0, 'expected_count': 0, 'collision_note': None})
                div_doc = build_divider_sheet(page_width, half_height, current_code, info['matched_count'], info['expected_count'], info['collision_note'])
                divider_docs.append(div_doc)
                grouped_halves.append({'doc_type': 'divider', 'doc_ref': div_doc, 'half': 'full'})

            grouped_halves.append(item)

        # NEW: Divider half marking the start of the Unmatched Pages group
        if unmatched_halves:
            unmatched_div_doc = build_unmatched_divider_sheet(page_width, half_height, len(unmatched_halves))
            divider_docs.append(unmatched_div_doc)
            unmatched_stack = [{'doc_type': 'divider', 'doc_ref': unmatched_div_doc, 'half': 'full'}] + unmatched_halves
        else:
            unmatched_stack = unmatched_halves

        master_stack = audit_halves + grouped_halves + unmatched_stack + blank_halves
    else:
        master_stack = audit_halves + ordered_matched + unmatched_halves + blank_halves
    
    total_halves = len(master_stack)
    halfway_point = math.ceil(total_halves / 2)
    top_stack = master_stack[:halfway_point]
    bottom_stack = master_stack[halfway_point:]
    
    audit_rect = fitz.Rect(0, 0, page_width, half_height)
    
    for i in range(halfway_point):
        new_page = final_shuffled_doc.new_page(width=page_width, height=page_height)
        
        # --- PASTE TOP HALF ---
        top_item = top_stack[i]
        if top_item['doc_type'] == 'audit':
            new_page.show_pdf_page(top_rect, audit_doc, top_item['page_num'], clip=audit_rect)
        elif top_item['doc_type'] == 'divider':
            # Handle the divider specifically by referencing its doc_ref!
            new_page.show_pdf_page(top_rect, top_item['doc_ref'], 0, clip=audit_rect)
        else:
            src_page = doc[top_item['page_num']]
            src_clip = top_rect if top_item['half'] == 'top' else bottom_rect
            new_page.show_pdf_page(top_rect, doc, src_page.number, clip=src_clip)
        
        # --- PASTE BOTTOM HALF ---
        if i < len(bottom_stack):
            bot_item = bottom_stack[i]
            if bot_item['doc_type'] == 'audit':
                new_page.show_pdf_page(bottom_rect, audit_doc, bot_item['page_num'], clip=audit_rect)
            elif bot_item['doc_type'] == 'divider':
                # Handle the divider specifically by referencing its doc_ref!
                new_page.show_pdf_page(bottom_rect, bot_item['doc_ref'], 0, clip=audit_rect)
            else:
                src_page_b = doc[bot_item['page_num']]
                src_clip_b = top_rect if bot_item['half'] == 'top' else bottom_rect
                new_page.show_pdf_page(bottom_rect, doc, src_page_b.number, clip=src_clip_b)
    audit_doc.close()
    # ADD THIS CLEANUP LOOP
    if add_dividers:
        for d in divider_docs:
            d.close()
    return final_shuffled_doc

# ==========================================
# ENGINE: STANDARD LAYOUT
# ==========================================
def process_standard_layout(doc, sorted_stores, store_mapping, page_width, page_height, pdf_filename, add_dividers):
    found_stores = set()
    current_store = None
    unmatched_pages = []
    blank_pages = []
    target = fitz.Rect(0, 0, 317.5, page_height)
    expected_extra_pages = 0

    # Every store should have exactly one label instance. A store's label can
    # legitimately span several physical pages ("Page 1/3, 2/3, 3/3"), and its
    # header may repeat on each of those - that's still ONE instance, tracked
    # via expected_extra_pages below. A match on the same store OUTSIDE that
    # continuation window is a second, separate instance - a sign something's
    # wrong - so pages are staged here first and only committed to a code
    # bucket once we know, at the end, how many real instances each store had.
    store_pages = collections.defaultdict(list)
    store_instance_count = collections.defaultdict(int)

    for page_num in range(len(doc)):
        page = doc[page_num]
        full_text = page.get_text("text").strip()

        if not full_text:
            print(f"[DEBUG] Blank Page detected on Page {page_num + 1}. Skipping.")
            blank_pages.append(page)
            continue

        text = page.get_text("text", clip=target).lower()
        match = next((store for store in sorted_stores if store.lower() in text), None)

        if match:
            is_continuation = (match == current_store and expected_extra_pages > 0)

            current_store = match
            found_stores.add(match)

            page_info = re.search(r"page\s+(\d+)\s*/\s*(\d+)", text)
            if page_info:
                expected_extra_pages = int(page_info.group(2)) - int(page_info.group(1))
            else:
                expected_extra_pages = 0

            if not is_continuation:
                store_instance_count[match] += 1

        elif expected_extra_pages > 0 and current_store:
            expected_extra_pages -= 1

        else:
            current_store = None
            expected_extra_pages = 0

        if not current_store:
            unmatched_pages.append(page)
            print(f"[DEBUG] Unmatched/Junk Label on Page: {page_num + 1}")
        else:
            store_pages[current_store].append(page_num)

    # Where a store shows up as 2+ separate instances, none of its pages can be
    # trusted to be the single "true" copy - move all of them to Unmatched for the
    # tool operator to sort out by hand. No audit-report callout is needed for this:
    # the store already shows up on the Missing Stores list below (since it has no
    # confirmed page in any bucket) and its actual pages are sitting right there in
    # Unmatched for review - a separate message would just repeat that.
    code_buckets = {}
    code_final_store_counts = collections.Counter()  # distinct stores actually left in each code's bucket
    confirmed_stores = set()  # stores with exactly one clean instance - a real page in a code bucket
    for store, page_nums in store_pages.items():
        if store_instance_count[store] > 1:
            for p_num in page_nums:
                unmatched_pages.append(doc[p_num])
        else:
            confirmed_stores.add(store)
            code = store_mapping[store]
            code_buckets.setdefault(code, []).extend(page_nums)
            code_final_store_counts[code] += 1
    unmatched_pages.sort(key=lambda p: p.number)  # keep original document order

    # Generate Universal Audit Report. Missing-store/collision checks run against
    # confirmed_stores (NOT the raw found_stores) so a store whose pages all got
    # pulled into Unmatched for being a duplicate correctly shows up as missing too
    # - it has no confirmed label in any bucket, exactly as if it were never found.
    missing_stores_grouped = group_missing_stores_by_code(store_mapping, confirmed_stores)
    divider_info_by_code = analyze_matches(store_mapping, sorted_stores, confirmed_stores)
    audit_doc = build_audit_report(pdf_filename, page_width, page_height, missing_stores_grouped, len(unmatched_pages), len(blank_pages), "PAGES")

    # Stitching
    final_shuffled_doc = fitz.open()

    if len(audit_doc) > 0:
        final_shuffled_doc.insert_pdf(audit_doc)

    # NEW: Sort by length first, then alphabetically (A, B, C... Z, AA, AB)
    sorted_codes = sorted(code_buckets.keys(), key=lambda x: (len(x), x))  # Sort by length then alphabetically
    for code in sorted_codes:
        if add_dividers:
            info = divider_info_by_code.get(code, {'matched_count': 0, 'expected_count': 0, 'collision_note': None})
            # Use the accurate post-dedup store count, not analyze_matches' found_stores-based
            # matched_count - that one still counts a store as "matched" even if every one of
            # its pages ended up diverted to Unmatched for being a duplicate.
            div_doc = build_divider_sheet(page_width, page_height, code, code_final_store_counts[code], info['expected_count'], info['collision_note'])
            final_shuffled_doc.insert_pdf(div_doc)
            div_doc.close()
        for p_num in code_buckets[code]:
            final_shuffled_doc.insert_pdf(doc, from_page=p_num, to_page=p_num)

    if unmatched_pages:
        if add_dividers:
            unmatched_div_doc = build_unmatched_divider_sheet(page_width, page_height, len(unmatched_pages))
            final_shuffled_doc.insert_pdf(unmatched_div_doc)
            unmatched_div_doc.close()
        unmatched_doc = fitz.open()
        for p in unmatched_pages:
            unmatched_doc.insert_pdf(p.parent, from_page=p.number, to_page=p.number)
        final_shuffled_doc.insert_pdf(unmatched_doc)
        unmatched_doc.close()

    if blank_pages:
        blank_doc = fitz.open()
        for p in blank_pages:
            blank_doc.insert_pdf(p.parent, from_page=p.number, to_page=p.number)
        final_shuffled_doc.insert_pdf(blank_doc)
        blank_doc.close()

    audit_doc.close()
    return final_shuffled_doc

# ==========================================
# MASTER ORCHESTRATOR
# ==========================================
def process_and_shuffle_pdf(input_pdf_path, store_mapping, output_pdf_path, signature_header="Unknown Header", add_dividers=False):
    """
    Main entry point. Detects layout and routes traffic to the correct processing engine.
    """
    sorted_stores = sorted(store_mapping.keys(), key=len, reverse=True)
    
    doc = fitz.open(input_pdf_path)
    first_page = doc[0]
    
    page_width = first_page.rect.width 
    page_height = first_page.rect.height
    is_split_layout = page_width < page_height
    pdf_filename = os.path.basename(input_pdf_path).strip()
    width_mm = math.ceil(page_width * 0.3528 * 10) / 10
    height_mm = math.ceil(page_height * 0.3528 * 10) / 10
    
    print(f"\n[ENGINE] Processing {pdf_filename} against Signature: '{signature_header}'")
    print(f"[ENGINE] Detected Page Size: {width_mm} x {height_mm} (Split Layout: {is_split_layout})")

    # Traffic Router
    if is_split_layout:
        final_shuffled_doc = process_split_layout(doc, sorted_stores, store_mapping, page_width, page_height, pdf_filename, add_dividers)
    else:
        final_shuffled_doc = process_standard_layout(doc, sorted_stores, store_mapping, page_width, page_height, pdf_filename, add_dividers)

    # Final Export Cleanup
    final_shuffled_doc.set_page_labels([])
    final_shuffled_doc.save(output_pdf_path)
    final_shuffled_doc.close()
    doc.close()
    
    print(f"\n[SUCCESS] Document saved to: {output_pdf_path}")
    return True