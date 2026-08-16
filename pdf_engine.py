import datetime
import pymupdf as fitz  # `fitz` is PyMuPDF's legacy, deprecated import name - `pymupdf` is current
import os
import math
import re

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
# UTILITY: PROFESSIONAL AUDIT REPORT GENERATOR
# ==========================================
def build_audit_report(pdf_filename, page_width, report_height, missing_stores, unmatched_count, blank_count, item_name="PAGES", collision_warnings=None):
    """Generates a highly professional, formatted warning page."""
    collision_warnings = collision_warnings or []
    audit_doc = fitz.open()
    audit_page = audit_doc.new_page(width=page_width, height=report_height)

    # 1. Determine Status & Colors (RGB normalized 0 to 1)
    has_issues = bool(missing_stores or unmatched_count or blank_count or collision_warnings)
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
    
    y_pos += 20
    audit_page.draw_line(fitz.Point(30, y_pos), fitz.Point(page_width - 30, y_pos), color=line_color, width=1.5)
    y_pos += 30
    
    # 5. Success Scenario
    if not has_issues:
        audit_page.insert_text((30, y_pos), "No issues detected. All labels matched the distribution list flawlessly.", fontname="helv", fontsize=12, color=text_dark)
        return audit_doc

    # --- HELPER: Smart Page Breaks ---
    def check_page_break(current_y, required_space):
        nonlocal audit_page, audit_doc
        if current_y + required_space > report_height - 40:
            audit_page = audit_doc.new_page(width=page_width, height=report_height)
            # Add a slim continuation header
            audit_page.draw_rect(fitz.Rect(0, 0, page_width, 15), color=bg_color, fill=bg_color)
            return 40 
        return current_y

    # 6. Missing Stores Section
    if missing_stores:
        y_pos = check_page_break(y_pos, 40)
        audit_page.insert_text((30, y_pos), "MISSING STORES:", fontname="hebo", fontsize=12, color=bg_color)
        y_pos += 20
        
        for store in missing_stores:
            y_pos = check_page_break(y_pos, 18)
            # Unpack the tuple we sent from the engines
            store_name, sig_code = store
            display_text = f"{store_name}   (Code: {sig_code})"

            audit_page.draw_circle(fitz.Point(36, y_pos - 4), 2, color=text_dark, fill=text_dark)
            audit_page.insert_text((48, y_pos), display_text, fontname="helv", fontsize=11, color=text_dark)
            y_pos += 18
            
        y_pos += 15

    # 6b. Possible Duplicate-Name Match Warnings
    if collision_warnings:
        header_height = 40
        y_pos = check_page_break(y_pos, header_height)
        header_rect = fitz.Rect(30, y_pos - 10, page_width - 30, y_pos - 10 + header_height)
        _fit_textbox(audit_page, header_rect, "PLEASE CHECK THOROUGHLY - POSSIBLE CODE GROUP ISSUES:", "hebo", 12, fitz.TEXT_ALIGN_LEFT, bg_color)
        y_pos += header_height

        for note in collision_warnings:
            note_height = 40
            y_pos = check_page_break(y_pos, note_height)
            audit_page.draw_circle(fitz.Point(36, y_pos - 4), 2, color=bg_color, fill=bg_color)
            note_rect = fitz.Rect(48, y_pos - 10, page_width - 30, y_pos - 10 + note_height)
            _fit_textbox(audit_page, note_rect, note, "helv", 10, fitz.TEXT_ALIGN_LEFT, text_dark)
            y_pos += note_height

        y_pos += 15

    # 7. Unmatched & Blank Summaries
    if unmatched_count or blank_count:
        y_pos = check_page_break(y_pos, 50)
        
        # Subtle divider line
        audit_page.draw_line(fitz.Point(30, y_pos), fitz.Point(page_width - 30, y_pos), color=line_color, width=1)
        y_pos += 25
        
        if unmatched_count:
            y_pos = check_page_break(y_pos, 20)
            audit_page.insert_text((30, y_pos), f"Unmatched {item_name.title()}:", fontname="hebo", fontsize=11, color=text_dark)
            audit_page.insert_text((170, y_pos), str(unmatched_count), fontname="hebo", fontsize=11, color=bg_color)
            y_pos += 20
            
        if blank_count:
            y_pos = check_page_break(y_pos, 20)
            audit_page.insert_text((30, y_pos), f"Blank {item_name.title()}:", fontname="hebo", fontsize=11, color=text_dark)
            audit_page.insert_text((170, y_pos), str(blank_count), fontname="hebo", fontsize=11, color=(0.6, 0.6, 0.6))

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
    other was not - the signature of a label being grouped under the wrong name."""
    expected_by_code = {}
    for store in sorted_stores:
        expected_by_code.setdefault(store_mapping[store], []).append(store)

    code_collision_notes = {}
    collision_warnings = []
    for name_a, name_b in find_name_collisions(store_mapping.keys()):
        a_found = name_a in found_stores
        b_found = name_b in found_stores
        if a_found == b_found:
            continue  # both found or both missing - not the ambiguous signal we care about

        found_name = name_a if a_found else name_b
        missing_name = name_b if a_found else name_a
        found_code = store_mapping[found_name]
        missing_code = store_mapping[missing_name]

        collision_warnings.append(
            f"'{missing_name}' (Code: {missing_code}) is missing, but similarly named '{found_name}' "
            f"(Code: {found_code}) was matched - pages may be grouped incorrectly, please check both code groups thoroughly."
        )
        code_collision_notes.setdefault(found_code, []).append(
            f"Possible mismatch with '{missing_name}' - some labels here may actually belong to Code {missing_code}."
        )
        code_collision_notes.setdefault(missing_code, []).append(
            f"Possible mismatch with '{found_name}' - check if labels were misrouted to Code {found_code}."
        )

    divider_info_by_code = {}
    for code, stores_for_code in expected_by_code.items():
        found_for_code = [s for s in stores_for_code if s in found_stores]
        missing_for_code = [s for s in stores_for_code if s not in found_stores]
        divider_info_by_code[code] = {
            'matched_count': len(found_for_code),
            'expected_count': len(stores_for_code),
            'missing': missing_for_code,
            'collision_note': " ".join(code_collision_notes.get(code, [])) or None,
        }

    return divider_info_by_code, collision_warnings


# ==========================================
# UTILITY: PROFESSIONAL DIVIDER SHEET
# ==========================================
def build_divider_sheet(page_width, page_height, signature_code, matched_count, expected_count, missing_for_code=None, collision_note=None):
    """Generates a professional divider page: green border if the number of distinct
    stores matched into this code equals the expected count from the signature links,
    red border otherwise (with a shortfall/excess disclaimer and the missing stores listed).
    Layout scales down on short pages (e.g. split-layout half-page dividers) so nothing
    overflows or gets skipped past the bottom edge."""
    missing_for_code = missing_for_code or []
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
    text = f"CODE {signature_code}\n\nMatched Labels: {matched_count}"
    _fit_textbox(page, header_rect, text, "hebo", 40, fitz.TEXT_ALIGN_CENTER, (0, 0, 0))

    # 3. Disclaimer + Missing Store List when it doesn't check out - each block only
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

        if collision_note and y + 15 < bottom_limit:
            note_rect = fitz.Rect(40, y, page_width - 40, min(y + 45, bottom_limit))
            _fit_textbox(page, note_rect, collision_note, "helv", 10, fitz.TEXT_ALIGN_CENTER, border_color)
            y += 45

        if missing_for_code and y + 15 < bottom_limit:
            list_text = "Missing Stores:\n" + "\n".join(f"- {s}" for s in missing_for_code)
            list_rect = fitz.Rect(50, y, page_width - 50, bottom_limit)
            _fit_textbox(page, list_rect, list_text, "helv", 10, fitz.TEXT_ALIGN_LEFT, (0.2, 0.2, 0.2))

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
    missing_stores = [(store, store_mapping[store]) for store in sorted_stores if store not in found_stores]
    divider_info_by_code, collision_warnings = analyze_matches(store_mapping, sorted_stores, found_stores)
    audit_doc = build_audit_report(pdf_filename, page_width, half_height, missing_stores, len(unmatched_halves), len(blank_halves), "HALVES", collision_warnings)

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
                info = divider_info_by_code.get(current_code, {'matched_count': 0, 'expected_count': 0, 'missing': [], 'collision_note': None})
                div_doc = build_divider_sheet(page_width, half_height, current_code, info['matched_count'], info['expected_count'], info['missing'], info['collision_note'])
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
    code_buckets = {}
    unmatched_pages = []
    blank_pages = []
    target = fitz.Rect(0, 0, 317.5, page_height)
    expected_extra_pages = 0

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
            current_store = match
            found_stores.add(match)
            page_info = re.search(r"page\s+(\d+)\s*/\s*(\d+)", text)
            if page_info:
                expected_extra_pages = int(page_info.group(2)) - int(page_info.group(1))
            else:
                expected_extra_pages = 0
                
        elif expected_extra_pages > 0 and current_store:
            expected_extra_pages -= 1 
            
        else:
            current_store = None 
            expected_extra_pages = 0
            
        if not current_store:
            sig_code = 'UNMATCHED'
        else:
            sig_code = store_mapping[current_store]
            
        if sig_code == 'UNMATCHED':
            unmatched_pages.append(page)
            print(f"[DEBUG] Unmatched/Junk Label on Page: {page_num + 1}")
        else:
            if sig_code not in code_buckets:
                code_buckets[sig_code] = []
            code_buckets[sig_code].append(page_num)

    # Generate Universal Audit Report
    missing_stores = [(store, store_mapping[store]) for store in sorted_stores if store not in found_stores]
    divider_info_by_code, collision_warnings = analyze_matches(store_mapping, sorted_stores, found_stores)
    audit_doc = build_audit_report(pdf_filename, page_width, page_height, missing_stores, len(unmatched_pages), len(blank_pages), "PAGES", collision_warnings)

    # Stitching
    final_shuffled_doc = fitz.open()

    if len(audit_doc) > 0:
        final_shuffled_doc.insert_pdf(audit_doc)

    # NEW: Sort by length first, then alphabetically (A, B, C... Z, AA, AB)
    sorted_codes = sorted(code_buckets.keys(), key=lambda x: (len(x), x))  # Sort by length then alphabetically
    for code in sorted_codes:
        if add_dividers:
            info = divider_info_by_code.get(code, {'matched_count': 0, 'expected_count': 0, 'missing': [], 'collision_note': None})
            div_doc = build_divider_sheet(page_width, page_height, code, info['matched_count'], info['expected_count'], info['missing'], info['collision_note'])
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