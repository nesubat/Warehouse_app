import fitz  # PyMuPDF
import os
import math
import re

def process_and_shuffle_pdf(input_pdf_path, store_mapping, output_pdf_path, signature_header="Unknown Header"):
    """
    Auto-detects layout (Standard vs Split), applies Longest-Match-First tracking,
    sorts/stitches the labels, and generates an Audit Report if missing/unmatched stores are found.
    """
    
    # ==========================================
    # PART 1: SETUP & SORTING
    # ==========================================
    # Sort stores from Longest to Shortest to prevent similar names overlapping
    sorted_stores = sorted(store_mapping.keys(), key=len, reverse=True)
    
    found_stores = set()
    
    doc = fitz.open(input_pdf_path)
    first_page = doc[0]
    
    # ==========================================
    # PART 2: AUTO-DETECT LAYOUT & PRE-CALCULATE
    # ==========================================
    page_width = first_page.rect.width
    page_height = first_page.rect.height
    is_split_layout = page_width < page_height
    # 2. GET THE SAFE FILE NAME AND UPDATE THE PRINT STATEMENT
    pdf_filename = os.path.basename(input_pdf_path).strip()
    
    print(f"\n[ENGINE] Processing {pdf_filename} against Signature: '{signature_header}'")
    print(f"[ENGINE] Detected Page Size: {page_width} x {page_height} (Split Layout: {is_split_layout})")

    # Variables for Split Mode
    all_extracted_halves = []

    half_height = page_height / 2
    top_rect = fitz.Rect(0, 0, page_width, half_height)
    bottom_rect = fitz.Rect(0, half_height, page_width, page_height)
    
    if is_split_layout:
        top_target = fitz.Rect(0, 0, 289, half_height)
        bottom_target = fitz.Rect(0, half_height, 289, page_height)
    else:
        target = fitz.Rect(0, 0, 317.5, page_height)

    # Variables for Standard Mode (Memory Tracker)
    current_store = None
    code_buckets = {}
    

    # ==========================================
    # PART 3A: SCAN EVERY PAGE 
    # ==========================================
    # --- TRACK A: SPLIT LAYOUT LOOP ---
    if is_split_layout:

        unmatched_halves = []
        blank_halves = []
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

            # --- EVALUATE BOTTOM HALF ---
            if bottom_match:
                found_stores.add(bottom_match)
                all_extracted_halves.append({'doc_type': 'original', 'page_num': page_num, 'half': 'bottom', 'code': store_mapping[bottom_match]})
            elif not bottom_text.strip():
                blank_halves.append({'doc_type': 'original', 'page_num': page_num, 'half': 'bottom'})
                print(f"[DEBUG] BLANK Bottom Half detected on Page {page_num + 1}")
            else:
                unmatched_halves.append({'doc_type': 'original', 'page_num': page_num, 'half': 'bottom'})
                print(f"[DEBUG] Unmatched Bottom Label on Page {page_num + 1}")
        # ==========================================
        # PART 4A: BUILD HALF-SIZED AUDIT REPORT
        # ==========================================
        audit_doc = fitz.open()
        audit_halves = []
        missing_stores = [store for store in sorted_stores if store not in found_stores]
        audit_page = audit_doc.new_page(width=page_width, height=half_height) 
        try:
            audit_page.insert_font(fontname="calibri", fontfile="C:/Windows/Fonts/calibri.ttf")
            font_style = "calibri"
        except:
            font_style = "helv" 

        y_pos = 30
        audit_page.insert_text((30, y_pos), f"WARNING: REPORT for ({pdf_filename})", fontsize=18, fontname=font_style, color=(1, 0, 0)) 
        y_pos += 30
        if not blank_halves and not unmatched_halves and not missing_stores:
            audit_page.insert_text((30, y_pos), "No issues detected. All labels matched successfully.", fontsize=14, fontname=font_style)
            y_pos += 15
        else:
        
            # --- 1. MISSING STORES ---
            if missing_stores:
                audit_page.insert_text((30, y_pos), "MISSING STORES:", fontsize=14, fontname=font_style)
                y_pos += 20
                for store in missing_stores:
                    # Check the ruler BEFORE writing the store name
                    if y_pos > half_height - 30:
                        audit_page = audit_doc.new_page(width=page_width, height=half_height)
                        y_pos = 30
            
                    audit_page.insert_text((50, y_pos), f"- {store}", fontsize=12, fontname=font_style)
                    y_pos += 15
            
            # --- 2. UNMATCHED HALVES ---
            if unmatched_halves:
                y_pos += 15 # Add a little blank space between sections
            
                # Check the ruler BEFORE writing the heading
                if y_pos > half_height - 30:
                    audit_page = audit_doc.new_page(width=page_width, height=half_height)
                    y_pos = 30
            
                audit_page.insert_text((30, y_pos), f"UNMATCHED HALVES: {len(unmatched_halves)}", fontsize=14, fontname=font_style)
                y_pos += 15
            
            # --- 3. BLANK HALVES ---
            if blank_halves:
                y_pos += 15 
            
                # Check the ruler BEFORE writing the heading
                if y_pos > half_height - 30:
                    audit_page = audit_doc.new_page(width=page_width, height=half_height)
                    y_pos = 30
        
                audit_page.insert_text((30, y_pos), f"BLANK HALVES: {len(blank_halves)}", fontsize=14, fontname=font_style)

        
        # Register all generated audit pages into our master stacking system
        for i in range(len(audit_doc)):
            audit_halves.append({'doc_type': 'audit', 'page_num': i, 'half': 'full'})

        # ==========================================
        # PART 5A: MASTER CUT-AND-STACK STITCHING
        # ==========================================
        final_shuffled_doc = fitz.open()
        
        # Sort the matched ones alphabetically A -> Z
        ordered_matched = sorted(all_extracted_halves, key=lambda x: x['code'])
        
        # Combine everything into one giant list in exact order
        master_stack = audit_halves + ordered_matched + unmatched_halves + blank_halves
        
        total_halves = len(master_stack)
        halfway_point = math.ceil(total_halves / 2)
        
        top_stack = master_stack[:halfway_point]
        bottom_stack = master_stack[halfway_point:]
        
        print(f"\n[CUT & STACK] Processing {total_halves} total halves (including Audit/Blanks)...")
        
        # The clip for audit pages since they are already half-sized
        audit_rect = fitz.Rect(0, 0, page_width, half_height)
        
        for i in range(halfway_point):
            new_page = final_shuffled_doc.new_page(width=page_width, height=page_height)
            
            # --- PASTE TOP HALF ---
            top_item = top_stack[i]
            if top_item['doc_type'] == 'audit':
                new_page.show_pdf_page(top_rect, audit_doc, top_item['page_num'], clip=audit_rect)
            else:
                src_page = doc[top_item['page_num']]
                src_clip = top_rect if top_item['half'] == 'top' else bottom_rect
                new_page.show_pdf_page(top_rect, doc, src_page.number, clip=src_clip)
            
            # --- PASTE BOTTOM HALF ---
            if i < len(bottom_stack):
                bot_item = bottom_stack[i]
                if bot_item['doc_type'] == 'audit':
                    new_page.show_pdf_page(bottom_rect, audit_doc, bot_item['page_num'], clip=audit_rect)
                else:
                    src_page_b = doc[bot_item['page_num']]
                    src_clip_b = top_rect if bot_item['half'] == 'top' else bottom_rect
                    new_page.show_pdf_page(bottom_rect, doc, src_page_b.number, clip=src_clip_b)

        # Clean up the temporary audit doc from memory
        audit_doc.close()

    # ==========================================
    # PART 3B: SCAN EVERY PAGE 
    # ==========================================
    # --- TRACK B: STANDARD LAYOUT LOOP ---
    else:
        
        unmatched_pages=[]
        blank_pages=[]
        for page_num in range(len(doc)):
            page = doc[page_num]
            
            # --- BLANK PAGE TRAP ---
            # Grab all text on the page and strip away empty spaces
            full_text = page.get_text("text").strip()
            if not full_text:
                print(f"[DEBUG] Blank Page detected on Page {page_num + 1}. Skipping.")
                blank_pages.append(page)
                # 'continue' instantly jumps to the next page in the loop. 
                # The countdown timer is NEVER touched!
                continue 
            
            # --- STANDARD LOGIC ---
            text = page.get_text("text", clip=target).lower()
            match = next((store for store in sorted_stores if store.lower() in text), None)
            
            # SCENARIO 1: We found a store name! 
            if match:
                current_store = match
                found_stores.add(match)
                
                # Look for the exact pattern "Page X/Y" 
                # \s+ means "one or more spaces"
                page_info = re.search(r"page\s+(\d+)\s*/\s*(\d+)", text)
                
                if page_info:
                    current_page_num = int(page_info.group(1))
                    total_pages = int(page_info.group(2))
                    expected_extra_pages = total_pages - current_page_num
                else:
                    expected_extra_pages = 0
                    
            # SCENARIO 2: No store name, BUT expecting more pages (Countdown)
            elif expected_extra_pages > 0 and current_store:
                expected_extra_pages -= 1 
                
            # SCENARIO 3: No store name, AND timer is 0. Wipe memory.
            else:
                current_store = None 
                expected_extra_pages = 0
                
            # --- ASSIGN TO BUCKETS ---
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


        # ==========================================
        # PART 4B: BUILD AUDIT REPORT (STANDARD)
        # ==========================================
        audit_doc = fitz.open()
        audit_pages = []
        missing_stores = [store for store in sorted_stores if store not in found_stores]
        audit_page = audit_doc.new_page(width=page_width, height=page_height) 
        
        # FIXED: Font setup is isolated cleanly
        try:
            audit_page.insert_font(fontname="calibri", fontfile="C:/Windows/Fonts/calibri.ttf")
            font_style = "calibri"
        except:
            font_style = "helv" 
        
        y_pos = 30
        audit_page.insert_text((30, y_pos), f"WARNING: REPORT for ({pdf_filename})", fontsize=18, fontname=font_style, color=(1, 0, 0)) 
        y_pos += 30
        
        # Check if there are actually any issues to report
        if not missing_stores and not unmatched_pages and not blank_pages:
            audit_page.insert_text((30, y_pos), "No issues detected. All labels matched successfully.", fontsize=14, fontname=font_style)
            y_pos += 15
            
        else:
            # --- 1. MISSING STORES ---
            if missing_stores:
                audit_page.insert_text((30, y_pos), "MISSING STORES:", fontsize=14, fontname=font_style)
                y_pos += 20
                for store in missing_stores:
                    # Check the ruler BEFORE writing the store name
                    if y_pos > page_height - 30:
                        audit_page = audit_doc.new_page(width=page_width, height=page_height)
                        y_pos = 30
    
                    audit_page.insert_text((50, y_pos), f"- {store}", fontsize=12, fontname=font_style)
                    y_pos += 15
    
            # --- 2. UNMATCHED PAGES ---
            if unmatched_pages:
                y_pos += 15 # Add a little blank space between sections
    
                # Check the ruler BEFORE writing the heading
                if y_pos > page_height - 30:
                    audit_page = audit_doc.new_page(width=page_width, height=page_height)
                    y_pos = 30
    
                audit_page.insert_text((30, y_pos), f"UNMATCHED PAGES: {len(unmatched_pages)}", fontsize=14, fontname=font_style)
                y_pos += 15
    
            # --- 3. BLANK PAGES ---
            if blank_pages:
                y_pos += 15 
    
                # Check the ruler BEFORE writing the heading
                if y_pos > page_height - 30:
                    audit_page = audit_doc.new_page(width=page_width, height=page_height)
                    y_pos = 30
    
                audit_page.insert_text((30, y_pos), f"BLANK PAGES: {len(blank_pages)}", fontsize=14, fontname=font_style)

            # Register all generated audit pages into our master stacking system
        for i in range(len(audit_doc)):
            audit_pages.append({'doc_type': 'audit', 'page_num': i})
        
    
    # ==========================================
    # PART 5B: MASTER STITCHING (STANDARD)
    # ==========================================
        final_shuffled_doc = fitz.open()
        
        # 1. APPEND AUDIT REPORT (First)
        # If the audit_doc has pages, paste them at the very beginning
        if len(audit_doc) > 0:
            final_shuffled_doc.insert_pdf(audit_doc)
            print(f"[STITCHER] Appended {len(audit_doc)} Audit Page(s).")
            
        # 2. APPEND MATCHED & SORTED PAGES (Second)
        sorted_codes = sorted(code_buckets.keys())
        print("\n[STITCHER] Beginning alphabetical merge...")
        for code in sorted_codes:
            for p_num in code_buckets[code]:
                final_shuffled_doc.insert_pdf(doc, from_page=p_num, to_page=p_num)
                print(f"Stitched Page {p_num + 1} from Bucket {code}")

        # 3. APPEND UNMATCHED PAGES (Third)
        if unmatched_pages:
            unmatched_doc = fitz.open()
            for p in unmatched_pages:
                unmatched_doc.insert_pdf(p.parent, from_page=p.number, to_page=p.number)
            final_shuffled_doc.insert_pdf(unmatched_doc)
            unmatched_doc.close()
            print(f"[STITCHER] Appended {len(unmatched_pages)} Unmatched Page(s) to the back.")
            
        # 4. APPEND BLANK PAGES (Fourth/Last)
        if blank_pages:
            blank_doc = fitz.open()
            for p in blank_pages:
                blank_doc.insert_pdf(p.parent, from_page=p.number, to_page=p.number)
            final_shuffled_doc.insert_pdf(blank_doc)
            blank_doc.close()
            print(f"[STITCHER] Appended {len(blank_pages)} Blank Page(s) to the very back.")

        # Clean up the temporary audit doc from memory
        audit_doc.close()

    # ==========================================
    # PART 6: EXPORT
    # ==========================================
    # Scrub the hidden page label metadata so Adobe counts sequentially!
    final_shuffled_doc.set_page_labels([])
    final_shuffled_doc.save(output_pdf_path)
    final_shuffled_doc.close()
    doc.close()
    
    print(f"\n[SUCCESS] Document saved to: {output_pdf_path}")
    return True