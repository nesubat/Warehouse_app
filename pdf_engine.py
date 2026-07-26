import fitz  # PyMuPDF
import os
import math

def process_standard_pdf(input_pdf_path, store_mapping, output_pdf_path):
    """
    Scans a standard/multi-page PDF, groups pages by signature code using a Memory Tracker,
    and stitches them into a final sorted PDF.
    
    :param input_pdf_path: The path to the uploaded PDF file.
    :param store_mapping: A dictionary linking stores to their codes (e.g., {'Fyshwick': 'E', 'Albury': 'B'}).
    :param output_pdf_path: Where to save the finalized, shuffled PDF.
    """
    
    # 1. LOAD THE PDF
    doc = fitz.open(input_pdf_path)
    
    # Capture the exact dimensions of the first page to guarantee sizing consistency
    first_page = doc[0]
    page_width = first_page.rect.width
    page_height = first_page.rect.height
    print(f"\n[ENGINE] Detected Page Size: {page_width} x {page_height}")
    
    # 2. THE MEMORY TRACKER SETUP
    current_store = None
    code_buckets = {}  # Will hold pages grouped by code, e.g., {'A': [0, 2], 'B': [1, 3]}

    # 3. SCAN EVERY PAGE
    for page_num in range(len(doc)):
        page = doc[page_num]
        text = page.get_text().lower() # Convert page text to lowercase for easier matching
        
        found_store = None
        
        # Cross-reference page text with our master list of store names
        for store_name in store_mapping.keys():
            # If the Excel says "Fyshwick" and the PDF says "Freedom Fyshwick", this matches it!
            if store_name.lower() in text:
                found_store = store_name
                break # Stop searching once we find a match
        
        # Apply the Memory Rule
        if found_store:
            # If we found a store, update the memory
            current_store = found_store
            
        # Determine the Signature Code for this page
        if not current_store:
            sig_code = 'UNMATCHED' # Failsafe for pages with no store before the first store is found
        else:
            sig_code = store_mapping[current_store]
            
        # Add the page number to the correct Signature Code bucket
        if sig_code not in code_buckets:
            code_buckets[sig_code] = []
            
        code_buckets[sig_code].append(page_num)
        
        # Debugging readout for the console
        store_display = current_store if current_store else "UNKNOWN"
        print(f"[SCANNER] Page {page_num + 1} -> Store: {store_display} -> Bucket: {sig_code}")

    # 4. SORT AND STITCH THE OUTPUT
    output_doc = fitz.open() # Create a brand new, empty PDF
    
    # Sort the buckets alphabetically (A, B, C, D...)
    sorted_codes = sorted(code_buckets.keys())
    
    print("\n[STITCHER] Beginning alphabetical merge...")
    for code in sorted_codes:
        pages_to_insert = code_buckets[code]
        for p_num in pages_to_insert:
            # The insert_pdf command perfectly clones the page. 
            # It retains 100% vector quality, selectable text, and the EXACT page dimensions (A4, A3, etc.)
            output_doc.insert_pdf(doc, from_page=p_num, to_page=p_num)
            print(f"Stitched Page {p_num + 1} from Bucket {code}")

    # 5. EXPORT THE FINAL FILE
    output_doc.save(output_pdf_path)
    output_doc.close()
    doc.close()
    
    print(f"\n[SUCCESS] Standard PDF saved to: {output_pdf_path}")
    return True


def process_split_pdf(input_pdf_path, store_mapping, output_pdf_path):
    """
    Scans a PDF with 2 labels per page. Slices them in half, groups them by signature code,
    and applies a Cut-and-Stack imposition so they guillotine perfectly in sequence.
    """
    
    # 1. LOAD THE PDF AND MEASURE
    doc = fitz.open(input_pdf_path)
    first_page = doc[0]
    
    page_width = first_page.rect.width
    page_height = first_page.rect.height
    
    # Define the digital "Crop Boxes" (Top Half and Bottom Half)
    top_rect = fitz.Rect(0, 0, page_width, page_height / 2)
    bottom_rect = fitz.Rect(0, page_height / 2, page_width, page_height)
    
    # We will store every valid half-page as a dictionary inside this array
    all_extracted_halves = []

    # 2. SCAN AND SLICE EVERY PAGE
    for page_num in range(len(doc)):
        page = doc[page_num]
        
        # Read the text strictly within the top half
        top_text = page.get_text("text", clip=top_rect).lower()
        # Read the text strictly within the bottom half
        bottom_text = page.get_text("text", clip=bottom_rect).lower()
        
        # Check Top Half
        for store_name in store_mapping.keys():
            if store_name.lower() in top_text:
                sig_code = store_mapping[store_name]
                # Save it to our master array
                all_extracted_halves.append({
                    'page_num': page_num, 
                    'half': 'top', 
                    'code': sig_code
                })
                break
                
        # Check Bottom Half
        for store_name in store_mapping.keys():
            if store_name.lower() in bottom_text:
                sig_code = store_mapping[store_name]
                all_extracted_halves.append({
                    'page_num': page_num, 
                    'half': 'bottom', 
                    'code': sig_code
                })
                break

    # 3. SORT ALPHABETICALLY BY SIGNATURE CODE
    # This takes our flat array and sorts it strictly by the 'code' (A -> Z)
    ordered_halves = sorted(all_extracted_halves, key=lambda x: x['code'])
    
    # 4. CUT AND STACK MATH
    total_labels = len(ordered_halves)
    # math.ceil ensures that if we have 47 labels, the top stack gets 24 and bottom gets 23
    halfway_point = math.ceil(total_labels / 2)
    
    top_stack = ordered_halves[:halfway_point]
    bottom_stack = ordered_halves[halfway_point:]
    
    # 5. THE RE-STITCH (New A4 Pages)
    output_doc = fitz.open()
    
    print(f"\n[CUT & STACK] Processing {total_labels} total labels...")
    
    for i in range(halfway_point):
        # Create a brand new blank page exactly matching the original size
        new_page = output_doc.new_page(width=page_width, height=page_height)
        
        # --- PASTE LABEL ON TOP HALF ---
        top_item = top_stack[i]
        src_page = doc[top_item['page_num']]
        # Determine if we need to grab the top or bottom clip from the source
        src_clip = top_rect if top_item['half'] == 'top' else bottom_rect
        # Paste it cleanly onto the top half of our new page
        new_page.show_pdf_page(top_rect, doc, src_page.number, clip=src_clip)
        
        # --- PASTE LABEL ON BOTTOM HALF (If available) ---
        if i < len(bottom_stack):
            bot_item = bottom_stack[i]
            src_page_b = doc[bot_item['page_num']]
            src_clip_b = top_rect if bot_item['half'] == 'top' else bottom_rect
            # Paste it cleanly onto the bottom half of our new page
            new_page.show_pdf_page(bottom_rect, doc, src_page_b.number, clip=src_clip_b)

    # 6. EXPORT
    output_doc.save(output_pdf_path)
    output_doc.close()
    doc.close()
    
    print(f"[SUCCESS] Cut-and-Stack PDF saved to: {output_pdf_path}")
    return True