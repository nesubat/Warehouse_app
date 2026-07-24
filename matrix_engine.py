import os
import shutil
import re
import openpyxl
import xlwings as xw
import pandas as pd
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string, get_column_letter

# =====================================================================
# UTILITY FUNCTIONS
# =====================================================================

def clean_file_name(raw_string):
    """
    Cleans the Job ID so Windows doesn't crash when saving the file.
    Only allows letters, numbers, spaces, hyphens, and underscores.
    """
    if not raw_string:
        return "UNKNOWN_JOB"
    cleaned = re.sub(r'[^A-Za-z0-9 _-]', '', str(raw_string))
    return cleaned.strip()

def scan_excel_tabs(file_path):
    """Quickly peeks into the file to list the tabs for the UI."""
    excel_file = pd.ExcelFile(file_path)
    excel_file.close()
    return excel_file.sheet_names



def generate_tab_map(file_path, sheet_name, start_cell, job_id_cell, store_col):
    """Reads the physical grid to show a blueprint preview."""
    wb = openpyxl.load_workbook(file_path, data_only=True)
    try:
        sheet = wb[sheet_name]
    
        # 1. ANCHORS
        start_coords = coordinate_from_string(start_cell) 
        stock_start_col = column_index_from_string(start_coords[0])
        pack_group_row = start_coords[1]
        
        # THE MAGIC FIX: Use Job ID Cell to anchor the Packs!
        job_coords = coordinate_from_string(job_id_cell)
        pack_start_col = column_index_from_string(job_coords[0])
        job_id_row = job_coords[1]
        
        store_col_idx = column_index_from_string(store_col)
        
        # Find boundaries
        max_col = sheet.max_column
        last_col_letter = get_column_letter(max_col)
            
        last_row = sheet.max_row
        while last_row > 0 and sheet.cell(row=last_row, column=store_col_idx).value is None:
            last_row -= 1
        if "total" in str(sheet.cell(row=last_row, column=store_col_idx).value).strip().lower():
            last_row -= 1
            
        total_stores = last_row - max(pack_group_row, job_id_row)
        
        # 2. FIND STOCKS (Scan right until blank)
        last_stock_col = stock_start_col
        while sheet.cell(row=pack_group_row, column=last_stock_col + 1).value is not None:
            last_stock_col += 1
            
        stock_start_ltr = get_column_letter(stock_start_col)
        stock_end_ltr = get_column_letter(last_stock_col)
            
        # 3. FIND PACKAGES (Scan right until blank)
        packages = []
        current_col = pack_start_col  
        while sheet.cell(row=pack_group_row, column=current_col).value is not None:
            cell_value = sheet.cell(row=pack_group_row, column=current_col).value
            pack_name = str(cell_value).strip()
            start_letter = get_column_letter(current_col)
            is_merged = False
            for merged_range in sheet.merged_cells.ranges:
                if (current_col >= merged_range.min_col and current_col <= merged_range.max_col and 
                    pack_group_row >= merged_range.min_row and pack_group_row <= merged_range.max_row):
                    end_letter = get_column_letter(merged_range.max_col)
                    
                    packages.append({
                        "name": pack_name, 
                        "label": f"{pack_name} spans from {start_letter} to {end_letter}"
                    })
                    current_col = merged_range.max_col + 1 
                    is_merged = True
                    break
            if not is_merged:
                packages.append({
                    "name": pack_name,
                    "label": f"{pack_name} is in column {start_letter}"
                })
                current_col += 1
                
    finally:
        wb.close()
            
    return {
        "sheet_name": sheet_name,
        "stocks_map": f"Packaging stocks are in {stock_start_ltr} to {stock_end_ltr}.",
        "packages_map": packages,
        "last_col": last_col_letter,
        "last_row": last_row,
        "total_stores": total_stores
    }
# =====================================================================
# THE CORE GENERATOR ENGINE (PHASE 1, 2, and 3)
# =====================================================================

def generate_all_outputs(file_path, original_filename, selected_tabs, user_inputs):
    base_dir = os.path.dirname(file_path)
    file_ext = os.path.splitext(file_path)[1]
    
    # ---------------------------------------------------------
    # MEMORY EXTRACTION: Read the file quickly with openpyxl 
    # to find all boundaries before we launch xlwings
    # ---------------------------------------------------------
    first_tab = selected_tabs[0]
    wb_temp = openpyxl.load_workbook(file_path, data_only=True)
    job_cell_coord = user_inputs[first_tab]["job"]
    raw_job_id = wb_temp[first_tab][job_cell_coord].value
    
    tab_data_memory = {}
    
    for tab_name in selected_tabs:
        sheet = wb_temp[tab_name]
        inputs = user_inputs.get(tab_name, {"start": "B8", "job": "E1", "store": "A", "selected_packs": []})
        
        # 1. ANCHORS
        start_coords = coordinate_from_string(inputs["start"])
        stock_start_col = column_index_from_string(start_coords[0])
        pack_group_row = start_coords[1]
        
        job_coords = coordinate_from_string(inputs["job"])
        pack_start_col = column_index_from_string(job_coords[0])
        job_id_row = job_coords[1]
        
        store_col_idx = column_index_from_string(inputs["store"])
        
        last_row = sheet.max_row
        while last_row > 0 and sheet.cell(row=last_row, column=store_col_idx).value is None:
            last_row -= 1
        if "total" in str(sheet.cell(row=last_row, column=store_col_idx).value).strip().lower():
            last_row -= 1
            
        # 2. Map Stocks (Scan right until blank)
        last_stock_col = stock_start_col
        while sheet.cell(row=pack_group_row, column=last_stock_col + 1).value is not None:
            last_stock_col += 1

        # 3. Map Packages (Scan right until blank)
        pack_ranges = []
        current_col = pack_start_col 
        while sheet.cell(row=pack_group_row, column=current_col).value is not None:
            cell_val = sheet.cell(row=pack_group_row, column=current_col).value
            start_idx = current_col
            end_idx = current_col
            for merged_range in sheet.merged_cells.ranges:
                if (current_col >= merged_range.min_col and current_col <= merged_range.max_col and 
                    pack_group_row >= merged_range.min_row and pack_group_row <= merged_range.max_row):
                    end_idx = merged_range.max_col
                    break
            pack_ranges.append({"name": str(cell_val).strip(), "start": start_idx, "end": end_idx})
            current_col = end_idx + 1

        tab_data_memory[tab_name] = {
            "job_id_row": job_id_row,
            "pack_group_row": pack_group_row,
            "last_row": last_row,
            "pack_ranges": pack_ranges,
            "store_col_idx": store_col_idx,
            "stock_start_col": stock_start_col,
            "last_stock_col": last_stock_col
            
        }
    wb_temp.close()
    # ---------------------------------------------------------
    # FILE SETUP: Create File Names and Base Clones
    # ---------------------------------------------------------
    cleaned_job_id = clean_file_name(raw_job_id)
    file1_name = f"Final {os.path.splitext(original_filename)[0]}{file_ext}"
    file2_name = f"Packing Sheet_{cleaned_job_id}{file_ext}"
    file3_name = f"Signature links_{cleaned_job_id}{file_ext}"
    
    file1_path = os.path.join(base_dir, file1_name)
    file2_path = os.path.join(base_dir, file2_name)
    file3_path = os.path.join(base_dir, file3_name)
    
    # Clone the RAW original file for File 1 and File 2 immediately
    shutil.copy(file_path, file1_path)
    shutil.copy(file_path, file2_path)
    
    # Dictionary to pass calculated summary data from Phase 1 to Phase 2
    tab_summaries = {}
    
    # Start the invisible Excel application
    app = xw.App(visible=False)
    
    try:
        # =====================================================
        # PHASE 1: GENERATE FILE 1 (Enhanced Original)
        # =====================================================
        wb1_xw = app.books.open(file1_path)
        
        for tab_name in selected_tabs:
            sheet_xw = wb1_xw.sheets[tab_name]
            tab_info = tab_data_memory[tab_name]
            raw_values = sheet_xw.used_range.value 
            tab_summaries[tab_name] = []
            # Safely fetch inputs, falling back to defaults if the web form dropped them
            inputs = user_inputs.get(tab_name, {
                "start": "B8", 
                "job": "E1", 
                "store": "A", 
                "selected_packs": []
            })
            
            # LOOP RIGHT TO LEFT to prevent insertion shifts from corrupting coordinates
            for pack in reversed(tab_info["pack_ranges"]):
                p_name = pack["name"]
                p_start = pack["start"]
                p_end = pack["end"]
                
                row_signatures = []
                store_rows = range(tab_info["pack_group_row"] + 1, tab_info["last_row"] + 1)
                
                # Extract allocation items per store
                for r in store_rows:
                    sig = tuple(raw_values[r-1][c-1] or 0 for c in range(p_start, p_end + 1))
                    
                    # If entirely blank/zeroes, do not assign a code
                    if all(v == 0 for v in sig):
                        row_signatures.append((r, None))
                    else:
                        row_signatures.append((r, sig))
                    
                unique_sigs = list(set(sig for r, sig in row_signatures if sig is not None))
                
                # SORTING: Force any '0' (blank) to act like Infinity so it sinks to the bottom
                unique_sigs.sort(key=lambda sig: tuple(float('inf') if val == 0 else val for val in sig))
                
                # Map letters and count totals
                sig_to_letter = {}
                summary_counts = {sig: 0 for sig in unique_sigs}
                
                for r, sig in row_signatures:
                    if sig is not None:
                        summary_counts[sig] += 1
                
                for index, sig in enumerate(unique_sigs):
                    letter = chr(65 + index) if index < 26 else chr(65 + (index // 26) - 1) + chr(65 + (index % 26))
                    sig_to_letter[sig] = letter
                    
                

                # Check if the user selected this pack for signatures
                
                selected_list = [p.strip() for p in inputs.get("selected_packs", [])]
                is_pack_selected = p_name.strip() in selected_list
                
                if is_pack_selected:
                    # INJECT COLUMN in File 1
                    col_letter = get_column_letter(p_start)
                    sheet_xw.range(f"{col_letter}:{col_letter}").insert('right')
                    sheet_xw.range(f"{col_letter}:{col_letter}").color = None 
                    sheet_xw.range(f"{col_letter}:{col_letter}").api.EntireColumn.AutoFit()
                    
                    # WRITE HEADERS AND LETTERS
                    sheet_xw.range(f"{col_letter}{tab_info['job_id_row']}").value = f"Code for {p_name}"
                    for r, sig in row_signatures:
                        if sig is not None:
                            sheet_xw.range(f"{col_letter}{r}").value = sig_to_letter[sig]
                            
                # [IMPORTANT] Notice how tab_summaries is OUTSIDE the 'if' statement!
                # We always save the data so Phase 2 can still build the Packing Sheet.
                tab_summaries[tab_name].append({
                    "name": p_name,
                    "start_idx": p_start,
                    "end_idx": p_end,
                    "unique_sigs": unique_sigs,
                    "counts": summary_counts,
                    "letters": sig_to_letter,
                    "is_selected": is_pack_selected # Pass this flag to Phase 2
                })

        wb1_xw.save()
        wb1_xw.close()
        
        # =====================================================
        # PHASE 2: GENERATE FILE 2 (Packing Sheet & Stocks Tab)
        # =====================================================
        # wb2_xw is already opened and cloned from the raw, unedited file!
        wb2_xw = app.books.open(file2_path)
        
        # Initialize our master bucket to hold all stock data across all tabs
        master_stock_data = []
        
        for tab_name in selected_tabs:
            sheet2 = wb2_xw.sheets[tab_name]
            tab_info = tab_data_memory[tab_name]
            pack_group_row = tab_info["pack_group_row"]
            job_id_row = tab_info["job_id_row"]
            
            # --- 1. EXTRACT STOCKS BEFORE DELETING ROWS ---
            stock_start_col = tab_info["stock_start_col"]
            for idx, raw_pack in enumerate(tab_info["pack_ranges"]):
                stock_col = stock_start_col + idx
                stock_values = sheet2.range((pack_group_row + 1, stock_col), (tab_info["last_row"], stock_col)).value
                if not isinstance(stock_values, list): stock_values = [stock_values]
                
                stock_counts = {}
                for val in stock_values:
                    if val: 
                        v_str = str(val).strip()
                        stock_counts[v_str] = stock_counts.get(v_str, 0) + 1
                        
                master_stock_data.append({
                    "header": f"{tab_name}_{raw_pack['name']}",
                    "counts": stock_counts
                })

            # --- 2. ROW WRECKING BALL (Delete Store Rows) ---
            start_del = pack_group_row + 1
            sheet2.range(f"{start_del}:1048576").api.EntireRow.Delete()
            
            # --- 3. RIGHT-TO-LEFT INSERTION LOOP ---
            summaries = tab_summaries.get(tab_name, [])
            for p_sum in summaries:
                p_name = p_sum["name"]
                
                # Fetch original coordinates (because File 2 is cloned from raw)
                raw_pack = next(p for p in tab_info["pack_ranges"] if p["name"] == p_name)
                p_start = raw_pack["start"]
                p_end = raw_pack["end"]
                
                # Insert two columns AT THE END of the pack for Count and Code
                col_letter_1 = get_column_letter(p_end + 1)
                col_letter_2 = get_column_letter(p_end + 2)
                
                sheet2.range(f"{col_letter_1}:{col_letter_2}").insert('right')
                sheet2.range(f"{col_letter_1}:{col_letter_2}").color = None
                    
                
            # 4. Write Headers (Moved Count to job_id_row as requested)
                sheet2.range((job_id_row, p_end + 1)).value = "Count"
                sheet2.range((job_id_row, p_end + 2)).value = f"Code for {p_name}"
                
                # 5. Build the Array and Calculate Totals
                write_data = []
                num_items = len(p_sum["unique_sigs"][0])
                totals = [0] * (num_items + 1) # Track totals for items + 1 for count
                
                for sig in p_sum["unique_sigs"]:
                    row_data = [val if val != 0 else None for val in sig]
                    row_data.append(p_sum["counts"][sig])   
                    row_data.append(p_sum["letters"][sig])  
                    write_data.append(row_data)
                    
                    # Accumulate totals (Corrected with multiplier)
                    for i, val in enumerate(sig):
                        if val: 
                            # Multiply the item amount by the number of stores getting this combo
                            try:
                                numeric_val = float(val)
                                totals[i] += (val * p_sum["counts"][sig])
                            except (ValueError, TypeError):
                                pass  # Skip if not a number
                    
                    # Add to the overall total store count
                    totals[-1] += p_sum["counts"][sig]
                
                # Append the Total Row at the bottom
                total_row = [t if t != 0 else None for t in totals[:-1]]
                total_row.append(totals[-1])
                total_row.append("Total") # Slap the word "Total" into the Code column slot
                write_data.append(total_row)
                    
                # 6. Paste Data and Apply Formatting
                if write_data:
                    # Define the exact block of cells we are pasting into
                    data_range = sheet2.range(
                        (pack_group_row + 1, p_start), 
                        (pack_group_row + len(write_data), p_end + 2)
                    )
                    
                    # Paste data and set Font 20
                    data_range.value = write_data
                    data_range.font.size = 20
                    
                    # Banded Rows (Light Blue = 220, 230, 241)
                    for r_offset in range(len(write_data)):
                        # If the row is an odd index (every other row), and NOT the total row
                        if r_offset % 2 == 1 and r_offset != len(write_data) - 1:
                            sheet2.range(
                                (pack_group_row + 1 + r_offset, p_start), 
                                (pack_group_row + 1 + r_offset, p_end + 2)
                            ).color = (220, 230, 241)
            # --- 4. COLUMN WRECKING BALL (Your New Request) ---
            # Delete every column from A up to the spacer column (last_stock_col + 1)
            # This isolates just the pack summaries!
            end_del_col_letter = get_column_letter(tab_info["last_stock_col"] + 1)
            sheet2.range(f"A:{end_del_col_letter}").api.EntireColumn.Delete()

        # --- 5. BUILD THE FINAL PACKAGING STOCKS TAB ---
        stock_sheet = wb2_xw.sheets.add(name="Packaging Stocks", after=wb2_xw.sheets[-1])
        out_col = 1
        
        for stock_info in master_stock_data:
            col_letter_1 = get_column_letter(out_col)
            col_letter_2 = get_column_letter(out_col + 1)
            
            # Merge and paint header green
            stock_sheet.range(f"{col_letter_1}1:{col_letter_2}1").merge()
            stock_sheet.range(f"{col_letter_1}1").value = stock_info["header"]
            stock_sheet.range(f"{col_letter_1}1").color = (217, 234, 211) 
            
            # Write stock item rows
            row_idx = 2
            for stock_name, count in stock_info["counts"].items():
                stock_sheet.range((row_idx, out_col)).value = stock_name
                stock_sheet.range((row_idx, out_col + 1)).value = count
                row_idx += 1
                
            # Autofit column widths nicely
            stock_sheet.range(f"{col_letter_1}:{col_letter_2}").api.EntireColumn.AutoFit()
            
            # Paint a 2-width Black Wall for separation
            wall_col = get_column_letter(out_col + 2)
            stock_sheet.range(f"{wall_col}:{wall_col}").color = (0, 0, 0)
            stock_sheet.range(f"{wall_col}:{wall_col}").column_width = 2
            
            # Jump 3 spaces to the right for the next pack's data
            out_col += 3
                    
        wb2_xw.save()
        wb2_xw.close()
    finally: 
        try:
            app.quit()  # Ensure Excel is closed to release file locks
            app.kill()
        except:
            pass

        # =====================================================
        # PHASE 3: GENERATE FILE 3 (Signature Links Roster)
        # =====================================================
        #cloning file1 
        shutil.copy(file1_path, file3_path)
        # Open File 3 with openpyxl 
    wb3_openpyxl = openpyxl.load_workbook(file3_path)
    
    for tab_name in selected_tabs:
        if tab_name not in wb3_openpyxl.sheetnames:
            continue
            
        sheet3 = wb3_openpyxl[tab_name]
        tab_info = tab_data_memory[tab_name]
        
        max_c = sheet3.max_column
        job_row = tab_info["job_id_row"]
        store_col = tab_info["store_col_idx"]
        pack_group_row = tab_info["pack_group_row"]
        
        # 1. Delete the middle gap (Bottom up)
        if pack_group_row > job_row + 1:
            sheet3.delete_rows(job_row + 1, pack_group_row - job_row)
            
        # 2. Delete the top gap
        if job_row > 1:
            sheet3.delete_rows(1, job_row - 1)
            
        # 3. Delete non-essential columns (Right to left)
        for c_idx in range(max_c, 0, -1):
            if c_idx == store_col:
                continue
                
            header_val = str(sheet3.cell(row=1, column=c_idx).value or "")
            
            if header_val.startswith("Code for"):
                continue
                
            sheet3.delete_cols(c_idx)
            
        # 4. Mathematical AutoFit (Fast approximation)
        for col in sheet3.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet3.column_dimensions[col_letter].width = adjusted_width

    wb3_openpyxl.save(file3_path)
    wb3_openpyxl.close()
    
    return file1_name, file2_name, file3_name