import os
import shutil
import re
import openpyxl
import xlwings as xw
import pandas as pd
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string, get_column_letter
import json


# =====================================================================
# UTILITY FUNCTIONS
# =====================================================================

def clean_file_name(raw_string):
    if not raw_string:
        return "UNKNOWN_JOB"
    # 1. Regex \s+ targets ALL weird whitespace (tabs, newlines, non-breaking spaces) and crushes them
    string_val = re.sub(r'\s+', ' ', str(raw_string)).strip()
    # 2. Strip invalid Windows filename characters
    cleaned = re.sub(r'[^A-Za-z0-9 _-]', '', string_val)
    return cleaned.strip()

def scan_excel_tabs(file_path):
    excel_file = pd.ExcelFile(file_path)
    excel_file.close()
    return excel_file.sheet_names

def generate_tab_map(file_path, sheet_name, start_cell, job_id_cell, store_col):
    wb = openpyxl.load_workbook(file_path, data_only=True)
    try:
        sheet = wb[sheet_name]
    
        start_coords = coordinate_from_string(start_cell) 
        stock_start_col = column_index_from_string(start_coords[0])
        pack_group_row = start_coords[1]
        
        job_coords = coordinate_from_string(job_id_cell)
        pack_start_col = column_index_from_string(job_coords[0])
        job_id_row = job_coords[1]
        
        raw_job_id = sheet[job_id_cell].value
        store_col_idx = column_index_from_string(store_col)
        
        true_max_col = 1
        for col in range(sheet.max_column, 0, -1):
            cell_val = sheet.cell(row=job_id_row, column=col).value
            if cell_val is not None and str(cell_val).strip() != "":
                true_max_col = col
                break
                
        max_col = true_max_col
        last_col_letter = get_column_letter(max_col)
            
        last_row = sheet.max_row
        while last_row > 0 and sheet.cell(row=last_row, column=store_col_idx).value is None:
            last_row -= 1
        if "total" in str(sheet.cell(row=last_row, column=store_col_idx).value).strip().lower():
            last_row -= 1
            
        total_stores = last_row - max(pack_group_row, job_id_row)
        
        last_stock_col = stock_start_col
        while sheet.cell(row=pack_group_row, column=last_stock_col + 1).value is not None:
            last_stock_col += 1
            
        stock_start_ltr = get_column_letter(stock_start_col)
        stock_end_ltr = get_column_letter(last_stock_col)
            
        last_pack_col = max_col
        while last_pack_col > pack_start_col and sheet.cell(row=job_id_row, column=last_pack_col).value is None:
            last_pack_col -= 1
            
        packages = []
        pack_ranges = [] 
        current_col = pack_start_col  
        
        while current_col <= last_pack_col:
            is_merged = False
            
            for merged_range in sheet.merged_cells.ranges:
                if (current_col >= merged_range.min_col and current_col <= merged_range.max_col and 
                    pack_group_row >= merged_range.min_row and pack_group_row <= merged_range.max_row):
                    
                    raw_name = sheet.cell(row=merged_range.min_row, column=merged_range.min_col).value
                    pack_name = str(raw_name).strip() if raw_name else f"Pack_{current_col}"
                    
                    start_letter = get_column_letter(merged_range.min_col)
                    end_letter = get_column_letter(merged_range.max_col)
                    
                    packages.append({
                        "name": pack_name, 
                        "label": f"{pack_name} spans from {start_letter} to {end_letter}"
                    })
                    pack_ranges.append({
                        "name": pack_name, 
                        "start": merged_range.min_col, 
                        "end": merged_range.max_col
                    })
                    
                    current_col = merged_range.max_col + 1 
                    is_merged = True
                    break
                    
            if not is_merged:
                raw_name = sheet.cell(row=pack_group_row, column=current_col).value
                if raw_name is not None and str(raw_name).strip() != "":
                    pack_name = str(raw_name).strip()
                    start_letter = get_column_letter(current_col)
                    
                    packages.append({
                        "name": pack_name,
                        "label": f"{pack_name} is in column {start_letter}"
                    })
                    pack_ranges.append({
                        "name": pack_name, 
                        "start": current_col, 
                        "end": current_col
                    })
                    
                current_col += 1
                
        backend_data = {
            "job_id_row": job_id_row,
            "pack_group_row": pack_group_row,
            "last_row": last_row,
            "pack_ranges": pack_ranges,
            "store_col_idx": store_col_idx,
            "stock_start_col": stock_start_col,
            "last_stock_col": last_stock_col,
            "raw_job_id": raw_job_id
        }
                
    finally:
        wb.close()
            
    return {
        "sheet_name": sheet_name,
        "stocks_map": f"Packaging stocks are in {stock_start_ltr} to {stock_end_ltr}.",
        "packages_map": packages,
        "last_col": last_col_letter,
        "last_row": last_row,
        "total_stores": total_stores,
        "backend_data": backend_data
    }

# =====================================================================
# THE CORE GENERATOR ENGINE (PHASE 1, 2, and 3)
# =====================================================================

def generate_all_outputs(file_path, original_filename, selected_tabs, user_inputs, tab_data_memory, project_dir):
    file_ext = os.path.splitext(file_path)[1]
    
    first_tab = selected_tabs[0]
    raw_job_id = tab_data_memory[first_tab]["raw_job_id"]
    cleaned_job_id = clean_file_name(raw_job_id)
    
    # ---------------------------------------------------------
    # MASTER SWITCH: Check if any tab has any packs selected
    # ---------------------------------------------------------
    any_packs_selected = False
    for tab in selected_tabs:
        inputs = user_inputs.get(tab, {})
        if inputs.get("selected_packs"):
            any_packs_selected = True
            break

    file1_name = f"Final {os.path.splitext(original_filename)[0]}{file_ext}" if any_packs_selected else None
    file2_name = f"Packing Sheet_{cleaned_job_id}{file_ext}"
    file3_name = f"Signature links_{cleaned_job_id}{file_ext}" if any_packs_selected else None
    
    file1_path = os.path.join(project_dir, file1_name) if any_packs_selected else None
    file2_path = os.path.join(project_dir, file2_name)
    file3_path = os.path.join(project_dir, file3_name) if any_packs_selected else None
    
    if any_packs_selected:
        shutil.copy(file_path, file1_path)
    shutil.copy(file_path, file2_path)
    
    tab_summaries = {}
    app = xw.App(visible=False)
    
    try:
        # =====================================================
        # PHASE 1 & 2: GENERATE FILE 1 & 2 SIMULTANEOUSLY 
        # =====================================================
        if any_packs_selected:
            wb1_xw = app.books.open(file1_path)
            
        wb2_xw = app.books.open(file2_path)
        master_stock_data = []
        
        for tab_name in selected_tabs:
            sheet2 = wb2_xw.sheets[tab_name]
            tab_info = tab_data_memory[tab_name]
            raw_values = sheet2.used_range.value 
            
            tab_summaries[tab_name] = []
            inputs = user_inputs.get(tab_name, {"selected_packs": []})
            selected_list = [p.strip() for p in inputs.get("selected_packs", [])]
            
            if any_packs_selected:
                sheet1_xw = wb1_xw.sheets[tab_name]
            
            # --- CALCULATE SIGNATURES & INJECT INTO FILE 1 (If Selected) ---
            for pack in reversed(tab_info["pack_ranges"]):
                p_name = pack["name"]
                p_start = pack["start"]
                p_end = pack["end"]
                
                row_signatures = []
                store_rows = range(tab_info["pack_group_row"] + 1, tab_info["last_row"] + 1)
                
                for r in store_rows:
                    sig = tuple(raw_values[r-1][c-1] or 0 for c in range(p_start, p_end + 1))
                    if all(v == 0 for v in sig):
                        row_signatures.append((r, None))
                    else:
                        row_signatures.append((r, sig))
                    
                unique_sigs = list(set(sig for r, sig in row_signatures if sig is not None))
                unique_sigs.sort(key=lambda sig: tuple(float('inf') if val == 0 else val for val in sig))
                
                sig_to_letter = {}
                summary_counts = {sig: 0 for sig in unique_sigs}
                
                for r, sig in row_signatures:
                    if sig is not None:
                        summary_counts[sig] += 1
                
                for index, sig in enumerate(unique_sigs):
                    letter = chr(65 + index) if index < 26 else chr(65 + (index // 26) - 1) + chr(65 + (index % 26))
                    sig_to_letter[sig] = letter
                    
                # Save ordered codes for Phase 3 (Pandas)
                ordered_codes = []
                for r, sig in row_signatures:
                    if sig is not None:
                        ordered_codes.append(sig_to_letter[sig])
                    else:
                        ordered_codes.append("")
                    
                is_pack_selected = p_name.strip() in selected_list
                
                if any_packs_selected and is_pack_selected:
                    col_letter = get_column_letter(p_start)
                    sheet1_xw.range(f"{col_letter}:{col_letter}").insert('right')
                    sheet1_xw.range(f"{col_letter}:{col_letter}").color = None 
                    sheet1_xw.range(f"{col_letter}:{col_letter}").api.EntireColumn.AutoFit()
                    
                    sheet1_xw.range(f"{col_letter}{tab_info['job_id_row']}").value = f"Code for {p_name}"
                    pack_group_row = tab_info["pack_group_row"]
                    pack_name_cell = sheet1_xw.range((pack_group_row, p_start+1))
                    original_color = pack_name_cell.color
                    
                    new_pack_range = sheet1_xw.range((pack_group_row, p_start), (pack_group_row, p_end + 1))
                    
                    try:
                        new_pack_range.unmerge()
                    except Exception:
                        pass
                        
                    new_pack_range.merge()
                    
                    if original_color:
                        new_pack_range.color = original_color
                        
                    new_pack_range.api.HorizontalAlignment = -4108
                    new_pack_range.api.VerticalAlignment = -4108
                    
                    # --- THE BATCH WRITE FIX ---
                    # Bundle all the letters into a 2D array (a vertical column list)
                    batch_data = []
                    for r, sig in row_signatures:
                        if sig is not None:
                            batch_data.append([sig_to_letter[sig]])
                        else:
                            batch_data.append([None])
                            
                    # Paste the entire column into Excel in exactly ONE command!
                    if batch_data:
                        start_r = row_signatures[0][0] # Grab the very first row number
                        sheet1_xw.range(f"{col_letter}{start_r}").value = batch_data
                        # setting the column size to auto-fit after the batch write
                    sheet1_xw.range(f"{col_letter}:{col_letter}").api.EntireColumn.AutoFit()
                            
                tab_summaries[tab_name].append({
                    "name": p_name,
                    "start_idx": p_start,
                    "end_idx": p_end,
                    "unique_sigs": unique_sigs,
                    "counts": summary_counts,
                    "letters": sig_to_letter,
                    "is_selected": is_pack_selected,
                    "ordered_codes": ordered_codes
                })

            # --- BUILD PACKING SHEET SUMMARY (FILE 2) ---
            pack_group_row = tab_info["pack_group_row"]
            job_id_row = tab_info["job_id_row"]
            
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

            start_del = pack_group_row + 1
            sheet2.range(f"{start_del}:1048576").api.EntireRow.Delete()
            
            summaries = tab_summaries.get(tab_name, [])
            for p_sum in summaries:
                p_name = p_sum["name"]
                
                raw_pack = next(p for p in tab_info["pack_ranges"] if p["name"] == p_name)
                p_start = raw_pack["start"]
                p_end = raw_pack["end"]
                
                col_letter_1 = get_column_letter(p_end + 1)
                col_letter_2 = get_column_letter(p_end + 2)
                
                sheet2.range(f"{col_letter_1}:{col_letter_2}").insert('right')
                sheet2.range(f"{col_letter_1}:{col_letter_2}").color = None
                    
                sheet2.range((job_id_row, p_end + 1)).value = "Count"
                sheet2.range((job_id_row, p_end + 2)).value = f"Code for {p_name}"
                
                # --- 1. AUTOFIT THE NEW COLUMNS ---
                sheet2.range(f"{col_letter_1}:{col_letter_2}").api.EntireColumn.AutoFit()
                
                
                write_data = []
                num_items = len(p_sum["unique_sigs"][0])
                totals = [0] * (num_items + 1) 
                
                for sig in p_sum["unique_sigs"]:
                    row_data = [val if val != 0 else None for val in sig]
                    row_data.append(p_sum["counts"][sig])   
                    row_data.append(p_sum["letters"][sig])  
                    write_data.append(row_data)
                    
                    for i, val in enumerate(sig):
                        if val: 
                            try:
                                numeric_val = float(val)
                                totals[i] += (val * p_sum["counts"][sig])
                            except (ValueError, TypeError):
                                pass  
                    totals[-1] += p_sum["counts"][sig]
                
                total_row = [t if t != 0 else None for t in totals[:-1]]
                total_row.append(totals[-1])
                total_row.append("Total")
                write_data.append(total_row)
                    
                if write_data:
                    data_range = sheet2.range(
                        (pack_group_row + 1, p_start), 
                        (pack_group_row + len(write_data), p_end + 2)
                    )
                    
                    data_range.value = write_data
                    data_range.font.size = 20
                    
                    # Banded Rows
                    for r_offset in range(len(write_data)):
                        if r_offset % 2 == 1 and r_offset != len(write_data) - 1:
                            sheet2.range(
                                (pack_group_row + 1 + r_offset, p_start), 
                                (pack_group_row + 1 + r_offset, p_end + 2)
                            ).color = (220, 230, 241)
                    
                    # Apply User Verified Borders!
                    for border_id in [7, 8, 9, 10, 11, 12]:
                        data_range.api.Borders(border_id).LineStyle = 1 
                        data_range.api.Borders(border_id).Weight = 2    
                            
            end_del_col_letter = get_column_letter(tab_info["last_stock_col"] + 1)
            sheet2.range(f"A:{end_del_col_letter}").api.EntireColumn.Delete()

        stock_sheet = wb2_xw.sheets.add(name="Packaging Stocks", after=wb2_xw.sheets[-1])
        out_col = 1
        
        for stock_info in master_stock_data:
            col_letter_1 = get_column_letter(out_col)
            col_letter_2 = get_column_letter(out_col + 1)
            
            stock_sheet.range(f"{col_letter_1}1:{col_letter_2}1").merge()
            stock_sheet.range(f"{col_letter_1}1").value = stock_info["header"]
            stock_sheet.range(f"{col_letter_1}1").color = (217, 234, 211) 
            
            row_idx = 2
            for stock_name, count in stock_info["counts"].items():
                stock_sheet.range((row_idx, out_col)).value = stock_name
                stock_sheet.range((row_idx, out_col + 1)).value = count
                row_idx += 1
                
            stock_sheet.range(f"{col_letter_1}:{col_letter_2}").api.EntireColumn.AutoFit()
            
            wall_col = get_column_letter(out_col + 2)
            stock_sheet.range(f"{wall_col}:{wall_col}").color = (0, 0, 0)
            stock_sheet.range(f"{wall_col}:{wall_col}").column_width = 2
            
            out_col += 3
                    
        if any_packs_selected:
            wb1_xw.save()
            wb1_xw.close()
            
        wb2_xw.save()
        wb2_xw.close()
    finally: 
        try:
            app.quit()
            app.kill()
        except:
            pass

    # =====================================================
    # PHASE 3: GENERATE FILE 3 (Pandas Rebuild Architecture)
    # =====================================================
    if any_packs_selected:
        wb_raw = openpyxl.load_workbook(file_path, data_only=True)
        
        with pd.ExcelWriter(file3_path, engine='xlsxwriter') as writer:
            for tab_name in selected_tabs:
                tab_info = tab_data_memory[tab_name]
                sheet_raw = wb_raw[tab_name]
                
                # Check if this specific tab has any selected packs
                has_codes = any(p_sum["is_selected"] for p_sum in tab_summaries.get(tab_name, []))
                if not has_codes:
                    continue # Skip writing this tab entirely if it has no codes
                
                store_col_idx = tab_info["store_col_idx"]
                start_r = tab_info["pack_group_row"] + 1
                end_r = tab_info["last_row"]
                
                # Extract Stores
                store_names = []
                for r in range(start_r, end_r + 1):
                    val = sheet_raw.cell(row=r, column=store_col_idx).value
                    store_names.append(str(val) if val is not None else "")
                
                df_dict = {"Store Name": store_names}
                
                # Extract Ordered Codes
                for p_sum in reversed(tab_summaries.get(tab_name, [])):
                    if p_sum["is_selected"]:
                        # Reverse list structure to write cleanly
                        p_name = p_sum["name"]
                        ordered_codes = p_sum["ordered_codes"]
                        # Pad or trim if lengths mismatch (failsafe)
                        if len(ordered_codes) < len(store_names):
                            ordered_codes.extend([""] * (len(store_names) - len(ordered_codes)))
                        elif len(ordered_codes) > len(store_names):
                            ordered_codes = ordered_codes[:len(store_names)]
                            
                        df_dict[f"Code for {p_name}"] = ordered_codes
                
                # Write to Excel
                df_out = pd.DataFrame(df_dict)
                df_out.to_excel(writer, sheet_name=tab_name, index=False)
                
                # Apply Formatting via xlsxwriter engine
                worksheet = writer.sheets[tab_name]
                workbook = writer.book
                format16 = workbook.add_format({'font_size': 16, 'align': 'center'})
                
                # Set Autofit Width to 25 and force Font Size 16
                worksheet.set_column(0, len(df_dict) - 1, 25, format16) 
                
        wb_raw.close()
    
    return file1_name, file2_name, file3_name