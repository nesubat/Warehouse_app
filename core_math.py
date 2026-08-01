import re

def clean_file_name(raw_string):
    """Cleans messy strings into safe file names."""
    if not raw_string:
        return "UNKNOWN_JOB"
    string_val = re.sub(r'\s+', ' ', str(raw_string)).strip()
    cleaned = re.sub(r'[^A-Za-z0-9 _-]', '', string_val)
    return cleaned.strip()

def sanitize_cell(val):
    """Cleans messy Excel data into strict numbers, text, or zero."""
    if val is None: return 0
    if isinstance(val, (int, float)): return val
    
    val_str = str(val).strip()
    if val_str in ("", "-", ".", "0", "0.0"): return 0
    
    try:
        num = float(val_str)
        return int(num) if num.is_integer() else num
    except ValueError:
        return val_str

def sort_key(sig):
    """Assigns a 3-part identity to prevent crashes: (Priority, Number, Text)"""
    keys = []
    for val in sig:
        if val == 0:
            keys.append((3, 0, ""))                 
        elif isinstance(val, (int, float)):
            keys.append((1, val, ""))               
        else:
            keys.append((2, 0, str(val).lower()))   
    return tuple(keys)

def generate_pack_signatures(raw_values, store_rows, p_start, p_end):
    """
    Takes a range of columns and rows, finds unique combinations, 
    and assigns A, B, AA signature codes.
    """
    row_signatures = []
    
    for r in store_rows:
        sig = tuple(sanitize_cell(raw_values[r-1][c-1]) for c in range(p_start, p_end + 1))
        if all(v == 0 for v in sig):
            row_signatures.append((r, None))
        else:
            row_signatures.append((r, sig))
            
    unique_sigs = list(set(sig for r, sig in row_signatures if sig is not None))
    unique_sigs.sort(key=sort_key)
    
    sig_to_letter = {}
    summary_counts = {sig: 0 for sig in unique_sigs}
    
    for r, sig in row_signatures:
        if sig is not None:
            summary_counts[sig] += 1
            
    for index, sig in enumerate(unique_sigs):
        letter = chr(65 + index) if index < 26 else chr(65 + (index // 26) - 1) + chr(65 + (index % 26))
        sig_to_letter[sig] = letter
        
    ordered_codes = []
    for r, sig in row_signatures:
        if sig is not None:
            ordered_codes.append(sig_to_letter[sig])
        else:
            ordered_codes.append("")
            
    return row_signatures, unique_sigs, sig_to_letter, summary_counts, ordered_codes
# Formatting Functions for Excel Sheets
def format_file1(sheet, col_letter, p_start, p_end, pack_group_row, job_id_row, last_row, p_name, original_color):
    """Applies standardized formatting for the newly inserted signature column in File 1."""
    target_col = sheet.range(f"{col_letter}:{col_letter}")
    
    # 1. Header Styling
    sheet.range(f"{col_letter}{job_id_row}").value = f"Code for {p_name}"
    sheet.range(f"{col_letter}{job_id_row}").color = (0, 0, 0)
    sheet.range(f"{col_letter}{job_id_row}").autofit()
    
    # 2. Reset and Align Column
    target_col.api.FormatConditions.Delete()
    target_col.color = (255, 255, 255)
    target_col.font.color = (0, 0, 0)
    target_col.api.EntireColumn.HorizontalAlignment = -4108
    
    # 3. Data Block Borders & Alignment
    start_row = pack_group_row + 1
    data_block = sheet.range(f"{col_letter}{start_row}:{col_letter}{last_row}")
    data_block.api.Borders.LineStyle = 1  
    data_block.api.Borders.Weight = 2     
    data_block.font.color = (0, 0, 0)
    data_block.font.bold = True  # FIXED: Added '= True' 
    data_block.api.HorizontalAlignment = -4108
    data_block.api.VerticalAlignment = -4108
    
    # 4. Handle Merged Pack Header
    new_pack_range = sheet.range((pack_group_row, p_start), (pack_group_row, p_end + 1))
    try:
        new_pack_range.unmerge()
    except Exception:
        pass
        
    new_pack_range.merge()
    if original_color:
        new_pack_range.color = original_color
        
    new_pack_range.api.HorizontalAlignment = -4108
    new_pack_range.api.VerticalAlignment = -4108


def format_file2(sheet, pack_group_row, p_start, p_end, write_data):
    """Applies standardized formatting for the summarized packing matrix in File 2."""
    if not write_data:
        return
        
    data_range = sheet.range(
        (pack_group_row + 1, p_start), 
        (pack_group_row + len(write_data), p_end + 2)
    )
    
    # 1. Base Font & Alignment
    data_range.value = write_data
    data_range.font.size = 20
    data_range.api.HorizontalAlignment = -4108
    data_range.api.VerticalAlignment = -4108
    data_range.font.bold = True
    data_range.font.color = (0, 0, 0)
    
    # 2. Banded Rows
    for r_offset in range(len(write_data)):
        if r_offset % 2 == 1 and r_offset != len(write_data) - 1:
            sheet.range(
                (pack_group_row + 1 + r_offset, p_start), 
                (pack_group_row + 1 + r_offset, p_end + 2)
            ).color = (220, 230, 241)
    
    # 3. Apply Borders
    for border_id in [7, 8, 9, 10, 11, 12]:
        data_range.api.Borders(border_id).LineStyle = 1 
        data_range.api.Borders(border_id).Weight = 2