import os
import re
import time
import xlwings as xw
import pandas as pd

# Extra tab in the Signature Links file (File 3) holding the job numbers behind every code of
# every pack in an installer tab. Only written when at least one tab is marked as going to an
# installer, and skipped by the Label Shuffler when it lists tabs/packs.
DIVIDER_BARCODE_SHEET = "Divider Barcodes"
DIVIDER_BARCODE_COLUMNS = ["Tab", "Pack", "Code", "Job Number", "Kind", "Qty", "Barcode"]

def clean_file_name(raw_string):
    """Cleans messy strings into safe file names."""
    if not raw_string:
        return "UNKNOWN_JOB"
    string_val = re.sub(r'\s+', ' ', str(raw_string)).strip()
    cleaned = re.sub(r'[^A-Za-z0-9 _-]', '', string_val)
    return cleaned.strip()
def clean_store_name(value):
    if value is None or str(value).strip().lower() == 'nan':
        return ""
    # Strip trailing period, hyphen and asteriks, keeping any periods elsewhere in the name
    cleaned = str(value).strip().rstrip('.*-')
    # Collapse multiple internal spaces into a single space
    return re.sub(r'\s+', ' ', cleaned).strip()

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

def format_quantity(val):
    """1.0 -> 1, "2.0" -> 2; anything that isn't a number is handed back as trimmed text."""
    try:
        num = float(val)
    except (ValueError, TypeError):
        return str(val).strip()
    return int(num) if num.is_integer() else num

def read_divider_barcodes(xls):
    """Reads the Divider Barcodes tab of an open Signature Links pd.ExcelFile into
    {(tab, "Code for <pack>"): {code: [{"job", "kind", "qty", "barcode"}, ...]}}, keyed the same
    way the Label Shuffler names its tab/pack uploads. Returns {} when the tab isn't there
    (i.e. no tab in this campaign is going to an installer)."""
    if DIVIDER_BARCODE_SHEET not in xls.sheet_names:
        return {}
    # dtype=str + keep_default_na=False so codes like "NA" and pack names like "1" stay exactly as written
    df = pd.read_excel(xls, sheet_name=DIVIDER_BARCODE_SHEET, dtype=str, keep_default_na=False)
    lookup = {}
    for _, row in df.iterrows():
        tab, pack, code, job = (str(row[c]).strip() for c in ("Tab", "Pack", "Code", "Job Number"))
        if not (tab and pack and code and job):
            continue
        lookup.setdefault((tab, f"Code for {pack}"), {}).setdefault(code, []).append({
            "job": job,
            "kind": format_quantity(row["Kind"]) if str(row["Kind"]).strip() else None,
            "qty": format_quantity(row["Qty"]),
            "barcode": str(row["Barcode"]).strip() or job,
        })
    return lookup

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
def format_file1(sheet, col_letter, p_start, p_end, pack_group_row, job_id_row, last_row):
    """Applies standardized formatting for the newly inserted signature column in File 1."""
    target_col = sheet.range(f"{col_letter}:{col_letter}")
    
    # 1. Header Styling
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
    data_range.rows.autofit()
    
    
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
        data_range.api.Borders(border_id).Weight = 3

def build_initial_metadata(tab_info, inputs, selected_list, any_packs_selected):
    """Builds the foundational JSON map tracking exact Excel column coordinates."""
    tab_meta = {
        "user_inputs": inputs,
        "store_col_idx": tab_info["store_col_idx"],
        "pack_group_row": tab_info["pack_group_row"],
        "job_id_row": tab_info["job_id_row"],
        "last_row": tab_info["last_row"],
        "packs": {}
    }
    
    current_shift = 0
    # Read left-to-right to track the cascading column shifts
    for pack in tab_info["pack_ranges"]:
        p_name = pack["name"]
        is_selected = p_name.strip() in selected_list
        
        if any_packs_selected and is_selected:
            final_start = pack["start"] + current_shift + 1
            final_end = pack["end"] + current_shift + 1
            code_col = pack["start"] + current_shift 
            current_shift += 1
        else:
            final_start = pack["start"] + current_shift
            final_end = pack["end"] + current_shift
            code_col = None
            
        tab_meta["packs"][p_name] = {
            "original_start": pack["start"],
            "original_end": pack["end"],
            "current_start": final_start,  # We use 'current' so we can update it later
            "current_end": final_end,
            "code_col": code_col,
            "is_selected": is_selected,
            "sub_groups": {} # An empty dictionary ready for infinite sub-grouping!
        }
        
    return tab_meta

def update_metadata_for_subgroup(metadata, tab_name, parent_pack_name, sub_group_name, insert_col_idx, sub_start, sub_end):
    """
    (For Stage 4) Shifts all columns in the metadata to the right by 1 to make room 
    for a newly inserted sub-group column.
    """
    tab_meta = metadata["tabs"][tab_name]
    
    # 1. Shift all existing packs to the right of the insertion
    for p_name, p_data in tab_meta["packs"].items():
        if p_data["current_start"] >= insert_col_idx:
            p_data["current_start"] += 1
            p_data["current_end"] += 1
            if p_data["code_col"] and p_data["code_col"] >= insert_col_idx:
                p_data["code_col"] += 1
        elif p_data["current_start"] < insert_col_idx <= p_data["current_end"]:
            # The insertion happened INSIDE this pack's boundaries
            p_data["current_end"] += 1
            
        # 2. Shift any previously created sub-groups
        for sg_name, sg_data in p_data.get("sub_groups", {}).items():
            if sg_data["current_start"] >= insert_col_idx:
                sg_data["current_start"] += 1
                sg_data["current_end"] += 1
                if sg_data["code_col"] and sg_data["code_col"] >= insert_col_idx:
                    sg_data["code_col"] += 1
            elif sg_data["current_start"] < insert_col_idx <= sg_data["current_end"]:
                sg_data["current_end"] += 1

    # 3. Register the brand new subgroup under its parent
    tab_meta["packs"][parent_pack_name]["sub_groups"][sub_group_name] = {
        "current_start": sub_start + 1, # +1 because it's shifted by its own code column
        "current_end": sub_end + 1,
        "code_col": insert_col_idx
    }
    return metadata
# function to generate the next stage filenames based on the current filename
def get_next_stage_filenames(current_filename):
    """
    Parses the current filename and increments the Stage number.
    Returns the new Excel filename and the matching JSON filename.
    """
    match = re.match(r"Stage - (\d+) - (.+)", current_filename)
    
    if match:
        current_stage = int(match.group(1))
        next_stage = current_stage + 1
        base_name = match.group(2)
        new_excel_name = f"Stage - {next_stage} - {base_name}"
    else:
        new_excel_name = f"Stage - 1 - {current_filename}"
        
    base_no_ext, _ = os.path.splitext(new_excel_name)
    new_json_name = f"{base_no_ext}.json"
    
    return new_excel_name, new_json_name

def close_if_open_elsewhere(file_path):
    """If this file is already open in another running Excel instance (e.g. the
    user has it open to look at it), force that copy closed first so our own
    automation gets exclusive write access. Any unsaved edits in that open
    copy are discarded - otherwise Excel silently hands us a read-only handle
    (or the OS refuses to move/rename the file at all) and our own edits/save
    never reach the file the user is actually looking at.
    """
    target = os.path.normcase(os.path.normpath(file_path))
    same = lambda path: os.path.normcase(os.path.normpath(path)) == target
    name = os.path.basename(file_path)
    closed = 0
    seen_apps = set()
    for window in _excel_windows():
        try:
            # A Protected View window ("Enable Editing" bar) knows only its source file
            source = os.path.join(window.SourcePath, window.SourceName)
        except Exception:
            source = None
        try:
            if source is not None:
                if same(source):
                    print(f"[DEBUG] '{name}' is open in Protected View — closing it.")
                    closed += _close_with_retry(window.Close, file_path)
                continue
            excel = window.Application
            if excel.Hwnd in seen_apps:
                continue
            seen_apps.add(excel.Hwnd)
            books = excel.Workbooks
            for i in range(books.Count, 0, -1):
                book = books.Item(i)
                if same(book.FullName):
                    print(f"[DEBUG] '{name}' is already open in Excel — closing it without saving.")
                    closed += _close_with_retry(lambda: book.Close(SaveChanges=False), file_path)
            views = excel.ProtectedViewWindows
            for i in range(views.Count, 0, -1):
                view = views.Item(i)
                if same(os.path.join(view.SourcePath, view.SourceName)):
                    print(f"[DEBUG] '{name}' is open in Protected View — closing it.")
                    closed += _close_with_retry(view.Close, file_path)
        except Exception as e:
            print(f"[WARNING] Skipped an Excel window while looking for {name}: {e}")
    if closed:
        _wait_until_released(file_path)
    return closed


def _excel_windows():
    """COM objects for every Excel window on screen or hidden, across all Excel processes.

    xlwings' own list (xw.apps) looks at one window per process and crashes when that window is
    a Protected View window, so the windows are walked here directly and each is handled on its own."""
    try:
        import pythoncom, win32gui
        from ctypes import byref, windll
        from win32com.client import Dispatch
        from xlwings import _xlwindows as xlw
    except ImportError:
        return
    pythoncom.CoInitialize()
    hwnd = windll.user32.GetTopWindow(None)
    while hwnd:
        try:
            desk = win32gui.FindWindowEx(hwnd, 0, "XLDESK", None)
            child = win32gui.FindWindowEx(desk, 0, "EXCEL7", None) if desk else 0
            if child:
                ptr = xlw.accessible_object_from_window(child)
                yield Dispatch(xlw._PyCom_PyObjectFromIUnknown(ptr, byref(xlw._IDISPATCH_GUID), True))
        except Exception:
            pass
        hwnd = windll.user32.GetWindow(hwnd, 2)  # next window


def _close_with_retry(close, file_path, attempts=6):
    """Excel refuses calls while it's busy (e.g. someone is typing in a cell), so try a few times."""
    for attempt in range(attempts):
        try:
            close()
            return 1
        except Exception as e:
            if attempt == attempts - 1:
                print(f"[WARNING] Could not close the already-open copy of {file_path}: {e}")
            time.sleep(0.5)
    return 0


def _wait_until_released(file_path, timeout=5.0):
    """Windows can hold Excel's lock for a moment after the workbook closes."""
    end = time.time() + timeout
    while time.time() < end:
        try:
            with open(file_path, 'r+b'):
                return True
        except PermissionError:
            time.sleep(0.2)
        except OSError:
            return True  # file gone or not openable for another reason; nothing to wait for
    return False


def get_available_project_files(project_dir):
    """
    Scans the project directory for JSON files.
    Filters out any system files if necessary.
    """
    if not os.path.exists(project_dir):
        return []
        
    valid_files = []
    for file in os.listdir(project_dir):
        if file.endswith(".json"):
            # Optional: Skip system JSONs if you have them (like user_settings.json)
            if file not in ["user_settings.json", "app_config.json"]:
                valid_files.append(file)
                
    # Sort them so "Stage - 1", "Stage - 2", etc., appear in order
    valid_files.sort()
    return valid_files