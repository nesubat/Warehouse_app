import time
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, send_file
import os
import shutil
import json
import pymupdf as fitz
import io
import re
import sys
import stat
import socket
import subprocess
import pprint
import pandas as pd
from werkzeug.utils import secure_filename
from pdf_engine import process_and_shuffle_pdf
from matrix_engine import clean_file_name, scan_excel_tabs, generate_tab_map, generate_all_outputs, convert_legacy_excel_to_xlsx
from core_math import clean_file_name, get_available_project_files, close_if_open_elsewhere, save_if_open_elsewhere, clean_store_name, DIVIDER_BARCODE_SHEET, read_divider_barcodes
from subgroup_engine import execute_subgroups, SubgroupValidationError
from packing_label_generator import (parse_packing_data, generate_packing_labels, PackCheckError, _norm, MAPPABLE_FIELDS,
                                     detect_columns, header_columns, similar_addresses)
from openpyxl.utils.cell import column_index_from_string, get_column_letter
from address_book import AddressBook, AddressBookError
from address_import import build_review, apply_review, all_label_maps, ImportProblem
from address_match import suggestions as closest_addresses, name_score
from packing_specs import SpecStore, SpecError, parse_formula, FORMULA_PREFIXES, FORMULA_LABELS, FORMULA_EXAMPLES, ITEM_TYPES
from courier_import import read_allocation, scan_columns, mapping_fields, reference_in_sheets, PACK_FIELD, \
    FIELD_NAMES as IMPORT_FIELD_NAMES
from label_stitcher import plan_stitch, render_stitch, match_session, packing_slots, StitchError, render_label
import threading
from collections import Counter
import secrets
from courier_export import (build_consignments, detect_series, write_label_map, one_line, DEFAULT_FIXED,
                            write_courier_csv, open360_item_references, clean_export, load_export_defaults,
                            save_export_defaults, EXPORT_DEFAULTS, EXPORT_COLUMNS,
                            open360_items, write_open360_csv,
                            source_destinations, sendable, fill_store_names,
                            EDITABLE_FIELDS, SERVICE_CODE_SET, load_service_usage, service_options, record_service_usage)
import openpyxl
import app_log
import logging





if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

        

app_log.setup(BASE_DIR)  # terminal + logs/warehouse.log (app_log.py)
log_pack, log_stitch, log_book = app_log.get('packing'), app_log.get('stitch'), app_log.get('book')
log_specs = app_log.get('specs')
log_import = app_log.get('import')
# The matching page's "finished yet?" checks and label pictures aren't worth a terminal line each
logging.getLogger('werkzeug').addFilter(lambda r: not re.search(r'"GET /stitch/[^ ]*/(status|label/)', r.getMessage()))

PROJECTS_FOLDER = os.path.join(BASE_DIR, 'projects')
SERVICE_USAGE_FILE = os.path.join(BASE_DIR, 'data', 'service_code_usage.json')
EXPORT_FILE = os.path.join(BASE_DIR, 'data', 'export_details.json')  # OpenFreight export details: saved defaults
ADDRESS_BOOK = AddressBook(os.path.join(BASE_DIR, 'data', 'address_book.db'))
SPEC_STORE = SpecStore(os.path.join(BASE_DIR, 'data', 'packing_specs.json'))
os.makedirs(PROJECTS_FOLDER, exist_ok=True)
temp_dir = os.path.join(BASE_DIR, 'temp_pdf_engine')
os.makedirs(temp_dir, exist_ok=True)

def force_delete_upload(file_path):
    """Deletes an abandoned upload, closing it in Excel first (unsaved edits there are discarded)."""
    try:
        close_if_open_elsewhere(file_path)
    except Exception as e:
        print(f"[WARNING] Could not check Excel for {file_path}: {e}")
    for _ in range(10):  # Excel can hold the lock for a moment after closing
        try:
            os.remove(file_path)
            break
        except FileNotFoundError:
            break
        except OSError:
            time.sleep(0.3)
    else:
        print(f"[WARNING] Could not delete {file_path}; the startup cleanup will retry.")
    lock_file = os.path.join(os.path.dirname(file_path), "~$" + os.path.basename(file_path))
    try:
        os.remove(lock_file)
    except OSError:
        pass


def is_abandoned_project(folder_path):
    """Nothing was ever generated here: the folder holds only spreadsheets (the uploaded input).
    Every real project has a generated .pdf or .json next to them."""
    files = [f for f in os.listdir(folder_path) if not f.startswith('~$')]
    return all(f.lower().endswith(('.xlsx', '.xls')) and os.path.isfile(os.path.join(folder_path, f)) for f in files)


def force_delete_project(folder_path):
    for f in os.listdir(folder_path):
        if not f.startswith('~$'):
            force_delete_upload(os.path.join(folder_path, f))
    shutil.rmtree(folder_path, ignore_errors=True)


# --- 7-DAY AUTO CLEANUP function---
def clean_old_projects():
    """Deletes any project folder older than 7 days on system boot."""
    if not os.path.exists(PROJECTS_FOLDER):
        return
        
    current_time = time.time()
    seven_days_in_seconds = 7 * 24 * 60 * 60  # 7 days in seconds
    
    for folder_name in os.listdir(PROJECTS_FOLDER):
        folder_path = os.path.join(PROJECTS_FOLDER, folder_name)
        
        if os.path.isdir(folder_path):
            creation_time = os.path.getctime(folder_path)
            if (current_time - creation_time) > 24 * 60 * 60 and is_abandoned_project(folder_path):
                force_delete_project(folder_path)
                print(f"Cleaned up abandoned project: {folder_name}")
            elif (current_time - creation_time) > seven_days_in_seconds:
                try:
                    shutil.rmtree(folder_path, ignore_errors=True)
                    print(f"Cleaned up old project: {folder_name}")
                except Exception as e:
                    print(f"Could not delete {folder_name}: {e}")
        elif folder_name.lower().endswith(('.xlsx', '.xls')) and not folder_name.startswith('~$'):
            if (current_time - os.path.getmtime(folder_path)) > 24 * 60 * 60:
                force_delete_upload(folder_path)
                print(f"Cleaned up abandoned upload: {folder_name}")

clean_old_projects()  # Retry cleanup if deletion fails
app = Flask(__name__, 
            template_folder=os.path.join(BASE_DIR, 'templates'),
            static_folder=os.path.join(BASE_DIR, 'static'))
app.config['UPLOAD_FOLDER'] = PROJECTS_FOLDER
app.config['TEMP_FOLDER'] = temp_dir
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['TEMP_FOLDER'], exist_ok=True)

# Project names go into the folder name and every file name: kept short enough that every file of a project still
# opens in Excel (whose limit is about 218 characters for a file's full path)
PROJECT_NAME_MAX = 50


def _project_name(raw, default):
    """The project name typed, cut to PROJECT_NAME_MAX characters (the box doesn't take more either)."""
    return (str(raw or '').strip() or default)[:PROJECT_NAME_MAX].strip()


# The courier CSVs Generate can write, chosen just above the project name ('courier_csv'); Open360 by default
CSV_FORMATS = {'open360': 'Open360', 'openfreight': 'OpenFreight'}


# --- PAGES THAT SURVIVE A REFRESH ---
# A page shown after a form is sent (an upload, a preview, a result) is kept in data/page_views and the browser is
# sent to a plain address for it (?view=<id>), so refreshing shows it again instead of sending the form again.
# Leaving a page with an upload not yet used asks for it to be deleted (discard-on-leave), but only after
# DISCARD_GRACE seconds: a refresh comes straight back to the page, which keeps the upload.
VIEW_FOLDER = os.path.join(BASE_DIR, 'data', 'page_views')
VIEW_DAYS = 2
DISCARD_GRACE = 120
PENDING_DISCARD = {}  # path -> threading.Timer that deletes it


def _keep(*paths):
    """A page is using these uploads again: don't delete them."""
    for path in paths:
        timer = PENDING_DISCARD.pop(os.path.normcase(os.path.abspath(path)), None) if path else None
        if timer:
            timer.cancel()


def _discard_later(path, delete):
    """delete() after DISCARD_GRACE seconds, unless a page uses the upload again before then (_keep)."""
    key = os.path.normcase(os.path.abspath(path))
    _keep(path)
    def run():
        if PENDING_DISCARD.pop(key, None) is not None:
            delete()
    timer = threading.Timer(DISCARD_GRACE, run)
    timer.daemon = True
    PENDING_DISCARD[key] = timer
    timer.start()


def _keep_from(context):
    """The uploads a page shows (Distribution Mapper / Packing Labels file, Label Shuffler project) stay."""
    name = os.path.basename(str(context.get('filename') or ''))
    project = os.path.basename(str(context.get('project_name') or context.get('duplicate_project_name') or ''))
    _keep(os.path.join(PROJECTS_FOLDER, name) if name else None, os.path.join(PROJECTS_FOLDER, project) if project else None)


def _show(template, _at=None, **context):
    """Shows a page. After a form post, the page is kept and the browser is sent to it with a GET (?view=<id> on
    _at, the page's own address by default), so a refresh shows the same page and the form isn't sent again."""
    _keep_from(context)
    if request.method != 'POST':
        return render_template(template, **context)
    os.makedirs(VIEW_FOLDER, exist_ok=True)
    token = secrets.token_urlsafe(12)
    try:
        with open(os.path.join(VIEW_FOLDER, f"{token}.json"), 'w', encoding='utf-8') as f:
            json.dump({'template': template, 'context': context}, f, ensure_ascii=False, default=str)
    except (OSError, TypeError, ValueError) as e:
        log_pack.warning("Couldn't keep the page for a refresh (%s); showing it as is", e)
        return render_template(template, **context)
    # Old kept pages go after VIEW_DAYS days
    cutoff = time.time() - VIEW_DAYS * 86400
    for f in os.listdir(VIEW_FOLDER):
        try:
            if os.path.getmtime(os.path.join(VIEW_FOLDER, f)) < cutoff:
                os.remove(os.path.join(VIEW_FOLDER, f))
        except OSError:
            pass
    return redirect(f"{_at or request.path}?view={token}", code=303)


def _kept_view():
    """The kept page for ?view=<id> (see _show), or None."""
    token = request.args.get('view', '')
    if request.method != 'GET' or not re.fullmatch(r'[\w-]{8,40}', token):
        return None
    data = _read_json(os.path.join(VIEW_FOLDER, f"{token}.json"), None)
    if not data:
        return None
    _keep_from(data['context'])
    return render_template(data['template'], **data['context'])


# --- PART 1: MATRIX ENGINE ---
@app.route('/matrix', methods=['GET', 'POST'])
def matrix():
    kept = _kept_view()
    if kept:
        return kept
    tabs = None
    filename = None
    
    if request.method == 'POST':
        if 'file' not in request.files:
            return "No file part"
            
        file = request.files['file']
        
        if file.filename != '':
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            try:
                close_if_open_elsewhere(filepath)
            except Exception:
                pass
            file.save(filepath)

            filepath = convert_legacy_excel_to_xlsx(filepath)
            filename = os.path.basename(filepath)

            tabs = scan_excel_tabs(filepath)
            
    return _show('matrix.html', tabs=tabs, filename=filename)

@app.route('/preview', methods=['POST'])
def preview():
    filename = request.form.get('filename')
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    
    all_tabs = request.form.getlist('all_tabs')
    selected_tabs = request.form.getlist('selected_tabs')
    
    previews = []
    user_inputs = {}
    blueprints = {}
    
    for tab in all_tabs:
        safe_tab = tab.replace(" ", "_")
        
        start_cell = request.form.get(f"start_{tab}") or request.form.get(f"start_{safe_tab}") or "B8"
        job_id_cell = request.form.get(f"job_{tab}") or request.form.get(f"job_{safe_tab}") or "E1"
        store_col = request.form.get(f"store_{tab}") or request.form.get(f"store_{safe_tab}") or "A"
        
        user_inputs[tab] = {
            "start": start_cell,
            "job": job_id_cell,
            "store": store_col,
            "selected": tab in selected_tabs
        }
        
        if tab in selected_tabs:
            try:
                blueprint = generate_tab_map(filepath, tab, start_cell, job_id_cell, store_col)
                # Extract the raw backend data into our dictionary, removing it from the frontend view
                blueprints[tab] = blueprint.pop("backend_data")
                previews.append(blueprint)
            except Exception as e:
                previews.append({
                    "sheet_name": tab,
                    "error": f"Failed to map. Check your coordinates! ({str(e)})"
                })
                
    blueprints_json = json.dumps(blueprints)
            
    return _show('matrix.html', _at='/matrix', tabs=all_tabs, previews=previews, filename=filename, user_inputs=user_inputs, blueprints_json=blueprints_json)

@app.route('/generate', methods=['POST'])
def generate():
    filename = request.form.get('filename')
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    
    all_tabs = request.form.getlist('all_tabs')
    selected_tabs = request.form.getlist('selected_tabs')
    
    blueprints_json = request.form.get('blueprints_json')
    blueprints = json.loads(blueprints_json) if blueprints_json else {}
    
    user_inputs = {}
    for tab in all_tabs:
        safe_tab = tab.replace(" ", "_")
        user_inputs[tab] = {
            "start": request.form.get(f"start_{tab}") or request.form.get(f"start_{safe_tab}"),
            "job": request.form.get(f"job_{tab}") or request.form.get(f"job_{safe_tab}"),
            "store": request.form.get(f"store_{tab}") or request.form.get(f"store_{safe_tab}"),
            "selected_packs": request.form.getlist(f"packs_{tab}") or request.form.getlist(f"packs_{safe_tab}")
        }
        # "Sent to an installer": every pack in the tab gets signature codes, so every divider
        # sheet can carry its job-number barcodes (the page locks those pack checkboxes on).
        if (request.form.get(f"installer_{tab}") or request.form.get(f"installer_{safe_tab}")) == "true" and tab in blueprints:
            user_inputs[tab]["installer"] = True
            user_inputs[tab]["selected_packs"] = [p["name"] for p in blueprints[tab]["pack_ranges"]]

   # 1. Grab the user's custom project name from the form
    raw_project_name = _project_name(request.form.get('project_name'), 'Untitled_Project')
    safe_project_name = clean_file_name(raw_project_name)
    
    # 2. Extract the Job ID from the first selected tab's blueprint
    first_tab = selected_tabs[0] if selected_tabs else None
    job_id = "UNKNOWN"
    if first_tab and first_tab in blueprints:
        job_id = clean_file_name(blueprints[first_tab].get("raw_job_id", "UNKNOWN"))
        
    # 3. Create a clean, readable timestamp (YYMMDD_HHMM)
    time_stamp = datetime.now().strftime("%y%m%d_%H%M")
    
    # 4. Build the final folder name: e.g., CampaignName_Job-12345_240725_1430
    final_folder_name = f"{safe_project_name}_Job-{job_id}_{time_stamp}"
    
    project_dir = os.path.join(app.config['UPLOAD_FOLDER'], final_folder_name)
    os.makedirs(project_dir, exist_ok=True)
    
    # 5. MOVE THE ORIGINAL FILE INTO THE PROJECT FOLDER
    new_filepath = os.path.join(project_dir, filename)
    if os.path.exists(filepath):
        # If the user still has this file open (e.g. via the "Open Excel File" button used to
        # fix duplicate store names), Excel's lock would make this move crash with a
        # PermissionError. Force-close that copy first, discarding any unsaved edits - by this
        # point the user has already saved what they meant to keep and clicked Generate.
        close_if_open_elsewhere(filepath)
        shutil.move(filepath, new_filepath)
    
    # 6. Pass the NEW filepath and project_dir to the engine
    file1_name, file2_name, file3_name, shared_job_warnings = generate_all_outputs(
        new_filepath, filename, selected_tabs, user_inputs, blueprints, project_dir
    )
    raw_files = [file1_name, file2_name, file3_name]

    # Filter the list to ONLY keep actual file names (removes None/Empty strings)
    files_to_download = [f for f in raw_files if f]

    # 3. Pass the clean list to the template
    return _show('matrix.html', _at='/matrix',
                           generation_complete=True,
                           project_folder=final_folder_name,
                           generated_files=files_to_download,
                           shared_job_warnings=shared_job_warnings)

def numbered_name(folder, kind, *parts, ext=".pdf"):
    """'Complete Labels 2 - J477161 - Lux Test 30 - 261005_2016.pdf': the kind of file, numbered after the ones of that
    kind already in the folder (so a project can hold several), then job number, project and timestamp."""
    taken = [int(m.group(1)) for f in (os.listdir(folder) if os.path.isdir(folder) else [])
             for m in [re.match(re.escape(kind) + r' (\d+) - ', f)] if m]
    return " - ".join([f"{kind} {max(taken, default=0) + 1}"] + [p for p in parts if p]) + ext


def newest_first(folder, names):
    return sorted(names, key=lambda f: os.path.getmtime(os.path.join(folder, f)), reverse=True)


def _generation_report(job, project, shipment_reference, sheet_warnings, courier_warnings,
                       all_cartons, cartons, consignments, by_number, service, sender, csv_formats=None):
    """The job report on page 1 of the packing labels: job number, project, and what needs a look while preparing
    labels (draw_report_pages). Read before Generate saves the new addresses to the address book."""
    def receiver(c):
        d = c['destination']
        return d['receiver'] + (f" (Attn {d['contact']})" if d.get('contact') else '')
    left_out = [x for x in all_cartons if x not in cartons]
    try:
        new_addresses = [c for c in consignments if not ADDRESS_BOOK.at_address(c['destination'])]
    except Exception:
        new_addresses = []
    no_size = [x for x in cartons if x.get('spec_kind') in ('unknown', 'incomplete')]
    installers = [x for x in all_cartons if x.get('install')]
    by_installer = {}
    for x in installers:  # one line per installer: the boxes to pack and label for them
        by_installer.setdefault(receiver(by_number[x['consignment']]), []).append(x['item_reference'])
    own = ("No address for ", "No size saved for packing spec")  # shown as their own sections below
    sections = [
        {'title': "Not in the courier CSV - no complete address, no courier label will come", 'level': 'bad', 'kind': 'left_out',
         'items': [f"{x['item_reference']} ({x['packing_spec']}) - {receiver(by_number[x['consignment']])}" for x in left_out]},
        {'title': "No size saved for the packing spec - dimensions blank in the CSV", 'level': 'warn',
         'items': [f"{x['item_reference']}: {x['packing_spec']}" for x in no_size]},
        {'title': "New to the address book - check before dispatch", 'level': 'warn',
         'items': [f"#{c['number']} {receiver(c)} - {one_line(c['destination'])}" for c in new_addresses]},
        {'title': "Courier checks", 'level': 'warn', 'items': [w for w in courier_warnings if not w.startswith(own)]},
        {'title': "Spreadsheet checks", 'level': 'info',
         'items': [", ".join(p for p in (f"Row {w['rows']}" if w.get('rows') else '', f"col {w['col']}" if w.get('col') else '') if p)
                   + (": " if w.get('rows') or w.get('col') else '') + w.get('text', '')
                   + (f" - {w['detail']}" if w.get('detail') else '') for w in sheet_warnings]},
        {'title': f"Going to installers - {len(installers)} pack{'s' if len(installers) != 1 else ''} labelled for the installer",
         'level': 'info', 'items': [f"{who} x{len(refs)}" for who, refs in by_installer.items()]},
    ]
    facts = [("Packing labels", len(all_cartons)), ("Courier labels to come", len(cartons)), ("Consignments", len(consignments)),
             ("Installer packs", len(installers)),
             # Every service in use, most used first: a consignment can have its own ("IPECX ×2, BORDERP ×1")
             ("Service", ", ".join(f"{code} ×{n}" for code, n in Counter(c['service_code_used'] for c in consignments).most_common())
                         if consignments else (service or '-'))]
    if csv_formats:
        facts.append(("Courier CSV", " + ".join(CSV_FORMATS[f] for f in csv_formats)))
    if left_out:
        facts.insert(2, ("Not in courier CSV", len(left_out), 'bad'))
    return {'title': "Packing Labels Only", 'job': job, 'project': project,
            'when': "Generated " + datetime.now().strftime("%d/%m/%Y %H:%M"), 'facts': facts, 'sections': sections}


def _stitch_project(project_name):
    """(project folder, label map file name, label map) for a project that has a label map, else None."""
    folder = os.path.join(app.config['UPLOAD_FOLDER'], os.path.basename(project_name))
    maps = sorted(f for f in os.listdir(folder) if f.endswith('.labelmap.json')) if os.path.isdir(folder) else []
    if not maps:
        return None
    with open(os.path.join(folder, maps[0]), encoding='utf-8') as f:
        return folder, maps[0], json.load(f)


SESSION_FILE, MATCHES_FILE = "stitch_session.json", "stitch_matches.json"
STITCH_JOBS = {}  # project folder -> {'state': 'running' | 'done' | 'error', ...} for the matching page's Save
STITCH_DONE = {}  # project folder -> threading.Event, set when its background stitch finishes (status waits on it)


def _read_json(path, default):
    try:
        with open(path, encoding='utf-8') as f:
            return json.load(f)
    except (OSError, ValueError):
        return default


def _write_json(path, data):
    tmp = f"{path}.tmp"
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=1, ensure_ascii=False)
    os.replace(tmp, path)


def _courier_pdfs(folder, names):
    """[(name, bytes)] for the courier label PDFs of a project that are still in its folder."""
    pdfs = []
    for name in names:
        path = os.path.join(folder, os.path.basename(name))
        if os.path.isfile(path):
            with open(path, 'rb') as f:
                pdfs.append((os.path.basename(name), f.read()))
    return pdfs


def _plan(folder, label_map, pdfs, add_unmatched=None):
    """Matches every courier label to its packing label from the labels' text and the matches saved by hand (fast:
    nothing is drawn) and keeps the result for the matching page. Returns (plan, session)."""
    started = time.perf_counter()
    manual = _read_json(os.path.join(folder, MATCHES_FILE), {})
    log_stitch.info("Match %s: %s, %d saved hand match(es)", os.path.basename(folder),
                    ", ".join(f"{n} ({len(d) // 1024} KB)" for n, d in pdfs), len(manual))
    plan = plan_stitch(folder, label_map, pdfs, manual)
    session = match_session(label_map, plan)
    old = _read_json(os.path.join(folder, SESSION_FILE), {})
    session['added_unmatched'] = old.get('added_unmatched', True) if add_unmatched is None else add_unmatched
    session['output'] = old.get('output') if old.get('output') and os.path.isfile(os.path.join(folder, old['output'])) else None
    _write_json(os.path.join(folder, SESSION_FILE), session)
    log_stitch.info("  matched in %.2fs: %d placed, %d courier label(s) unmatched, %d packing label(s) waiting%s",
                    time.perf_counter() - started, len(plan['placed']), len(plan['unmatched']), len(plan['missing']),
                    " - complete" if plan['complete'] else '')
    return plan, session


def _remove_old_labels(folder, keep=None):
    """Deletes the earlier Complete Labels PDFs of a project (a new one replaces them), so the folder holds one.
    Returns the ones that couldn't be deleted (open in a PDF viewer)."""
    locked = []
    for f in os.listdir(folder):
        if f.startswith("Complete Labels") and f.endswith(".pdf") and f != keep:
            try:
                os.remove(os.path.join(folder, f))
                log_stitch.info("  removed the earlier %s", f)
            except OSError as e:
                log_stitch.warning("  couldn't remove the earlier %s (%s) - open in a PDF viewer?", f, e)
                locked.append(f)
    return locked


def _finalise(folder, label_map, pdfs, add_unmatched=True):
    """Stitches once: the plan (codes + saved hand matches) drawn into a new Complete Labels PDF, which replaces the
    earlier ones. Returns (output name, result, earlier files that were open elsewhere and stayed)."""
    plan, session = _plan(folder, label_map, pdfs, add_unmatched)
    project_label = (label_map.get('project') or {}).get('name') or re.sub(r'_\d{6}_\d{4}$', '', os.path.basename(folder))
    output_name = numbered_name(folder, "Complete Labels", label_map.get('consignment_reference'), project_label,
                                datetime.now().strftime("%y%m%d_%H%M"))
    started = time.perf_counter()
    try:
        result = render_stitch(folder, label_map, pdfs, plan, os.path.join(folder, output_name), add_unmatched)
    except Exception:
        log_stitch.exception("Stitch FAILED after %.2fs", time.perf_counter() - started)
        raise
    locked = _remove_old_labels(folder, keep=output_name)
    session.update(output=output_name, added_unmatched=add_unmatched, finalised=True)
    _write_json(os.path.join(folder, SESSION_FILE), session)
    log_stitch.info("Stitched in %.2fs -> %s: %d placed, %d courier label(s) unmatched, %d packing label(s) without one, %d duplicate(s)",
                    time.perf_counter() - started, output_name, len(result['placed']), len(result['unmatched']),
                    len(result['without_label']), len(result['duplicates']))
    # The first few in the terminal; all of them in logs/warehouse.log
    for n, u in enumerate(result['unmatched']):
        log_stitch.log(logging.INFO if n < 8 else logging.DEBUG, "  unmatched %s p%s: %s", u['file'], u['page'], u['reason'])
    if len(result['unmatched']) > 8:
        log_stitch.info("  ... and %d more unmatched (all in logs/warehouse.log)", len(result['unmatched']) - 8)
    for n, r in enumerate(result['without_label']):
        log_stitch.log(logging.INFO if n < 8 else logging.DEBUG, "  without a courier label: %s", r)
    if len(result['without_label']) > 8:
        log_stitch.info("  ... and %d more without a courier label (all in logs/warehouse.log)", len(result['without_label']) - 8)
    for d in result['duplicates']:
        log_stitch.warning("  duplicate %s p%s: %s (first used %s)", d['file'], d['page'], d['item_reference'], d['first'])
    return output_name, {**result, 'session': session, 'complete': plan['complete']}, locked


def _start_finalise(folder, label_map, pdfs, add_unmatched, **extra):
    """_finalise in the background; the page asks /status (which waits for it)."""
    STITCH_JOBS[folder] = {'state': 'running', **extra}
    STITCH_DONE[folder] = finished = threading.Event()

    def work():
        try:
            output_name, result, locked = _finalise(folder, label_map, pdfs, add_unmatched)
            STITCH_JOBS[folder] = {'state': 'done', 'output': output_name, 'placed': len(result['placed']),
                                   'added_unmatched': add_unmatched, 'locked': locked, **extra,
                                   'unmatched': len(result['unmatched']),  # really matched nothing (hand matches aren't)
                                   'waiting': sum(1 for w in result['without_label'] if not w.endswith("(not in the courier CSV)"))}
        except Exception as e:  # shown on the page; the traceback is in the log (_finalise)
            STITCH_JOBS[folder] = {'state': 'error', 'error': str(e)}
        finally:
            finished.set()
    threading.Thread(target=work, daemon=True, name="stitch").start()


@app.route('/stitch/<project_name>', methods=['GET', 'POST'])
def stitch_page(project_name):
    """Stitch Labels: the courier portal's label PDFs onto the project's packing labels (label_stitcher.py).
    Every courier label is read and matched first (text only). All matched: the Complete Labels PDF is made straight
    away and the page just offers it. Anything left over: straight on to the matching page."""
    kept = _kept_view()
    if kept:
        return kept
    found = _stitch_project(project_name)
    if not found:
        return _show('stitch.html', project=project_name, error="This project has no label map, so there's nothing to stitch to.")
    folder, map_name, label_map = found
    stitched = newest_first(folder, [f for f in os.listdir(folder) if f.startswith("Complete Labels") and f.endswith(".pdf")])
    own = {(label_map.get('files') or {}).get('packing_labels_pdf')} | set(stitched)
    candidates = newest_first(folder, [f for f in os.listdir(folder) if f.lower().endswith('.pdf') and f not in own])
    slots = packing_slots(label_map)
    session = _read_json(os.path.join(folder, SESSION_FILE), None)
    info = {'packs': sum(1 for sl in slots if not sl['not_in_csv']),
            'boxed': sum(1 for sl in slots if sl['boxes'] > 1 and sl['box'] == 1),
            'with_code': sum(1 for pk in (p for ps in label_map.get('stores', {}).values() for p in ps) if pk.get('open360_code')),
            'left_out': len(label_map.get('not_in_courier_csv') or []),
            'latest': stitched[0] if stitched else None, 'stitched': stitched, 'candidates': candidates,
            'matching': bool(session and not session.get('complete') and (session.get('unmatched') or session.get('waiting'))
                             and not session.get('finalised'))}
    if request.method == 'GET':
        return _show('stitch.html', project=os.path.basename(folder), info=info)

    pdfs = []
    for upload in request.files.getlist('labels'):
        if upload and upload.filename:
            name = secure_filename(upload.filename) or 'courier-labels.pdf'
            if not name.lower().endswith('.pdf'):
                return _show('stitch.html', project=os.path.basename(folder), info=info,
                                       error=f"{upload.filename} isn't a PDF. Upload the label PDF from the courier portal.")
            if name in own:
                name = f"courier-{name}"
            data = upload.read()
            with open(os.path.join(folder, name), 'wb') as f:  # kept with the project
                f.write(data)
            pdfs.append((name, data))
    for name in request.form.getlist('existing'):
        name = os.path.basename(name)
        if name in candidates and name not in {n for n, _ in pdfs}:
            with open(os.path.join(folder, name), 'rb') as f:
                pdfs.append((name, f.read()))
    if not pdfs:
        log_stitch.warning("Stitch %s: no courier label PDF chosen", os.path.basename(folder))
        return _show('stitch.html', project=os.path.basename(folder), info=info,
                               error="Choose the courier label PDF (upload it, or tick one already in this project).")
    add_unmatched = request.form.get('unmatched', 'add') != 'leave'
    try:
        plan, session = _plan(folder, label_map, pdfs, add_unmatched)
        if not plan['complete']:
            # Something to match by hand: straight to the matching page; nothing is stitched until Finalise
            return redirect(url_for('stitch_match_page', project_name=os.path.basename(folder)))
        output_name, report, locked = _finalise(folder, label_map, pdfs, add_unmatched)
    except StitchError as e:
        log_stitch.warning("Stitch stopped: %s", e)
        return _show('stitch.html', project=os.path.basename(folder), info=info, error=str(e))
    info.update(latest=output_name, stitched=[output_name], matching=False,
                candidates=newest_first(folder, list(set(candidates) | {n for n, _ in pdfs})))
    return _show('stitch.html', project=os.path.basename(folder), info=info, report=report,
                           files=[n for n, _ in pdfs], locked=locked)


@app.route('/stitch/<project_name>/match')
def stitch_match_page(project_name):
    """Match by hand: unmatched courier labels on the left, packing labels still waiting on the right."""
    found = _stitch_project(project_name)
    if not found:
        return redirect(url_for('stitch_page', project_name=project_name))
    folder, _, label_map = found
    session = _read_json(os.path.join(folder, SESSION_FILE), None)
    if not session:
        return redirect(url_for('stitch_page', project_name=project_name))
    return render_template('stitch_match.html', project=os.path.basename(folder), session=session,
                           session_json=json.dumps(session), job=label_map.get('consignment_reference'),
                           project_label=(label_map.get('project') or {}).get('name') or os.path.basename(folder))


_LABEL_IMAGES = {}


@app.route('/stitch/<project_name>/label/<int:index>.png')
def stitch_label_image(project_name, index):
    """An unmatched courier label as a picture for the matching page: upright, margins cropped (render_label)."""
    found = _stitch_project(project_name)
    if not found:
        return "Not found", 404
    folder = found[0]
    session = _read_json(os.path.join(folder, SESSION_FILE), {})
    rows = session.get('unmatched') or []
    if not 0 <= index < len(rows):
        return "Not found", 404
    row, dpi = rows[index], 200 if request.args.get('size') == 'zoom' else 70
    path = os.path.join(folder, os.path.basename(row['file']))
    cache_key = (path, row['page'], dpi, os.path.getmtime(path) if os.path.exists(path) else 0)
    if cache_key not in _LABEL_IMAGES:
        try:
            doc = fitz.open(path)
            img = render_label(doc[row['page'] - 1], dpi=dpi)
            doc.close()
        except Exception:
            log_stitch.exception("Could not draw courier label %s p%s for the matching page", row['file'], row['page'])
            return "Not found", 404
        buf = io.BytesIO()
        img.save(buf, format="PNG", optimize=True)
        if len(_LABEL_IMAGES) > 400:
            _LABEL_IMAGES.clear()
        _LABEL_IMAGES[cache_key] = buf.getvalue()
    return send_file(io.BytesIO(_LABEL_IMAGES[cache_key]), mimetype='image/png', max_age=300)


def _valid_matches(folder, matches):
    """The hand-made matches ({courier page key: slot}) that still fit what's on the matching page."""
    session = _read_json(os.path.join(folder, SESSION_FILE), {})
    known_keys = {u['key'] for u in session.get('unmatched') or []}
    known_slots = {w['item_reference'] for w in session.get('waiting') or []}
    return {k: v for k, v in (matches or {}).items() if k in known_keys and v in known_slots}, session, known_keys


def _save_matches(folder, matches, known_keys):
    """The page's pairs replace what was saved for its courier labels (a pair undone there is dropped too)."""
    saved = {k: v for k, v in _read_json(os.path.join(folder, MATCHES_FILE), {}).items() if k not in known_keys}
    saved.update(matches)
    _write_json(os.path.join(folder, MATCHES_FILE), saved)
    for k, v in matches.items():
        log_stitch.debug("  pair %s -> %s", k, v)


@app.route('/stitch/<project_name>/match', methods=['POST'])
def stitch_match_save(project_name):
    """Save changes: keeps the hand-made matches ({courier page key: slot}) for this project. Nothing is stitched
    (that's Finalise); the matching is just worked out again, which takes a moment."""
    found = _stitch_project(project_name)
    if not found:
        return {'error': "Project not found."}, 404
    folder, _, label_map = found
    if (STITCH_JOBS.get(folder) or {}).get('state') == 'running':
        return {'error': "Still stitching - wait for it to finish."}, 409
    sent = (request.get_json(silent=True) or {}).get('matches') or {}
    matches, session, known_keys = _valid_matches(folder, sent)
    log_stitch.info("Matching page save %s: %d pair(s) sent, %d valid (%d ignored: label or pack no longer waiting)",
                    os.path.basename(folder), len(sent), len(matches), len(sent) - len(matches))
    if not known_keys:
        return {'error': "Nothing to match on this page any more: open it again."}, 409
    _save_matches(folder, matches, known_keys)
    try:
        plan, _ = _plan(folder, label_map, _courier_pdfs(folder, session.get('files') or []))
    except StitchError as e:
        return {'error': str(e)}, 400
    return {'saved': len(matches), 'complete': plan['complete'], 'placed': len(plan['placed']),
            'unmatched': len(plan['unmatched']), 'waiting': len(plan['missing'])}


@app.route('/stitch/<project_name>/finalise', methods=['POST'])
def stitch_finalise(project_name):
    """Finalise: saves the matches sent with it, then stitches once, in the background, into a Complete Labels PDF
    that replaces the earlier one."""
    found = _stitch_project(project_name)
    if not found:
        return {'error': "Project not found."}, 404
    folder, _, label_map = found
    if (STITCH_JOBS.get(folder) or {}).get('state') == 'running':
        return {'error': "Still stitching - wait for it to finish."}, 409
    body = request.get_json(silent=True) or {}
    matches, session, known_keys = _valid_matches(folder, body.get('matches'))
    if known_keys and 'matches' in body:
        _save_matches(folder, matches, known_keys)
    add_unmatched = bool(body.get('add_unmatched', True))
    pdfs = _courier_pdfs(folder, session.get('files') or [])
    if not pdfs:
        return {'error': "The courier label PDFs aren't in the project folder any more: add them again."}, 400
    log_stitch.info("Finalise %s: %d new match(es), unmatched labels %s", os.path.basename(folder), len(matches),
                    "added 4-up" if add_unmatched else "left out")
    _start_finalise(folder, label_map, pdfs, add_unmatched, matched=len(matches))
    return {'state': 'running', 'saved': len(matches)}, 202


@app.route('/stitch/<project_name>/add', methods=['POST'])
def stitch_add_pdfs(project_name):
    """More courier label PDFs from the matching page: kept in the project, and the matching is worked out again with
    them and the ones already used. The earlier Complete Labels PDF is removed (it's out of date). If everything
    now matches, it's stitched straight away (in the background)."""
    found = _stitch_project(project_name)
    if not found:
        return {'error': "Project not found."}, 404
    folder, _, label_map = found
    if (STITCH_JOBS.get(folder) or {}).get('state') == 'running':
        return {'error': "Still stitching - wait for it to finish."}, 409
    session = _read_json(os.path.join(folder, SESSION_FILE), {})
    own = {(label_map.get('files') or {}).get('packing_labels_pdf')}
    names = [os.path.basename(n) for n in session.get('files') or []]
    added = []
    for upload in request.files.getlist('labels'):
        if not (upload and upload.filename):
            continue
        name = secure_filename(upload.filename) or 'courier-labels.pdf'
        if not name.lower().endswith('.pdf'):
            return {'error': f"{upload.filename} isn't a PDF."}, 400
        if name in own or name.startswith("Complete Labels"):
            name = f"courier-{name}"
        upload.save(os.path.join(folder, name))
        if name not in names:
            names.append(name)
        added.append(name)
    if not added:
        return {'error': "Choose one or more courier label PDFs."}, 400
    log_stitch.info("Matching page: courier label PDF(s) added to %s: %s", os.path.basename(folder), added)
    locked = _remove_old_labels(folder)
    pdfs = _courier_pdfs(folder, names)
    add_unmatched = session.get('added_unmatched', True)
    try:
        plan, _ = _plan(folder, label_map, pdfs, add_unmatched)
    except StitchError as e:
        return {'error': str(e)}, 400
    if plan['complete']:
        _start_finalise(folder, label_map, pdfs, add_unmatched, added_files=added)
        return {'state': 'running', 'added': added, 'complete': True, 'locked': locked}, 202
    return {'state': 'matching', 'added': added, 'complete': False, 'locked': locked,
            'unmatched': len(plan['unmatched']), 'waiting': len(plan['missing'])}


@app.route('/stitch/<project_name>/status')
def stitch_match_status(project_name):
    """The background stitch's state. With ?wait=1 it waits (up to 25 s) for the stitch to finish before answering,
    so the matching page asks once in a while instead of every second."""
    found = _stitch_project(project_name)
    if not found:
        return {'state': 'none'}
    job = STITCH_JOBS.get(found[0]) or {'state': 'none'}
    if request.args.get('wait') and job.get('state') == 'running' and found[0] in STITCH_DONE:
        STITCH_DONE[found[0]].wait(timeout=25)
        job = STITCH_JOBS.get(found[0]) or {'state': 'none'}
    return job


@app.route('/download/<folder_name>/<filename>')
def download_file(folder_name, filename):
    """Secure endpoint allowing users to pull individual output sheets from their specific project folder."""
    safe_folder = os.path.basename(folder_name)
    safe_filename = os.path.basename(filename)
    
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], safe_folder, safe_filename)
    
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
        
    return "File Not Found", 404


@app.route('/')
def dashboard():
    """Main dashboard displaying job history."""
    projects = []
    if os.path.exists(app.config['UPLOAD_FOLDER']):
        for folder_name in os.listdir(app.config['UPLOAD_FOLDER']):
            folder_path = os.path.join(app.config['UPLOAD_FOLDER'], folder_name)
            if os.path.isdir(folder_path):
                
                # Get files inside the project
                files = os.listdir(folder_path)
                
                # --- THE ZOMBIE SWEEPER ---
                if len(files) == 0:
                    try:
                        # If Windows has finally unlocked the folder, delete it permanently!
                        shutil.rmdir(folder_path, ignore_errors=True)
                    except Exception:
                        pass
                    continue # Always skip showing empty folders in the UI

                # --- FILTER AND GROUP FILES ---
                files = newest_first(folder_path, files)  # last modified first
                display_files = [f for f in files if not f.endswith('.json')]
                # Sort JSON files by creation time (newest first)
                json_files = sorted([f for f in files if f.endswith('.json') and not f.endswith('.labelmap.json')], key=lambda x: os.path.getctime(os.path.join(folder_path, x)), reverse=True )

                
                excel_files = [f for f in display_files if f.lower().endswith(('.xlsx', '.xls'))]
                pdf_files = [f for f in display_files if f.lower().endswith('.pdf')]
                csv_files = [f for f in display_files if f.lower().endswith('.csv')]
                
                # Get human-readable date
                timestamp = os.path.getctime(folder_path)
                date_str = datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M')

                # Extract the Job ID dynamically from File 2's name
                job_id = "N/A"
                for f in display_files:
                    if f.startswith("Packing Sheet_"):
                        job_id = f.replace("Packing Sheet_", "").rsplit(".", 1)[0]
                        break
                        
                display_name = folder_name.split('_Job-')[0] if '_Job-' in folder_name else folder_name
                has_label_map = any(f.endswith('.labelmap.json') for f in files)
                courier_only = bool(csv_files) and not pdf_files and not has_label_map

                # The features the project was made with, for its colours on the dashboard (the same as their tiles
                # on Home): Courier Import, Packing Labels, or Packing Sheets, plus the Label Shuffler when its labels
                # were shuffled there too (both shown, sharing the colour band equally)
                sheets = '_Job-' in folder_name or any(f.startswith('Packing Sheet_') for f in files)
                if courier_only:
                    kinds = ['import']
                elif has_label_map or any(f.startswith(('Packing Labels', 'Complete Labels')) for f in pdf_files):
                    kinds = ['labels']
                elif sheets:
                    kinds = ['sheets'] + (['shuffler'] if any('_Shuffled' in f for f in pdf_files) else [])
                else:
                    kinds = ['labels'] if pdf_files else ['sheets']
                
                projects.append({
                    "name": folder_name,          
                    "display_name": display_name, 
                    "date": date_str,
                    "timestamp": timestamp,
                    "job_id": job_id,
                    "excel_files": excel_files,  # Pass the grouped Excel files
                    "pdf_files": pdf_files,       # Pass the grouped PDFs
                    "csv_files": csv_files,
                    "has_label_map": has_label_map,
                    "courier_only": courier_only,  # Courier Import: the courier CSV only, nothing to stitch or sub-group
                    "kinds": kinds,
                    "json_files": json_files     # Pass the grouped JSON files
                })
                
    # Sort projects newest to oldest
    projects.sort(key=lambda x: x['timestamp'], reverse=True)
    
    return render_template('index.html', projects=projects)

@app.route('/delete/<folder_name>', methods=['POST'])
def delete_project(folder_name):
    """Aggressively deletes a specific project folder."""
    safe_folder = os.path.basename(folder_name)
    folder_path = os.path.join(app.config['UPLOAD_FOLDER'], safe_folder)
    
    if os.path.exists(folder_path) and os.path.isdir(folder_path):
        # 1. Forcefully delete all files inside and strip read-only locks
        for root, dirs, files in os.walk(folder_path, topdown=False):
            for name in files:
                file_path = os.path.join(root, name)
                try:
                    os.chmod(file_path, stat.S_IWRITE)
                    os.remove(file_path)
                except Exception:
                    pass
        try:
            shutil.rmtree(folder_path)
        except Exception as e1:
            print(f"[DEBUG] FAILED: shutil.rmtree error -> {e1}")
            print("[DEBUG] Falling back to Method B (os.rmdir)...")
            
    return redirect(url_for('dashboard'))
@app.route('/delete_file/<folder_name>/<filename>', methods=['POST'])
def delete_single_file(folder_name, filename):
    """Deletes a specific file inside a project."""
    safe_folder = os.path.basename(folder_name)
    safe_filename = os.path.basename(filename)
    
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], safe_folder, safe_filename)
    
    if os.path.exists(file_path):
        try:
            os.chmod(file_path, stat.S_IWRITE)
            os.remove(file_path)
        except Exception as e:
            print(f"[DEBUG] Could not delete file: {e}")
            
            
    return redirect(url_for('dashboard'))

@app.route('/open_local/<folder_name>/<filename>')
def open_local_file(folder_name, filename):
    """Commands Windows to open the file directly using its default application."""
    safe_folder = os.path.basename(folder_name)
    safe_filename = os.path.basename(filename)
    
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], safe_folder, safe_filename)
    
    if os.path.exists(file_path):
        try:
            # os.startfile is a built-in Windows command that opens a file natively
            os.startfile(file_path)
        except Exception as e:
            print(f"[DEBUG] Could not open file locally: {e}")

    return '', 204  # Prevents the browser from reloading the page

@app.route('/discard_upload', methods=['POST'])
def discard_upload():
    """Deletes a scanned-but-never-generated upload when the user leaves the page."""
    safe_filename = os.path.basename(request.form.get('filename', ''))
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], safe_filename)
    if safe_filename.lower().endswith(('.xlsx', '.xls')) and os.path.isfile(file_path):
        # Not at once: a refresh also leaves the page, and comes straight back to it (which keeps the upload)
        _discard_later(file_path, lambda: os.path.isfile(file_path) and force_delete_upload(file_path))
    return '', 204

@app.route('/discard_project', methods=['POST'])
def discard_project():
    """Deletes a Label Shuffler project folder the user left before generating anything."""
    safe_folder = os.path.basename(request.form.get('project', ''))
    folder_path = os.path.join(PROJECTS_FOLDER, safe_folder)
    if safe_folder not in ('', '.', '..') and os.path.isdir(folder_path) and is_abandoned_project(folder_path):
        _discard_later(folder_path, lambda: os.path.isdir(folder_path) and is_abandoned_project(folder_path)
                       and force_delete_project(folder_path))
    return '', 204

@app.route('/open_upload/<filename>')
def open_upload_file(filename):
    """Same as /open_local, but for a freshly-uploaded file that still sits directly in the
    uploads root (Distribution Mapper flow, before a project folder is created at /generate)."""
    safe_filename = os.path.basename(filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], safe_filename)

    if os.path.exists(file_path):
        try:
            os.startfile(file_path)
        except Exception as e:
            print(f"[DEBUG] Could not open file locally: {e}")

    return '', 204  # Prevents the browser from reloading the page

@app.route('/subgroup/<project_name>', methods=['GET', 'POST'])
def setup_subgroup(project_name):
    safe_project = os.path.basename(project_name)
    project_dir = os.path.join(app.config['UPLOAD_FOLDER'], safe_project)

    # ADD THIS: Get the specific JSON file chosen by the user (defaults to project_metadata.json)
    target_json = request.values.get('target_json')
    # 2. Dynamic Fallback: If target_json is missing, grab the first available .json file in the project folder
    if not target_json or target_json == 'default':
        json_files = [f for f in os.listdir(project_dir) if f.endswith('.json') and not f.endswith('.labelmap.json')]
        target_json = json_files[0] if json_files else f"{safe_project}.json"
        print(f"[DEBUG] No target_json specified. Defaulting to: {target_json}")
    metadata_path = os.path.join(project_dir, target_json)
    
    

    
    if not os.path.exists(metadata_path):
        return "Metadata not found for this project. Cannot create sub-groups.", 404
        
    with open(metadata_path, 'r') as f:
        metadata = json.load(f)
        
    if request.method == 'POST':
        # 1. Grab the list of selected tabs
        selected_tabs = request.form.getlist('selected_tabs[]')
        
        # 2. Build a clean dictionary of instructions for the engine
        subgroup_instructions = {}
        
        for tab in selected_tabs:
            item_row = int(request.form.get(f'item_row_{tab}'))
            selected_packs = request.form.getlist(f'target_pack_{tab}[]')
            
            tab_instructions = {
                "item_row": item_row,
                "packs": {}
            }
            
            for pack in selected_packs:
                # Grab the paired arrays of start and end numbers
                starts = request.form.getlist(f'start_item_{tab}_{pack}[]')
                ends = request.form.getlist(f'end_item_{tab}_{pack}[]')
                
                # Zip them together into neat pairs (e.g., [[1, 5], [6, 10]])
                ranges = [[int(s), int(e)] for s, e in zip(starts, ends) if s and e]
                
                if ranges:
                    tab_instructions["packs"][pack] = ranges
                    
            if tab_instructions["packs"]:
                subgroup_instructions[tab] = tab_instructions
                
        # --- DEBUG PRINT ---
        print("\n--- SUBGROUP INSTRUCTIONS ---")
        pprint.pprint(subgroup_instructions)
        print("-----------------------------\n")

        # 3. Trigger the Engine (We will build this function next!)
        try:
            execute_subgroups(project_dir, metadata, subgroup_instructions)
        except SubgroupValidationError as e:
            return render_template('sub-group.html', project_name=project_name, metadata=metadata,
                                   metadata_json=json.dumps(metadata), target_json=target_json,
                                   error=str(e))

        # 4. Redirect back to dashboard upon completion
        return redirect(url_for('dashboard'))
        
    return render_template('sub-group.html', project_name=project_name, metadata=metadata, metadata_json=json.dumps(metadata), target_json=target_json)

@app.route('/pdf', methods=['GET', 'POST'])
def pdf_engine():
    kept = _kept_view()
    if kept:
        return kept
    os.makedirs(PROJECTS_FOLDER, exist_ok=True)
    
    if request.method == 'POST':
        step = request.form.get('step')
        
       # STEP 1: Process Excel and Project Name
        if step == '1':
            excel_file = request.files.get('excel_file')
            existing_project = request.form.get('existing_project')
            new_project = _project_name(request.form.get('new_project'), '') or None
            time_stamp = datetime.now().strftime("%y%m%d_%H%M")
            safe_project_name = clean_file_name(new_project)
            final_folder_name = f"{safe_project_name}_{time_stamp}"

            
            project_name = final_folder_name.strip() if new_project and new_project.strip() else existing_project
            
            if not project_name:
                return "Please select or enter a Project Name.", 400
            
            # --- THE FIX: Point directly to the final Project Folder ---
            project_folder = os.path.join(PROJECTS_FOLDER, os.path.basename(project_name))
            os.makedirs(project_folder, exist_ok=True)
            
            excel_path = None
            
            # Scenario A: User uploaded a new file (Save it directly to the project folder)
            if excel_file and excel_file.filename != '':
                filename = secure_filename(excel_file.filename)
                excel_path = os.path.join(project_folder, filename)
                excel_file.save(excel_path)
            
            # Scenario B: Existing project selected (Just point to the file already in the folder!)
            elif existing_project:
                # If we got here from the duplicate modal's "Recheck File" button, it tells us
                # the exact filename to reload instead of guessing by naming convention.
                resume_filename = request.form.get('resume_filename')
                if resume_filename:
                    candidate_path = os.path.join(project_folder, secure_filename(resume_filename))
                    if os.path.exists(candidate_path):
                        excel_path = candidate_path

                if not excel_path and os.path.exists(project_folder):
                    for f in os.listdir(project_folder):
                        if f.lower().startswith("signature links") and (f.lower().endswith(".xlsx") or f.lower().endswith(".xls")):
                            excel_path = os.path.join(project_folder, f)
                            break
                            
            # Failsafe if no file was uploaded AND no file was found
            if not excel_path or not os.path.exists(excel_path):
                return "No Signature links file found or uploaded. Please try again.", 400
            
            tabs_data = {}
            duplicate_errors = []
            installer_tabs = []
            try:
                with pd.ExcelFile(excel_path) as xls:
                    installer_tabs = sorted({tab for tab, _ in read_divider_barcodes(xls)})
                    for sheet in xls.sheet_names:
                        if sheet == DIVIDER_BARCODE_SHEET:
                            continue  # barcode data for the dividers, not a tab of store codes
                        df = pd.read_excel(xls, sheet_name=sheet)
                        packs = df.columns[1:].tolist()
                        tabs_data[sheet] = packs
                        # --- CHECK FOR DUPLICATE STORE NAMES (COLUMN 0) ---
                        if not df.empty:
                            # Store names live in column 0
                            store_col = df.iloc[:, 0].dropna().astype(str).str.strip()
                            store_col = store_col[store_col.str.lower() != 'nan']  # Remove string 'nan'
                            store_col = store_col.apply(clean_store_name)

                            # Identify duplicates
                            dupes = store_col[store_col.duplicated()].unique().tolist()
                            if dupes:
                                duplicate_errors.append(f"Tab '{sheet}': {', '.join(dupes)}")
            except Exception as e:
                return f"Error reading Excel file: {e}", 500
            # --- HALT PROCESS IF DUPLICATES EXIST ---
            if duplicate_errors:
                # NOTE: We deliberately do NOT delete project_folder here. Deleting it used to
                # wipe out prior work whenever an existing project's replacement Excel still had
                # duplicates. The folder and the uploaded file are left in place so the user can
                # open the file, fix it, and recheck without losing anything.

                # Reload project list to safely render Step 1 again
                projects_info = {}
                if os.path.exists(PROJECTS_FOLDER):
                    for folder_name in os.listdir(PROJECTS_FOLDER):
                        folder_path = os.path.join(PROJECTS_FOLDER, folder_name)
                        if os.path.isdir(folder_path):
                            sig_file = None
                            for f in os.listdir(folder_path):
                                if f.lower().startswith("signature links") and (f.lower().endswith(".xlsx") or f.lower().endswith(".xls")):
                                    sig_file = f
                                    break
                            projects_info[folder_name] = sig_file

                projects_json = json.dumps(projects_info)

                # Re-render Step 1 with duplicate errors, plus enough context for the
                # "Open Excel File" / "Recheck File" buttons to target the exact file.
                return _show('pdf.html',
                                       step=1,
                                       existing_projects=list(projects_info.keys()),
                                       projects_json=projects_json,
                                       duplicate_errors=duplicate_errors,
                                       duplicate_project_name=os.path.basename(project_folder),
                                       duplicate_excel_filename=os.path.basename(excel_path))
                
            return _show('pdf.html', step=2, tabs_data=tabs_data, excel_path=excel_path, project_name=project_name, installer_tabs=installer_tabs)
            
        # STEP 2: Process the PDFs and Go to Success Screen
        elif step == '2':
            excel_path = request.form.get('excel_path')
            project_name = request.form.get('project_name')
            # Close any copy of the Signature Links file open in Excel, without saving
            if excel_path and os.path.isfile(excel_path):
                close_if_open_elsewhere(excel_path)
            temp_dir = os.path.join(BASE_DIR, 'temp_pdf_engine')
            os.makedirs(temp_dir, exist_ok=True)
            # ---> THE FIX: Define project_folder HERE, before any loops!
            project_folder = os.path.join(PROJECTS_FOLDER, os.path.basename(project_name))
            os.makedirs(project_folder, exist_ok=True)

            all_sheets_data = {}
            barcode_lookup = {}

            try:
                with pd.ExcelFile(excel_path) as xls:
                    barcode_lookup = read_divider_barcodes(xls)
                    for sheet in xls.sheet_names:
                        all_sheets_data[sheet] = pd.read_excel(xls, sheet_name=sheet)
            except Exception as e:
                return f"Could not load Excel file for mapping: {e}", 500

            generated_files = []

            
                # Loop through ONLY the files that were actually uploaded
            for key, file in request.files.items():
                
                # THE FIX: Look for the '---' separator
                if key.startswith('pdf---') and file and file.filename != '':
                    parts = key.split('---')
                    
                    if len(parts) == 3:
                        tab_name = parts[1]
                        pack_name = parts[2]
                        
                        df = all_sheets_data.get(tab_name)
                        
                        # Failsafe: Prevent crashes if the tab name is completely invalid
                        if df is None:
                            print(f"[ERROR] Could not find sheet '{tab_name}' in Excel data.")
                            continue 
                        
                        # 1. Build the mapping for THIS specific tab and pack
                        store_mapping = {}
                        
                        # Failsafe: Prevent KeyErrors if the pack name isn't found
                        if pack_name not in df.columns:
                            print(f"[ERROR] Pack '{pack_name}' not found in Tab '{tab_name}'.")
                            continue
                            
                        for index, row in df.iterrows():
                            store_cell = str(row.iloc[0]).strip()
                            code_cell = str(row[pack_name]).strip()
                            
                            if store_cell != 'nan' and code_cell != 'nan':
                                store_mapping[store_cell] = code_cell

                        # 2. Check the divider flag from the form
                        checkbox_key = f"divider---{tab_name}---{pack_name}"
                        add_dividers_flag = request.form.get(checkbox_key) == "true"
                        
                        # 3. Generate the clean "_Shuffled" filename
                        safe_orig = secure_filename(file.filename)
                        name_part, ext_part = os.path.splitext(safe_orig) 
                        final_filename = f"{name_part}_Shuffled{ext_part}"
                        
                        # 4. Define Paths & Save temp file
                        temp_pdf_path = os.path.join(temp_dir, secure_filename(file.filename))
                        output_pdf_path = os.path.join(project_folder, final_filename)
                        file.save(temp_pdf_path)
                        
                        # 5. Run the Engine!
                        process_and_shuffle_pdf(
                            input_pdf_path=temp_pdf_path,
                            store_mapping=store_mapping,
                            output_pdf_path=output_pdf_path,
                            signature_header=pack_name,
                            add_dividers=add_dividers_flag,
                            # Installer tabs only: job-number barcodes on each code's divider
                            divider_barcodes=barcode_lookup.get((tab_name, pack_name)) if add_dividers_flag else None
                        )
                        
                        # Log it for the download screen
                        generated_files.append(final_filename)

            try:
                shutil.rmtree(temp_dir)
            except Exception:
                pass

            return _show('pdf.html', step=3, project_name=project_name, generated_files=generated_files)

    # GET REQUEST: Fetch existing projects AND look for their Signature links files
    projects_info = {}
    if os.path.exists(PROJECTS_FOLDER):
        for folder_name in os.listdir(PROJECTS_FOLDER):
            folder_path = os.path.join(PROJECTS_FOLDER, folder_name)
            if os.path.isdir(folder_path):
                sig_file = None
                for f in os.listdir(folder_path):
                    # FIXED: Case-insensitive check for the GET request as well
                    if f.lower().startswith("signature links") and (f.lower().endswith(".xlsx") or f.lower().endswith(".xls")):
                        sig_file = f
                        break
                projects_info[folder_name] = sig_file
                
    projects_json = json.dumps(projects_info)
    return _show('pdf.html', step=1, existing_projects=list(projects_info.keys()), projects_json=projects_json)


# Make sure you have an upload folder configured in your app
# app.config['UPLOAD_FOLDER'] = 'uploads/'



def _edits_field(form):
    try:
        data = json.loads(form.get('consignment_edits') or '{}')
    except ValueError:
        return {}
    return data if isinstance(data, dict) else {}


def _consignment_edits(form):
    """Edits made in the preview's consignment table, sent back as JSON in a hidden field."""
    return {str(k): v for k, v in _edits_field(form).items() if isinstance(v, dict)}


def _read_tab(filepath, header_row, tab, columns=None, mode='labels'):
    """parse_packing_data() (Courier Import: read_allocation()), plus store names for files whose only address info
    is one combined cell. columns: the tab's Column mapping set by hand in the preview (_column_choices)."""
    read = read_allocation if mode == 'courier' else parse_packing_data
    pack_groups, last_row, found_headers, warnings = read(filepath, header_row, sheet_name=tab, columns=columns)
    return fill_store_names(pack_groups), last_row, found_headers, warnings


def _column_choices(form, tab, mode='labels'):
    """The preview's Column mapping for a tab: {field: column letter, or '-' for not used}; automatic ones left out.
    Sent as 'colmap::<tab>::<field>' (the page also puts back what was chosen for this file last time). Courier
    Import's fields are the address list's, and 'pack_<n>' for its pack columns."""
    choices = {}
    prefix = f'colmap::{tab}::'
    if mode == 'courier':
        fields = [k[len(prefix):] for k in form if k.startswith(prefix)]
        fields = [f for f in fields if f in IMPORT_FIELD_NAMES or PACK_FIELD.match(f)]
    else:
        fields = [f for f, _, _ in MAPPABLE_FIELDS]
    for field in fields:
        value = form.get(prefix + field, '').strip().upper()
        if value == '-' or re.fullmatch(r'[A-Z]{1,3}', value):
            choices[field] = value
    return choices


def _mapping_view(filepath, header_row, tab, choices, mode='labels'):
    """What the Column mapping shows for a tab: each field, the header it's matched to and that header's cell."""
    try:
        columns = header_columns(filepath, header_row, tab)
    except Exception as e:
        log_pack.warning("  tab %s: couldn't read header row %s for the column mapping: %s", tab, header_row, e)
        columns = []
    text_of = dict(columns)
    if mode == 'courier':
        try:
            found, packs, _ = scan_columns(filepath, header_row, tab)
        except Exception as e:
            log_import.warning("  tab %s: couldn't find the columns in header row %s: %s", tab, header_row, e)
            found, packs = {}, {}
        auto = {f: get_column_letter(i) for f, i in found.items()} | {f"pack_{n}": get_column_letter(i) for n, i in packs.items()}
        fields = mapping_fields(packs, choices)
    else:
        auto = {f: get_column_letter(i) for f, i in detect_columns([(column_index_from_string(l), t) for l, t in columns]).items()}
        fields = MAPPABLE_FIELDS
    items = []
    for field, label, required in fields:
        chosen = choices.get(field, '')
        used = None if chosen == '-' else (chosen or auto.get(field))
        items.append({'field': field, 'label': label, 'required': required, 'auto': auto.get(field, ''),
                      'auto_text': text_of.get(auto.get(field, ''), ''), 'chosen': chosen, 'used': used or '',
                      'used_text': text_of.get(used, '') if used else '',
                      'cell': f"{used}{header_row}" if used else ''})
    return {'header_row': header_row, 'columns': columns, 'items': items, 'by_hand': sum(1 for i in items if i['chosen'])}


def _receiver_groups(consignments):
    """Consignments for the same receiver (a store, or an installer) at different addresses, put next to each other
    so they can be merged into one in the preview. Names count as the same when they're equal ignoring capitals,
    spacing and punctuation, or one small typo apart. Returns the consignments in the new order, each in a group
    carrying c['group'] = {'id', 'first', 'size', 'receiver', 'similar', 'members'}; 'similar' says whether every
    address looks like the same place written differently (similar_addresses) or some really differ."""
    from difflib import SequenceMatcher
    keys = []  # one per receiver name, in order of first appearance
    group_of = {}
    for c in consignments:
        name = _norm(c['destination']['receiver'])
        if not name:
            continue
        key = next((k for k in keys if k == name or (min(len(k), len(name)) > 5 and SequenceMatcher(None, k, name).ratio() >= 0.92)), None)
        if key is None:
            keys.append(name)
            key = name
        group_of[c['number']] = key
    sizes = {}
    for k in group_of.values():
        sizes[k] = sizes.get(k, 0) + 1
    ordered, placed = [], set()
    for c in consignments:
        if c['number'] in placed:
            continue
        key = group_of.get(c['number'])
        members = [m for m in consignments if group_of.get(m['number']) == key] if key and sizes[key] > 1 else [c]
        if len(members) > 1:
            addressed = [m for m in members if not m.get('no_address')]
            similar = all(similar_addresses(one_line(addressed[0]['destination']), one_line(m['destination'])) for m in addressed[1:])
            info = [{'number': m['number'], 'receiver': m['destination']['receiver'],
                     'address': one_line(m['destination']) or 'no address', 'cartons': len(m['cartons']),
                     'no_address': bool(m.get('no_address')),
                     'service': m.get('service_code') or '', 'service_used': m.get('service_code_used') or '',
                     'sources': sorted({x['source_id'] for x in m['cartons']}),
                     'fields': {k: m['destination'][k] for k in ('line1', 'line2', 'suburb', 'state', 'postcode', 'country')}}
                    for m in members]
            for i, m in enumerate(members):
                m['group'] = {'id': f"g{members[0]['number']}", 'first': i == 0, 'size': len(members),
                              'receiver': members[0]['destination']['receiver'], 'similar': similar,
                              'services': sorted({x['service_used'] for x in info if x['service_used']}),
                              'members': info if i == 0 else None}
        for m in members:
            ordered.append(m)
            placed.add(m['number'])
    return ordered


def _import_refs(groups, tab_number):
    """Courier Import: each pack's code, T<tab>P<pack> ('T1P2' = pack 2 of the first selected tab)."""
    for g in groups.values():
        g['item_ref'] = f"T{tab_number}P{g['pack_no']}"
    return groups


def _import_item_refs(cartons, groups):
    """Courier Import: each carton's Item Reference, its pack's code and its store: 'T1P2 Corio Village' (the Store
    Name, else the Receiver Name). In the table and both CSVs."""
    for x in cartons:
        x['ref_code'] = groups[x['pack_key']]['item_ref']
        x['item_reference'] = f"{x['ref_code']} {x['store']}".strip()


def _import_csvs(project_dir, reference, project, csv_formats, cartons, consignments, fixed, form, log):
    """Courier Import's Generate: the courier CSVs chosen, with each pack's T<tab>P<pack> as its Item Reference
    (Open360's Item Reference, OpenFreight's Reference). Returns the file names."""
    base_name = f"{reference} - {project}" if reference else project
    refs = [(x['item_reference'], x['item_reference']) for x in cartons]
    csv_files = []
    if 'open360' in csv_formats:
        name = f"{base_name} - Open360.csv"
        with app_log.step(log, f"  Open360 CSV {name}"):
            write_open360_csv(os.path.join(project_dir, name), open360_items(cartons, consignments), reference, item_refs=refs)
        log.info("    %d row(s), Shipment Reference %s", len(refs), reference)
        csv_files.append(name)
    if 'openfreight' in csv_formats:
        name = f"{base_name} - OpenFreight.csv"
        with app_log.step(log, f"  OpenFreight CSV {name}"):
            export, _ = clean_export(_export_values(form))
            write_courier_csv(os.path.join(project_dir, name), cartons, consignments, reference, fixed, export=export)
            abroad = sum(1 for c in consignments if (c['destination'].get('country') or 'AU').upper() != 'AU')
            if abroad:
                log.info("    export details on %d consignment(s) outside Australia: %s", abroad, export)
        csv_files.append(name)
    return csv_files


def _carton_weights(form):
    """Carton weights typed in the preview's consignment table: {pack key: kg}, from a hidden JSON field."""
    try:
        data = json.loads(form.get('carton_weights') or '{}')
    except ValueError:
        return {}
    return {str(k): v for k, v in data.items()} if isinstance(data, dict) else {}


def _book_offered(form):
    """Excel addresses the address book has already been asked about, so a Reset to Excel sticks."""
    offered = _edits_field(form).get('__offered__', [])
    return [str(x) for x in offered] if isinstance(offered, list) else []


def _same_address(a, b):
    return all(str(a.get(k) or '').strip() == str(b.get(k) or '').strip() for k in EDITABLE_FIELDS) \
        and bool(a.get('authority_to_leave')) == bool(b.get('authority_to_leave'))


def _apply_address_book(pack_groups, edits, offered):
    """Fills in addresses the book has learned for these Excel spellings (each spelling is offered once).

    A pack with a receiver name but no address (a file with only receiver names) that the book hasn't learned
    yet is looked up by receiver: when every saved entry for that receiver is at one address, it's filled in.
    Several different addresses are left for the user to choose from in the preview."""
    sources = source_destinations(pack_groups)
    new = [sid for sid in sources if sid not in edits and sid not in offered]
    try:
        learned = ADDRESS_BOOK.lookup_aliases(new)
        for sid in new:
            dest = sources[sid]
            if sid in learned:
                continue
            if dest['postcode']:
                if not dest['receiver']:
                    # Address with no name in it: the book knows who's at that address
                    here = ADDRESS_BOOK.at_address(dest)
                    if here:
                        learned[sid] = {**dest, **{k: here[0][k] for k in ('receiver', 'contact')},
                                        'authority_to_leave': dest['authority_to_leave']}
                continue
            saved = ADDRESS_BOOK.find_by_receiver(dest['receiver'])
            places = {(e['line1'], e['line2'], e['suburb'], e['state'], e['postcode']) for e in saved}
            if len(places) == 1:
                learned[sid] = saved[0]  # most used first
    except Exception as e:
        log_book.exception("Address book lookup failed: %s", e)
        learned = {}
    filled = 0
    for sid, entry in learned.items():
        if not _same_address(entry, sources[sid]):
            edits[sid] = {**{k: entry[k] for k in EDITABLE_FIELDS}, 'authority_to_leave': entry['authority_to_leave'], 'source': 'book'}
            filled += 1
            log_book.debug("Filled from the book: %s -> %s", sid, one_line(entry))
    log_book.info("Address book: %d Excel addresses, %d new to ask about, %d filled from the book", len(sources), len(new), filled)
    return edits, offered + new


SUGGESTION_FIELDS = EDITABLE_FIELDS + ('authority_to_leave',)


def _closest(address, cache=None):
    """Closest saved addresses for the ✏️ form, best first, each with the fields a click fills in."""
    return [{**s, 'json': json.dumps({k: s[k] for k in SUGGESTION_FIELDS})}
            for s in closest_addresses(address, ADDRESS_BOOK.candidates(address, cache))]


def _check_against_book(consignments, edits):
    """Sets each consignment's 'book' status for the preview:
      book       filled in from the address book (a spelling it learned)
      matched    the book holds this address, under this receiver or another one (an address can have several
                 receivers / Attn names). A receiver or Attn new to it gets 'book_note': it's added on Generate.
      checked    not in the book, but the user has reviewed it in the edit form
      unverified not in the book: listed first, with 'suggestions' (closest saved addresses, best first)
      missing    no address at all (only a receiver name): listed first of all, with the receiver's saved
                 addresses as suggestions. Left out of the courier CSV until it has one."""
    cache = {}
    for c in consignments:
        d = c['destination']
        c['book'], c['book_note'], c['suggestions'] = 'unverified', '', []
        # Every ✏️ form lists the closest saved addresses (for a row filled from the book, that entry comes
        # first), so undoing a fill or a wrong edit leaves the good address one click away
        try:
            c['suggestions'] = _closest(d, cache)
        except Exception as e:
            log_book.exception("Address book suggestions failed for #%s: %s", c.get('number'), e)
        if c.get('no_address'):
            c['book'] = 'missing'
            continue
        if c['edit_source'] == 'book':
            c['book'] = 'book'
            continue
        try:
            here = ADDRESS_BOOK.at_address(d)
            if here:
                c['book'] = 'matched'
                same_receiver = [e for e in here if _norm(e['receiver']) == _norm(d['receiver'])]
                if not same_receiver:
                    others = list(dict.fromkeys(e['receiver'] for e in here))
                    c['book_note'] = (f"Address saved for {', '.join(others[:3])}{' …' if len(others) > 3 else ''}. "
                                      f"This receiver will be added to the address book on Generate.")
                elif d['contact'] and not any(_norm(e['contact']) == _norm(d['contact']) for e in same_receiver):
                    saved = [e['contact'] for e in same_receiver if e['contact']]
                    c['book_note'] = (f"New Attn for this address{' (saved: ' + ', '.join(saved[:3]) + ')' if saved else ''}. "
                                      f"It will be added to the address book on Generate.")
            elif c['edited'] or edits.get(c['id'], {}).get('checked'):
                c['book'] = 'checked'
        except Exception as e:
            log_book.exception("Address book check failed for #%s: %s", c.get('number'), e)
            c['book'] = 'checked'  # don't flag every row because the book couldn't be read
    # No address first, then not in the book; otherwise keep consignment order
    return sorted(consignments, key=lambda c: {'missing': 0, 'unverified': 1}.get(c['book'], 2))


@app.route('/address-book/update', methods=['GET', 'POST'])
def bulk_update_addresses():
    """Update the address book in bulk from the courier portal's CSV of verified addresses.
    Rows are matched against every project's label map and the address book (see address_import.py)."""
    kept = _kept_view()
    if kept:
        return kept
    if request.method == 'POST' and 'apply' in request.form:
        try:
            proposals = json.loads(request.form.get('proposals') or '[]')
        except ValueError:
            proposals = []
        actions = {k[len('action_'):]: v for k, v in request.form.items() if k.startswith('action_')}
        with app_log.step(log_book, f"Bulk update: apply {len(proposals)} reviewed address(es)"):
            applied, confirmed, added, merged, joined, problems = apply_review(
                proposals, set(request.form.getlist('choose')), ADDRESS_BOOK, actions if actions else None)
        log_book.info("  updated %s, confirmed %s, added %s, merged %s, joined %s, problems %d",
                      *(len(v) if isinstance(v, (list, dict, set)) else v for v in (applied, confirmed, added, merged, joined)),
                      len(problems))
        for p in problems:
            log_book.warning("  problem: %s", p)
        return _show('address_verify.html', done=True, applied=applied, confirmed=confirmed, added=added,
                               merged=merged, joined=joined, problems=problems)

    if request.method == 'POST':
        upload = request.files.get('file')
        if not upload or not upload.filename:
            return _show('address_verify.html', error="Choose the file of verified addresses (CSV or Excel).")
        if not upload.filename.lower().endswith(('.csv', '.xlsx', '.xlsm', '.xls')):
            return _show('address_verify.html', error="Use a CSV or Excel file (.csv, .xlsx or .xls).")
        try:
            with app_log.step(log_book, f"Bulk update: read {upload.filename}"):
                proposals, columns = build_review(upload.read(), upload.filename, all_label_maps(PROJECTS_FOLDER), ADDRESS_BOOK)
            log_book.info("  %d address(es) to review; columns used %s", len(proposals), sorted(columns))
        except ImportProblem as e:
            log_book.warning("Bulk update: %s", e)
            return _show('address_verify.html', error=str(e))
        if not proposals:
            return _show('address_verify.html', error="No addresses found in that file.")
        return _show('address_verify.html', proposals=proposals, proposals_json=json.dumps(proposals),
                               filename=upload.filename, columns=sorted(columns))

    return _show('address_verify.html')


def _spec_rows_from(form):
    formula = {p: {'weight': form.get(f'formula_{p}_weight'), 'item_type': form.get(f'formula_{p}_item_type'),
                   'third_size': form.get(f'formula_{p}_third_size')} for p in FORMULA_PREFIXES}
    named = []
    for i in form.getlist('row'):
        named.append({k: form.get(f'named_{i}_{k}') for k in ('name', 'length', 'width', 'height', 'weight', 'item_type')}
                     | {'remove': form.get(f'named_{i}_remove'), 'seen_at': form.get(f'named_{i}_seen_at')})
    return formula, named


@app.route('/packing-specs', methods=['GET', 'POST'])
def packing_specs_page():
    """Sizes, weights and item types for every Packing Spec, used for the courier CSV."""
    errors, saved = {}, request.args.get('saved') == '1'
    if request.method == 'POST':
        formula, named = _spec_rows_from(request.form)
        errors = SPEC_STORE.save_page(formula, named)
        log_specs.info("Packing Specs saved: %d named row(s), %s", len(named), f"{len(errors)} error(s): {errors}" if errors else "no errors")
        if not errors:
            return redirect(url_for('packing_specs_page', saved=1))
        data = {'formula': formula, 'named': [r for r in named if not r.get('remove')]}
        named_rows = list(enumerate(named))
    else:
        data = SPEC_STORE.load()
        named_rows = list(enumerate(data['named']))
    code = request.args.get('code', '').strip()
    tried = None
    if code:
        found = parse_formula(code, SPEC_STORE.load()['formula']['FP']['third_size'])
        tried = {'code': code, 'kind': found[0] if found else None, 'size': found[1] if found else None,
                 'named': SPEC_STORE.resolver()(code) if not found else None}
    return render_template('packing_specs.html', formula=data['formula'], named_rows=named_rows, errors=errors,
                           saved=saved, prefixes=FORMULA_PREFIXES, labels=FORMULA_LABELS, examples=FORMULA_EXAMPLES,
                           item_types=ITEM_TYPES, tried=tried)


@app.route('/address-book')
def address_book_page():
    return render_template('address_book.html', total=ADDRESS_BOOK.count())


@app.route('/api/addresses')
def api_addresses():
    rows, has_more = ADDRESS_BOOK.search(request.args.get('q', ''),
                                         request.args.get('limit', 50, type=int),
                                         request.args.get('offset', 0, type=int))
    return {'rows': rows, 'has_more': has_more}


def _export_values(form):
    """The export details on the page: what's been typed (export_<field>), else the saved defaults."""
    typed = {k: form.get(f'export_{k}') for k in EXPORT_DEFAULTS if form.get(f'export_{k}') is not None}
    return clean_export({**load_export_defaults(EXPORT_FILE), **typed})[0]


@app.route('/api/export-defaults', methods=['GET', 'POST'])
def api_export_defaults():
    """The saved export defaults (GET), or keep these as the defaults (POST {field: value})."""
    if request.method == 'GET':
        return load_export_defaults(EXPORT_FILE)
    saved, problems = save_export_defaults(EXPORT_FILE, request.get_json(silent=True) or {})
    log_pack.info("Export defaults saved: %s%s", saved, f" (not usable, kept as before: {problems})" if problems else '')
    return {'saved': saved, 'problems': problems}


@app.route('/api/addresses/suggest')
def api_address_suggest():
    """Closest saved addresses to what's typed in a consignment's ✏️ form so far (any field), best first."""
    address = {k: request.args.get(k, '').strip() for k in EDITABLE_FIELDS}
    if not any(address.values()):
        return {'rows': []}
    return {'rows': [{k: s[k] for k in SUGGESTION_FIELDS + ('id', 'address', 'why', 'verified_at')} for s in _closest(address)]}


@app.route('/api/addresses', methods=['POST'])
def api_address_create():
    try:
        row = ADDRESS_BOOK.create(request.get_json(silent=True) or {})
        log_book.info("Address added #%s: %s", row.get('id'), row.get('receiver'))
        return row, 201
    except AddressBookError as e:
        log_book.warning("Address not added: %s", e)
        return {'error': str(e)}, 400


@app.route('/api/addresses/<int:address_id>', methods=['PUT'])
def api_address_update(address_id):
    try:
        row = ADDRESS_BOOK.update(address_id, request.get_json(silent=True) or {})
        log_book.info("Address #%s updated: %s", address_id, row.get('receiver') if isinstance(row, dict) else row)
        return row
    except AddressBookError as e:
        log_book.warning("Address #%s not updated: %s", address_id, e)
        return {'error': str(e)}, 400


@app.route('/api/addresses/<int:address_id>', methods=['DELETE'])
def api_address_delete(address_id):
    if ADDRESS_BOOK.delete(address_id):
        log_book.info("Address #%s deleted", address_id)
        return '', 204
    return {'error': 'That address no longer exists.'}, 404


@app.route('/packing-labels', methods=['GET', 'POST'])
def create_packing_labels():
    """Packing labels (PDF) and courier CSVs from a packing list (a Packing Spec column, rows merged per box)."""
    return _distribution_page('labels')


@app.route('/courier-import', methods=['GET', 'POST'])
def courier_import():
    """Courier Import: only the courier CSV, from an address list with a column per pack (courier_import.py)."""
    return _distribution_page('courier')


def _distribution_page(mode):
    """Packing Labels and Courier Import share one page and its steps: upload, tabs and header rows, the preview
    (Column mapping, consignment table with the address book, ✏️ edits, merges, service codes, weights), the
    courier CSV choice and Generate. mode 'labels': packing labels PDF + courier CSVs + label map, from a packing
    list. mode 'courier': the courier CSV(s) only, from an address list whose pack columns are headed 1, 2, 3 …;
    each pack's Item Reference is T<tab>P<pack> ('T1P2')."""
    kept = _kept_view()
    if kept:
        return kept
    courier_mode = mode == 'courier'
    log = log_import if courier_mode else log_pack
    show = lambda **context: _show('packing_labels.html', mode=mode, **context)
    if request.method == 'GET':
        return show()

    # STEP 1: Handle File Upload
    if 'file' in request.files:
        file = request.files['file']
        original_filename = file.filename or ''
        extension = os.path.splitext(original_filename)[1].lower()
        if not original_filename:
            return show(page_error="Choose an Excel file to scan.")
        if extension not in ('.xlsx', '.xls'):
            return show(page_error="Unsupported file type. Choose an .xlsx or .xls file.")
        filename = secure_filename(original_filename)
        if not filename:
            return show(page_error="The uploaded filename is not valid.")
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        try:
            close_if_open_elsewhere(filepath)
            file.save(filepath)
            filepath = convert_legacy_excel_to_xlsx(filepath)
            filename = os.path.basename(filepath)

            wb = openpyxl.load_workbook(filepath, read_only=True)
            try:
                tabs = wb.sheetnames
            finally:
                wb.close()

            log.info("Upload %s: %d tab(s) %s", filename, len(tabs), tabs)
            return show(tabs=tabs, filename=filename)
        except Exception as e:
            log.exception("Could not scan upload %s", original_filename)
            return show(page_error=f"Could not scan '{original_filename}': {e}")
    # Base variables for Step 2 & 3
    filename = request.form.get('filename')
    if not filename:
        return redirect(url_for(request.endpoint))
        
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], os.path.basename(filename))
    if not os.path.isfile(filepath):
        return show(page_error="The uploaded file is no longer available. Scan it again.")
    all_tabs = request.form.getlist('all_tabs')
    selected_tabs = request.form.getlist('selected_tabs')
    
    # Save inputs so UI remembers what user typed
    user_inputs = {}
    for tab in all_tabs:
        user_inputs[tab] = {
            'selected': tab in selected_tabs,
            'header_row': request.form.get(f'header_row_{tab}', '1')
        }

    # Action: Update Previews
    if 'preview' in request.form:
        log.info("Preview %s: tabs %s", filename, selected_tabs)
        started = time.perf_counter()
        # The upload may be open in Excel (the Open button): changes not saved yet are saved first, so the preview
        # shows what's in Excel. The page says which saved version it read.
        excel = save_if_open_elsewhere(filepath)
        saved_at = datetime.fromtimestamp(os.path.getmtime(filepath))
        file_note = {'saved_at': saved_at.strftime('%d/%m/%Y %H:%M:%S'), 'excel': excel}
        log.info("  file saved %s; open in Excel: %s", file_note['saved_at'], excel or 'no')
        previews = []
        header_rows = {}  # each tab's header row as a number (a blank, 0 or text one is read as row 1)
        for tab in selected_tabs:
            choices = _column_choices(request.form, tab, mode)
            try:
                header_row = max(1, int(user_inputs[tab]['header_row']))
            except ValueError:
                header_row = 1
            header_rows[tab] = header_row
            mapping = _mapping_view(filepath, header_row, tab, choices, mode)
            if choices:
                log.info("  tab %s: columns set by hand %s", tab, choices)
            try:
                pack_groups, last_row, found_headers, warnings = _read_tab(filepath, header_row, tab, choices, mode)
                
                total_packs = len(pack_groups)
                total_stores = len(set(g['store_name'] for g in pack_groups.values()))
                log.info("  tab %s (header row %d): %d packs, %d stores, last row %s, %d warning(s)",
                              tab, header_row, total_packs, total_stores, last_row, len(warnings))
                for w in warnings:
                    log.debug("    warning: %s", w)
                
                previews.append({
                    'sheet_name': tab, 
                    'total_packs': total_packs, 
                    'total_stores': total_stores, 
                    'last_row': last_row,
                    'found_headers': found_headers, # Pass headers to the UI
                    'warnings': warnings,
                    'mapping': mapping,
                    'error': None
                })
            except PackCheckError as e:
                log.warning("  tab %s: %d problem(s) stop it: %s", tab, len(e.issues),
                                 "; ".join(f"{i.get('rows', '')} {i.get('col', '')} {i.get('text', '')}".strip() for i in e.issues[:5]))
                previews.append({'sheet_name': tab, 'error': True, 'issues': e.issues, 'warnings': e.warnings, 'mapping': mapping})
            except ValueError as e:  # a column that can't be found: picked in the Column mapping
                log.warning("  tab %s: %s", tab, e)
                previews.append({'sheet_name': tab, 'error': True, 'issues': [{'rows': '', 'col': '', 'text': str(e), 'detail': ''}],
                                 'warnings': [], 'mapping': mapping})
            except Exception as e:
                log.exception("  tab %s could not be read", tab)
                previews.append({'sheet_name': tab, 'error': True, 'issues': [{'rows': '', 'col': '', 'text': str(e), 'detail': ''}],
                                 'warnings': [], 'mapping': mapping})

        courier = None
        if previews and not any(p['error'] for p in previews):
            combined = {}
            for tab in selected_tabs:
                groups, _, _, _ = _read_tab(filepath, header_rows[tab], tab, _column_choices(request.form, tab, mode), mode)
                combined.update({f"{tab} - {k}": v for k, v in groups.items()})
                if courier_mode:
                    _import_refs(groups, selected_tabs.index(tab) + 1)
            edits, offered = _apply_address_book(combined, _consignment_edits(request.form), _book_offered(request.form))
            options = service_options(load_service_usage(SERVICE_USAGE_FILE))
            most_used = options[0]['code'] if options[0]['used'] else ''
            fixed = {**DEFAULT_FIXED, 'service_code': most_used,
                     **{k: request.form[k] for k in ('who_pays', 'charge_account', 'service_code', 'reference') if k in request.form}}
            weights = _carton_weights(request.form)
            consignments, cartons, courier_warnings = build_consignments(combined, edits, fixed['service_code'],
                                                                         SPEC_STORE.resolver(), weights)
            if courier_mode:
                _import_item_refs(cartons, combined)
            consignments = _receiver_groups(_check_against_book(consignments, edits))  # one receiver's next to each other
            # Specs the Packing Specs page doesn't know yet are listed there for the user to fill in
            try:
                SPEC_STORE.note_unknown(x['packing_spec'] for x in cartons if x['spec_kind'] == 'unknown')
            except OSError as e:
                log_specs.warning("Could not note new packing specs: %s", e)
            # Courier Import: the job numbers above the header rows (or in the file name)
            series, all_series = (reference_in_sheets(filepath, [(t, header_rows[t]) for t in selected_tabs], filename)
                                  if courier_mode else detect_series(combined))
            if len(all_series) > 1:
                courier_warnings.insert(0, f"Job numbers use more than one series ({', '.join(all_series)}). Using {series}; change it below if needed.")
            status = {}
            for c in consignments:
                status[c['book']] = status.get(c['book'], 0) + 1
            log.info("  %d cartons in %d consignments | job %s (found %s) | service %s | address book: %s",
                          len(cartons), len(consignments), series or '-', all_series or 'none', fixed.get('service_code') or '-',
                          ", ".join(f"{k} {v}" for k, v in sorted(status.items())))
            for w in courier_warnings:
                log.info("  courier check: %s", w)
            log.info("Preview ready in %.2fs", time.perf_counter() - started)
            courier = {
                'consignments': [{**c, 'address': one_line(c['destination']),
                                  'source_text': (c['original']['raw'] if c['original']['postcode'] in c['original']['raw']
                                                  else ", ".join(x for x in (c['original']['raw'], c['original']['suburb'], c['original']['state'], c['original']['postcode']) if x))
                                                 or f"{c['original']['receiver']} (no address in the file)",
                                  'original_json': json.dumps({k: c['original'][k] for k in EDITABLE_FIELDS + ('authority_to_leave',)})}
                                 for c in consignments],
                'states': sorted({c['destination']['state'] for c in consignments if c['destination']['state']}),
                'carton_count': len(cartons),
                'weights_json': json.dumps(weights),
                'spec_gaps': any(x['spec_kind'] in ('unknown', 'incomplete') for x in cartons),
                'unverified': sum(c['book'] == 'unverified' for c in consignments),
                'missing': sum(c['book'] == 'missing' for c in consignments),
                'warnings': courier_warnings,
                'fixed': {'reference': series, **fixed},
                'edits_json': json.dumps({**edits, '__offered__': offered}),
                # Export details for deliveries outside Australia (OpenFreight): as typed so far, else the defaults
                'export': _export_values(request.form),
                'international': sum(1 for c in consignments if (c['destination'].get('country') or 'AU').upper() != 'AU'),
                'service_options': options,
            }

        return show(tabs=all_tabs, filename=filename, user_inputs=user_inputs, previews=previews, courier=courier,
                               file_note=file_note)

    # Action: Generate Final PDF & Save to Project
    if 'generate' in request.form:
        combined_pack_groups = {}
        
        # Grab the sorted layout order from the frontend, and filter out 'empty' slots
        attribute_order_str = request.form.get('attribute_order', 'thumbnail,desc,dimension,job_no,barcode,qty')
        attribute_order = [attr for attr in attribute_order_str.split(',') if attr and attr != 'empty']
        
        missing = [label for key, label in (('who_pays', 'Who Pays'), ('service_code', 'Service Code'), ('reference', 'Consignment Reference'),
                                            ('project_name', 'Project Name'))
                   if not request.form.get(key, '').strip()]
        csv_formats = [f for f in CSV_FORMATS if f in request.form.getlist('courier_csv')]
        if not csv_formats:
            missing.append('Courier CSV (Open360, OpenFreight or both)')
        log.info("Generate %s: tabs %s, project %r, reference %r, service %r", filename, selected_tabs,
                      request.form.get('project_name', ''), request.form.get('reference', ''), request.form.get('service_code', ''))
        if missing:
            log.warning("Generate stopped: %s not filled in", ", ".join(missing))
            return show(tabs=all_tabs, filename=filename, user_inputs=user_inputs,
                                   page_error=f"Fill in {', '.join(missing)} before generating.")
        sender = None  # the portal uses the account's sender (the Open360 CSV leaves it empty)

        # Edits made in Excel and not saved yet are saved first (what the preview showed is what's generated);
        # then that copy is closed, so nothing stops the file moving into the project folder
        excel = save_if_open_elsewhere(filepath)
        if excel == 'busy':
            log.warning("Generate stopped: %s is open in Excel in the middle of editing a cell", filename)
            return show(tabs=all_tabs, filename=filename, user_inputs=user_inputs,
                                   page_error=f"{filename} is open in Excel in the middle of editing a cell, so it can't be saved. "
                                              f"Press Enter (or Esc) in Excel, then Update Previews and Generate again.")
        if excel == 'saved':
            log.info("  saved the changes made in Excel to %s before generating", filename)
        close_if_open_elsewhere(filepath)

        project_dir = None
        started = time.perf_counter()
        tab_of = {}  # pack key -> its tab's number among the selected tabs, for the Item Reference ('01T2QWXK Store')
        try:
            # 1. Parse Data
            sheet_warnings = []
            for tab in selected_tabs:
                header_row = int(user_inputs[tab]['header_row'])
                pack_groups, _, _, tab_warnings = _read_tab(filepath, header_row, tab, _column_choices(request.form, tab, mode), mode)
                if courier_mode:
                    _import_refs(pack_groups, selected_tabs.index(tab) + 1)
                sheet_warnings += [{**w, 'col': f"{tab} {w.get('col', '')}".strip() if len(selected_tabs) > 1 else w.get('col', '')}
                                   for w in tab_warnings]
                for key, val in pack_groups.items():
                    combined_pack_groups[f"{tab} - {key}"] = val
                    tab_of[f"{tab} - {key}"] = selected_tabs.index(tab) + 1  # T1 = the first selected tab, to the right

            # Each pack's serial in its tab, in the distribution file's order: '01T1', '02T1' ... '01T2'. It's printed
            # before the store name on the packing label and starts the label's Open360 Item Reference
            per_tab = Counter(tab_of.values())
            seen = Counter()
            serials = {}
            for key in combined_pack_groups:
                n = tab_of[key]
                seen[n] += 1
                serials[key] = f"{seen[n]:0{max(2, len(str(per_tab[n])))}d}T{n}"

            # 2. Courier consignments, checked before anything is written or moved
            fixed = {k: request.form.get(k, '').strip() for k in ('who_pays', 'charge_account', 'service_code')}
            reference = request.form.get('reference', '').strip() or detect_series(combined_pack_groups)[0]
            consignments, cartons, courier_warnings = build_consignments(combined_pack_groups, _consignment_edits(request.form),
                                                                         fixed['service_code'], SPEC_STORE.resolver(),
                                                                         _carton_weights(request.form))
            if courier_mode:
                _import_item_refs(cartons, combined_pack_groups)
            all_consignments, all_cartons = consignments, cartons
            consignments, cartons = sendable(consignments, cartons)  # still no address: left out of the CSV
            log.info("  %d packing labels; %d cartons in %d consignments go in the courier CSV, %d left out (no complete address)",
                          len(all_cartons), len(cartons), len(consignments), len(all_cartons) - len(cartons))
            for x in all_cartons:
                if x not in cartons:
                    log.info("    left out: %s (%s)", x['item_reference'], x.get('packing_spec', ''))
            unknown = sorted({c['service_code_used'] for c in consignments} - SERVICE_CODE_SET)
            if unknown:
                raise ValueError(f"Unknown service code {', '.join(repr(u) for u in unknown)}. Pick a service from the list.")

            # 3. Setup Project Folder structure
            raw_project_name = _project_name(request.form.get('project_name'), 'Courier_Import_Job' if courier_mode else 'Packing_Labels_Job')
            safe_project_name = clean_file_name(raw_project_name)
            time_stamp = datetime.now().strftime("%y%m%d_%H%M")
            final_folder_name = f"{safe_project_name}_{time_stamp}"
            project_dir = os.path.join(app.config['UPLOAD_FOLDER'], final_folder_name)
            os.makedirs(project_dir, exist_ok=True)

            # 4. Move the Excel File into the Project
            new_filepath = os.path.join(project_dir, filename)
            shutil.move(filepath, new_filepath)
            log.info("  project folder %s", final_folder_name)

            if courier_mode:
                csv_files = _import_csvs(project_dir, reference, safe_project_name, csv_formats, cartons, consignments,
                                         fixed, request.form, log)
                output_files = csv_files
            else:
                # 5. Generate the PDF with Custom Layout
                # 'Packing Labels Only 1 - J477161 - Lux Test 30 - 261005_1732.pdf' (stitched: 'Complete Labels 1 - ...')
                output_pdf_name = numbered_name(project_dir, "Packing Labels Only", reference, safe_project_name, time_stamp)
                output_pdf = os.path.join(project_dir, output_pdf_name)
                # Each label shows who and where its consignment goes (after address-book fills and ✏️ edits),
                # exactly as in the preview, not the raw Excel text. Every pack gets its label, sent or not.
                by_number = {c['number']: c for c in all_consignments}
                label_addresses = {}
                for x in all_cartons:
                    c = by_number[x['consignment']]
                    d = c['destination']
                    label_addresses[x['pack_key']] = {'address': one_line(d), 'receiver': d['receiver'], 'contact': d['contact'],
                                                      'sent': not c.get('no_address')}
                shipment_reference = reference  # the Open360 Shipment Reference is the Consignment Reference
                report = _generation_report(reference, safe_project_name, shipment_reference, sheet_warnings, courier_warnings,
                                            all_cartons, cartons, consignments, by_number, fixed['service_code'], sender, csv_formats)
                with app_log.step(log, f"  packing labels PDF {output_pdf_name}"):
                    page_info = generate_packing_labels(combined_pack_groups, output_pdf, attribute_order, label_addresses, report,
                                                    serials, {k: (n, selected_tabs[n - 1]) for k, n in tab_of.items()})
                log.info("    %d page(s), %d report page(s)", len(fitz.open(output_pdf)), page_info.get('report_pages', 0))

                # 6. The courier CSVs chosen (TIG Open360's bulk upload and / or OpenFreight's), and the label map used to
                # match courier labels back to packing label pages. Both files carry the same Item References
                # ('01T1QWXK Store': serial + the row's own code), so courier labels booked from either stitch by code.
                base_name = f"{reference} - {safe_project_name}" if reference else safe_project_name
                map_name = f"{base_name}.labelmap.json"
                items = open360_items(cartons, consignments, tab_of, serials)
                open360_refs = open360_item_references(items)
                ref_of = {x['item_reference']: r[0] for x, r in zip(cartons, open360_refs)}
                csv_files = []
                if 'open360' in csv_formats:
                    open360_name = f"{base_name} - Open360.csv"
                    with app_log.step(log, f"  Open360 CSV {open360_name}"):
                        write_open360_csv(os.path.join(project_dir, open360_name), items, shipment_reference, item_refs=open360_refs)
                    csv_files.append(open360_name)
                    log.info("    %d row(s), Shipment Reference %s", len(open360_refs), shipment_reference)
                if 'openfreight' in csv_formats:
                    openfreight_name = f"{base_name} - OpenFreight.csv"
                    with app_log.step(log, f"  OpenFreight CSV {openfreight_name}"):
                        export, export_problems = clean_export(_export_values(request.form))
                        write_courier_csv(os.path.join(project_dir, openfreight_name), cartons, consignments, reference, fixed,
                                          item_refs=ref_of, export=export)
                        abroad = sum(1 for c in consignments if (c['destination'].get('country') or 'AU').upper() != 'AU')
                        if abroad:
                            log.info("    export details on %d consignment(s) outside Australia: %s", abroad, export)
                    csv_files.append(openfreight_name)
                for x, r in zip(cartons, open360_refs):
                    log.debug("    %s -> %s", x['item_reference'], r)
                left_out = [x for x in all_cartons if x not in cartons]
                write_label_map(os.path.join(project_dir, map_name), cartons, consignments, reference,
                                output_pdf_name, csv_files[0], page_info, left_out, sender,
                                {'file': csv_files[0] if 'open360' in csv_formats else None, 'shipment_reference': shipment_reference,
                                 'item_references': {x['item_reference']: r for x, r in zip(cartons, open360_refs)},
                                 'csv_files': csv_files},
                                {'name': safe_project_name, 'generated': time_stamp, 'report': report,
                                 'report_pages': page_info.get('report_pages', 0)})
                log.info("  label map %s", map_name)
                output_files = [output_pdf_name] + csv_files
            record_service_usage(SERVICE_USAGE_FILE, consignments)
            try:
                ADDRESS_BOOK.record_used([(c['destination'], [x['source_id'] for x in c['cartons']]) for c in consignments])
                log_book.info("Address book: %d consignment address(es) saved / marked used", len(consignments))
            except Exception as e:
                log_book.exception("Could not update the address book after Generate: %s", e)

            log.info("Generate done in %.2fs: %s", time.perf_counter() - started, final_folder_name)
            left_out = [x for x in all_cartons if x not in cartons]
            summary = {'items': len(cartons), 'consignments': len(consignments), 'left_out': [
                f"{x['item_reference']} ({x['packing_spec']}) - {x['store']}" for x in left_out]} if courier_mode else None
            return show(generation_complete=True, project_folder=final_folder_name, generated_files=output_files,
                        summary=summary)

        except Exception as e:
            log.exception("Generate FAILED after %.2fs (project folder undone)", time.perf_counter() - started)
            # Undo a half-made project so the upload is back where it was and Generate can be retried
            if project_dir and os.path.isdir(project_dir):
                moved = os.path.join(project_dir, filename)
                if os.path.exists(moved) and not os.path.exists(filepath):
                    shutil.move(moved, filepath)
                shutil.rmtree(project_dir, ignore_errors=True)
            return show(tabs=all_tabs, filename=filename, user_inputs=user_inputs, page_error=str(e))

    return show()

PORT = 5001


def free_port(port):
    """Stops every process listening on the port (e.g. an earlier copy of this app left running),
    then waits until the port is free so this copy can start."""
    try:
        netstat = subprocess.run(['netstat', '-ano', '-p', 'TCP'], capture_output=True, text=True, timeout=15).stdout
    except Exception as e:
        print(f"[WARNING] Could not check port {port}: {e}")
        return
    pids = set()
    for line in netstat.splitlines():
        parts = line.split()
        # Proto, Local Address, Foreign Address, State, PID. A listening socket has no foreign
        # address (0.0.0.0:0 / [::]:0); checking that instead of the word LISTENING works in any language.
        if (len(parts) >= 5 and parts[0] == 'TCP' and parts[1].rsplit(':', 1)[-1] == str(port)
                and parts[2].rsplit(':', 1)[-1] == '0' and parts[-1].isdigit()):
            pids.add(int(parts[-1]))
    pids.discard(os.getpid())
    pids.discard(0)

    for pid in sorted(pids):
        info = subprocess.run(['tasklist', '/FI', f'PID eq {pid}', '/FO', 'CSV', '/NH'],
                              capture_output=True, text=True).stdout.strip()
        name = info.split('","')[0].strip('"') if info.startswith('"') else 'unknown program'
        result = subprocess.run(['taskkill', '/PID', str(pid), '/T', '/F'], capture_output=True, text=True)
        if result.returncode == 0:
            print(f"Stopped {name} (PID {pid}), which was using port {port}.")
        else:
            print(f"[WARNING] Could not stop {name} (PID {pid}) on port {port}: {result.stderr.strip() or result.stdout.strip()}")

    if pids:
        for _ in range(50):  # Windows can take a moment to release the port
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
                try:
                    probe.bind(('127.0.0.1', port))
                    return
                except OSError:
                    time.sleep(0.2)
        print(f"[WARNING] Port {port} is still in use.")


if __name__ == '__main__':
    # In debug mode Flask runs this file twice: a watcher that restarts the server on code changes, and the
    # server itself (WERKZEUG_RUN_MAIN=true). Only the first may clear the port, or the server would stop its own watcher.
    if os.environ.get('WERKZEUG_RUN_MAIN') != 'true':
        free_port(PORT)
    if getattr(sys, 'frozen', False):
        app.run(debug=False, port=PORT)
    else:
        app.run(debug=True, port=PORT)

    
