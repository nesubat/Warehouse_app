from engine.core import paths
import time
from datetime import datetime
from flask import has_request_context, Flask, render_template, request, redirect, url_for, send_file, abort
import os
import filecmp
import shutil
import json
import pymupdf as fitz
import io
import re
import sys
import stat
import socket
import subprocess
import pandas as pd
from werkzeug.utils import secure_filename
from werkzeug.exceptions import HTTPException, InternalServerError
from urllib.parse import urlencode, urlsplit
from engine.label_shuffler.pdf_engine import process_and_shuffle_pdf
from engine.packing_sheets.matrix_engine import scan_excel_tabs, generate_tab_map, generate_all_outputs, convert_legacy_excel_to_xlsx
from engine.core.core_math import job_series, JOB_NUMBER, unused_name, clean_file_name, close_if_open_elsewhere, save_if_open_elsewhere, clean_store_name, NOT_STORE_TABS, read_divider_barcodes, \
    free_name
from engine.packing_sheets.subgroup_engine import execute_subgroups, SubgroupValidationError
from engine.packing_labels.packing_label_generator import (_address_key, parse_packing_data, generate_packing_labels, PackCheckError, _norm, MAPPABLE_FIELDS,
                                     group_issues, issue_where, issue_text, _row_ranges,
                                     detect_columns, header_columns, similar_addresses, cell_preview)
from openpyxl.utils.cell import column_index_from_string, get_column_letter
from engine.addresses.address_book import AddressBook, AddressBookError
from engine.addresses.address_import import build_review, apply_review, all_label_maps, ImportProblem
from engine.addresses.address_match import suggestions as closest_addresses
from engine.courier.packing_specs import SpecStore, parse_formula, FORMULA_PREFIXES, FORMULA_LABELS, FORMULA_EXAMPLES, ITEM_TYPES
from engine.courier.courier_import import read_allocation, scan_columns, mapping_fields, reference_in_sheets, PACK_FIELD, \
    FIELD_NAMES as IMPORT_FIELD_NAMES
from engine.addresses import address_reference
from engine.core import data_transfer
from engine.addresses.address_book import one_line as address_one_line
from engine.packing_labels.label_stitcher import (read_courier_csv, courier_csv_names, address_changes, plan_stitch, render_stitch, match_session, packing_slots, StitchError, render_label,
                                                 pack_services, csv_kind)
import threading
from collections import Counter
import secrets
from engine.courier.courier_export import (unique_codes, build_consignments, detect_series, write_label_map, one_line, DEFAULT_FIXED,
                            write_courier_csv, open360_item_references, clean_export, load_export_defaults,
                            save_export_defaults, EXPORT_DEFAULTS, open360_items, write_open360_csv,
                            source_destinations, sendable, fill_store_names,
                            EDITABLE_FIELDS, load_service_usage, service_options, record_service_usage,
                            write_csv_recipe, csv_from_recipe, RECIPE_SUFFIX, CSV_KINDS)
from engine.courier.courier_services import ServiceStore, ServiceError, LOGO_MAX_BYTES, LOGO_DIR_NAME, LEGACY_LOGO_DIR_NAME
import openpyxl
from engine.core import app_log, project_record, log_view, progress, env_file
from engine.packing_labels.horizontal_distribution import read_horizontal, read_signature_file
from engine.packing_labels.horizontal_labels import generate_horizontal_labels
import logging





BASE_DIR = paths.BASE_DIR  # the app folder (next to the .exe when built)
env_file.load(os.path.join(BASE_DIR, '.env'))  # this computer's settings (env_file.py), before anything reads them

        

app_log.setup(BASE_DIR)  # terminal + logs/warehouse.log (app_log.py)
LOG_DIR = app_log.log_dir(BASE_DIR)  # read back by the Logs page (log_view.py)
log_pack, log_stitch, log_book = app_log.get('packing'), app_log.get('stitch'), app_log.get('book')
log_specs = app_log.get('specs')
log_import = app_log.get('import')
log_sheets, log_subgroup, log_shuffler = app_log.get('sheets'), app_log.get('subgroup'), app_log.get('shuffler')
log_app = app_log.get('app')  # anything that failed without its own handling (_failed)
# The matching page's "finished yet?" checks and label pictures aren't worth a terminal line each
logging.getLogger('werkzeug').addFilter(lambda r: not re.search(r'"GET /stitch/[^ ]*/(status|label/)', r.getMessage()))

PROJECTS_FOLDER = os.path.join(BASE_DIR, 'projects')
SERVICE_USAGE_FILE = os.path.join(BASE_DIR, 'data', 'service_code_usage.json')
EXPORT_FILE = os.path.join(BASE_DIR, 'data', 'export_details.json')  # OpenFreight export details: saved defaults
ADDRESS_BOOK = AddressBook(os.path.join(BASE_DIR, 'data', 'address_book.db'))
SPEC_STORE = SpecStore(os.path.join(BASE_DIR, 'data', 'packing_specs.json'))
# Courier services and their groups' logos (the logo Stitch Labels puts on each label)
SERVICE_STORE = ServiceStore(os.path.join(BASE_DIR, 'data', 'courier_services.json'), os.path.join(BASE_DIR, 'data', LOGO_DIR_NAME),
                             legacy_dir=os.path.join(BASE_DIR, 'data', LEGACY_LOGO_DIR_NAME))
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

app = Flask(__name__,
            template_folder=os.path.join(BASE_DIR, 'templates'),
            static_folder=os.path.join(BASE_DIR, 'static'))
app.config['UPLOAD_FOLDER'] = PROJECTS_FOLDER
app.config['TEMP_FOLDER'] = temp_dir
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['TEMP_FOLDER'], exist_ok=True)
# Uploads are Excel files and courier label PDFs: anything far bigger is a mistake (or an attack), refused early
app.config['MAX_CONTENT_LENGTH'] = int(os.environ.get('WAREHOUSE_MAX_UPLOAD_MB', '500')) * 1024 * 1024

def _help_url(url):
    """WAREHOUSE_HELP_URL from .env: only a web address (http or https) becomes the nav's Help link."""
    url = str(url or '').strip()
    return url if re.match(r'https?://[^\s"<>]+$', url, re.I) else ''


HELP_URL = _help_url(os.environ.get('WAREHOUSE_HELP_URL'))


@app.context_processor
def nav_settings():
    """What the nav bar on every page needs: the Help link (none when .env doesn't set one)."""
    return {'help_url': HELP_URL}


# Routes that do something (every POST, and the GET links that open a file in Windows). Any web page open in the
# same browser could otherwise make the browser call them on 127.0.0.1 (cross-site request forgery).
ACTION_GETS = {'open_local_file', 'open_upload_file', 'stitch_page'}  # stitch_page: ?run= picks the run


def script_json(value):
    """JSON safe to put inside a <script> tag: '<', '>' and '&' escaped, so a folder or file name holding
    '</script>' can't end the tag and run as code. JSON.parse reads it back unchanged."""
    return json.dumps(value).replace('<', '\\u003c').replace('>', '\\u003e').replace('&', '\\u0026')


@app.before_request
def same_origin_only():
    """Refuses actions started by another website. Browsers say where a request comes from (Sec-Fetch-Site, else
    Origin); this app's own pages are same-origin, typing a URL or a bookmark is 'none'."""
    if request.method != 'POST' and request.endpoint not in ACTION_GETS:
        return None
    site = request.headers.get('Sec-Fetch-Site')
    if site in ('cross-site', 'same-site'):
        return 'Refused: this request came from another website.', 403
    origin = request.headers.get('Origin')
    if not site and origin and origin != 'null' and origin.split('://', 1)[-1] != request.host:
        return 'Refused: this request came from another website.', 403
    return None


# Live progress for the loading overlay (progress.py): a form posted with a progress_id is followed by
# /api/progress/<id> while it runs. Only form posts: a JSON body is left for the route to read.
logging.getLogger(app_log.ROOT).addHandler(progress.Handler())


def _logs_for_error(back=None, error=None):
    """Where a failed action lands: the Logs page opened at what went wrong (the last error this thread logged, or
    `error` = (time, message)), with a link back to the page (`back`: this app's own address only)."""
    error = error or app_log.last_error()
    args = {'at': error[0], 'find': error[1][:120]} if error else {}
    if back:
        parts = urlsplit(back)
        if not parts.netloc or (has_request_context() and parts.netloc == request.host):
            local = parts.path + (f"?{parts.query}" if parts.query else '')
            if local.startswith('/') and not local.startswith('//'):
                args['back'] = local
    return '/logs' + (f"?{urlencode(args)}" if args else '')


def _wants_json():
    """A page's script asked (fetch), not the browser for a page."""
    return (request.path.startswith('/api/') or request.is_json or 'X-Progress-Id' in request.headers
            or 'application/json' in request.headers.get('Accept', ''))


@app.errorhandler(Exception)
def _failed(e):
    """A task behind the loading overlay (it carries a progress id) that failed without its own handling: logged with
    its traceback, then the browser lands on the Logs page at it (a script gets the address in X-Log-Url and goes
    there itself). Anything else keeps the usual error page (it's in the log too); 404s and the like stay as they are."""
    if isinstance(e, HTTPException):
        return e
    log_app.exception("%s %s FAILED: %s", request.method, request.path, e)
    if not progress.current():
        return InternalServerError(original_exception=e)
    url = _logs_for_error(back=request.referrer if request.method == 'POST' else None)
    if _wants_json():
        return {'error': f"That didn't work: {e}. What happened is on the Logs page.", 'log_url': url}, 500, {'X-Log-Url': url}
    return redirect(url, code=303)


@app.before_request
def follow_progress():
    app_log.forget_error()  # each request's own errors (_logs_for_error)
    if request.method != 'POST':
        return
    # A page's script (the matching page's Finalise) sends it in a header; a form, as a field
    run_id = request.headers.get('X-Progress-Id', '')
    if not run_id and request.mimetype in ('multipart/form-data', 'application/x-www-form-urlencoded'):
        run_id = request.form.get('progress_id', '')
    progress.start(run_id)


@app.teardown_request
def end_progress(_error=None):
    progress.finish()


@app.route('/api/progress/<run_id>')
def api_progress(run_id):
    run = progress.get(run_id)
    return run if run is not None else ({'unknown': True}, 404)

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
            log_sheets.info("Upload %s: %d tab(s) %s", filename, len(tabs or []), tabs)

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
    log_sheets.info("Preview %s: tabs %s", filename, selected_tabs)
    started = time.perf_counter()

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
                log_sheets.info("  tab %s (start %s, job %s, stores in %s): %d pack(s), stores to row %s", tab, start_cell,
                                job_id_cell, store_col, len(blueprints[tab].get("pack_ranges") or []),
                                blueprints[tab].get("last_row"))
            except Exception as e:
                log_sheets.warning("  tab %s could not be mapped (start %s, job %s, stores in %s): %s", tab, start_cell,
                                   job_id_cell, store_col, e)
                previews.append({
                    "sheet_name": tab,
                    "error": f"Failed to map. Check your coordinates! ({str(e)})"
                })
                
    blueprints_json = json.dumps(blueprints)
    log_sheets.info("Preview ready in %.2fs", time.perf_counter() - started)

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
        # The job number itself (J476523, not J476523-01 or -001): the folder name the dashboard reads it from
        job_id = clean_file_name(job_series(blueprints[first_tab].get("raw_job_id", "UNKNOWN")) or "UNKNOWN")
        
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
    started = time.perf_counter()
    log_sheets.info("Packing Sheets %s: tabs %s, project '%s', packs %s", filename, selected_tabs, raw_project_name,
                    {t: user_inputs.get(t, {}).get("selected_packs") for t in selected_tabs})
    try:
        file1_name, file2_name, file3_name, shared_job_warnings = generate_all_outputs(
            new_filepath, filename, selected_tabs, user_inputs, blueprints, project_dir
        )
    except Exception:
        log_sheets.exception("Packing Sheets FAILED after %.2fs", time.perf_counter() - started)
        raise
    log_sheets.info("Packing Sheets done in %.2fs: %s", time.perf_counter() - started, final_folder_name)
    raw_files = [file1_name, file2_name, file3_name]

    # Filter the list to ONLY keep actual file names (removes None/Empty strings)
    files_to_download = [f for f in raw_files if f]
    _record(project_dir, 'packing_sheets', files_to_download + [filename])

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


def _project_job(folder, files, project):
    """A project's job number for the dashboard, the number itself (J477130, not J477130-03: job_series)."""
    return job_series(_project_job_text(folder, files, project))


def _project_job_text(folder, files, project):
    """Where the dashboard finds a project's job: the one its record keeps (a Label Shuffler project), else the Packing
    Sheet's ('Packing Sheet_J477130-03.xlsx'), else the Consignment Reference Generate Labels / Courier Import put in
    front of its files' names ('J477130 - Macpac - Open360.csv'), else the label map's own record of it. '' when none."""
    kept = project_record.job(folder)
    if kept:
        return kept
    for f in files:
        if f.startswith("Packing Sheet_"):
            return re.sub(r' \(\d+\)$', '', f[len("Packing Sheet_"):].rsplit(".", 1)[0])
    for f in files:
        for tail in (".labelmap.json", " - Open360.csv", " - OpenFreight.csv"):
            if f.endswith(tail):
                base = re.sub(r' \(\d+\)$', '', f[:-len(tail)])
                if base.endswith(f" - {project}"):
                    return base[:-len(f" - {project}")]
    for f in files:
        if f.endswith(".labelmap.json"):
            try:
                with open(os.path.join(folder, f), encoding='utf-8') as fh:
                    return str(json.load(fh).get('consignment_reference') or '')
            except (OSError, ValueError):
                pass
    return ''


def newest_first(folder, names):
    return sorted(names, key=lambda f: os.path.getmtime(os.path.join(folder, f)), reverse=True)


def _generation_report(job, project, shipment_reference, sheet_warnings, courier_warnings,
                       all_cartons, cartons, consignments, by_number, service, sender, csv_formats=None, ref_issues=None):
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
        {'title': "Don't match the official address data (G-NAF / LINZ), or the street isn't on record - check before dispatch", 'level': 'warn',
         'items': list(ref_issues or [])},
        {'title': "New to the address book - check before dispatch", 'level': 'warn',
         'items': [f"#{c['number']} {receiver(c)} - {one_line(c['destination'])}" for c in new_addresses]},
        {'title': "Courier checks", 'level': 'warn', 'items': [w for w in courier_warnings if not w.startswith(own)]},
        {'title': "Spreadsheet checks", 'level': 'info',
         'items': [", ".join(p for p in (issue_where(w), f"col {w['col']}" if w.get('col') else '') if p)
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


def _label_maps(folder):
    """The project's label maps, newest first: one per Generate run (a Packing Sheets project can have one per Final
    file, Stage files included)."""
    return newest_first(folder, [f for f in os.listdir(folder) if f.endswith('.labelmap.json')]) if os.path.isdir(folder) else []


def _current_run(folder):
    """The run (label map) the Stitch pages are on: the one picked on the Stitch page, else the newest."""
    maps = _label_maps(folder)
    chosen = _read_json(os.path.join(folder, CURRENT_RUN_FILE), {}).get('label_map')
    return chosen if chosen in maps else (maps[0] if maps else None)


def _stitch_project(project_name):
    """(project folder, label map file name, label map) of the run being stitched (_current_run), else None."""
    folder = os.path.join(app.config['UPLOAD_FOLDER'], os.path.basename(project_name))
    run = _current_run(folder)
    if not run:
        return None
    with open(os.path.join(folder, run), encoding='utf-8') as f:
        label_map = json.load(f)
    # The run this stitch is of, kept with it to the end: picking another run meanwhile (another tab) doesn't move it
    label_map['_run'] = run
    return folder, run, label_map


def _run_file(folder, kind, run=None):
    """This run's stitching file: its session ('stitch_session - <label map>.json') or its hand matches
    ('stitch_matches - …'), so each run of a project is stitched on its own. A project stitched before runs were kept
    apart has one 'stitch_session.json': it belongs to the run that was newest then (_keep_earlier_session)."""
    run = run or _current_run(folder)
    legacy = os.path.join(folder, f"stitch_{kind}.json")
    if not run:
        return legacy
    path = os.path.join(folder, f"stitch_{kind} - {run[:-len('.labelmap.json')]}.json")
    if not os.path.exists(path) and os.path.exists(legacy) and run == _label_maps(folder)[0]:
        return legacy
    return path


def _keep_earlier_session(folder, new_map):
    """A new run in a project: the stitching files of before (one per project then) are the earlier run's, so they
    take its name, and the Stitch pages move on to the new run."""
    earlier = [m for m in _label_maps(folder) if m != new_map]
    for kind in ('session', 'matches'):
        legacy = os.path.join(folder, f"stitch_{kind}.json")
        if os.path.exists(legacy):
            if earlier:
                os.replace(legacy, os.path.join(folder, f"stitch_{kind} - {earlier[0][:-len('.labelmap.json')]}.json"))
            else:
                os.remove(legacy)
    if os.path.exists(os.path.join(folder, CURRENT_RUN_FILE)):
        os.remove(os.path.join(folder, CURRENT_RUN_FILE))


def _stitch_file(name):
    """The app's own stitching files, kept out of the project's file lists."""
    return name == CURRENT_RUN_FILE or (name.startswith(('stitch_session', 'stitch_matches')) and name.endswith('.json'))


CURRENT_RUN_FILE = "stitch_run.json"  # the run picked on the Stitch page
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
    progress.stage("Reading the courier labels and matching them", 2, 10)
    manual = _read_json(_run_file(folder, 'matches', label_map.get('_run')), {})
    log_stitch.info("Match %s: %s, %d saved hand match(es)", os.path.basename(folder),
                    ", ".join(f"{n} ({len(d) // 1024} KB)" for n, d in pdfs), len(manual))
    plan = plan_stitch(folder, label_map, pdfs, manual)
    session = match_session(label_map, plan)
    old = _read_json(_run_file(folder, 'session', label_map.get('_run')), {})
    session['added_unmatched'] = old.get('added_unmatched', True) if add_unmatched is None else add_unmatched
    session['output'] = old.get('output') if old.get('output') and os.path.isfile(os.path.join(folder, old['output'])) else None
    _write_json(_run_file(folder, 'session', label_map.get('_run')), session)
    log_stitch.info("  matched in %.2fs: %d placed, %d courier label(s) unmatched, %d packing label(s) waiting%s",
                    time.perf_counter() - started, len(plan['placed']), len(plan['unmatched']), len(plan['missing']),
                    " - complete" if plan['complete'] else '')
    return plan, session


def _remove_old_labels(folder, keep=None, only=None):
    """Deletes the earlier Complete Labels PDFs (a new one replaces them), so the folder holds one per run. only: the
    names of this run's earlier ones, when the project has several runs (another run's are never touched).
    Returns the ones that couldn't be deleted (open in a PDF viewer)."""
    locked = []
    for f in os.listdir(folder):
        if f.startswith("Complete Labels") and f.endswith(".pdf") and f != keep and (only is None or f in only):
            try:
                os.remove(os.path.join(folder, f))
                log_stitch.info("  removed the earlier %s", f)
            except OSError as e:
                log_stitch.warning("  couldn't remove the earlier %s (%s) - open in a PDF viewer?", f, e)
                locked.append(f)
    return locked


def _stitch_logos(folder, label_map):
    """{pack item reference: (service code, its group's logo or None)}: the service each label books, read from the
    courier CSV picked on the Stitch page (_service_csv)."""
    source = _service_csv(folder, label_map)
    return {ref: (code, SERVICE_STORE.logo_for(code) if code else None)
            for ref, code in pack_services(folder, label_map, source).items()}


CSV_UPLOAD_MAX = 20 * 1024 * 1024


def _keep_courier_csv(folder, map_name, label_map, upload):
    """The courier CSV attached on the Stitch page (changed after Generate, e.g. in the courier portal's file): checked
    to be an Open360 or OpenFreight CSV with rows of this job, then kept in the project beside the others under its own
    name (never over the project's copy) and listed in the label map. Returns (saved name, None) or (None, problem)."""
    raw = upload.read(CSV_UPLOAD_MAX + 1)
    if len(raw) > CSV_UPLOAD_MAX:
        return None, f"{upload.filename} is over 20 MB: that isn't a courier CSV of one job."
    name = unused_name(folder, secure_filename(upload.filename) or "courier.csv")
    path = os.path.join(folder, name)
    with open(path, 'wb') as f:  # as attached, byte for byte
        f.write(raw)
    kind, rows = read_courier_csv(path)
    problem = None
    if not kind:
        problem = f"{upload.filename} isn't an Open360 or OpenFreight courier CSV (its header row isn't one)."
    elif not any(pack_services(folder, {**label_map, 'consignments': []}, name).values()):
        problem = (f"None of the rows in {upload.filename} are this job's labels (their Item References don't match): "
                   f"attach the courier CSV of this job.")
    if problem:
        os.remove(path)
        return None, problem
    files = label_map.setdefault('files', {})
    files['courier_csvs'] = list(dict.fromkeys(list(files.get('courier_csvs') or []) + [name]))
    _write_json(os.path.join(folder, map_name), label_map)
    log_stitch.info("Stitch %s: courier CSV attached: %s, kept as %s (%s, %d rows)", os.path.basename(folder),
                    upload.filename, name, CSV_KINDS[kind], len(rows) - 1)
    return name, None


def _run_choices(folder, current):
    """The project's runs for the Stitch page when it has more than one: [{'map', 'pdf', 'when', 'current',
    'stitched'}], newest first ([] for one run)."""
    maps = _label_maps(folder)
    if len(maps) < 2:
        return []
    out = []
    for m in maps:
        lm = _read_json(os.path.join(folder, m), {}) or {}
        files = lm.get('files') or {}
        out.append({'map': m, 'pdf': files.get('packing_labels_pdf') or m, 'current': m == current,
                    'when': datetime.fromtimestamp(os.path.getmtime(os.path.join(folder, m))).strftime('%d/%m/%Y %I:%M %p').replace(' 0', ' '),
                    'stitched': [f for f in files.get('complete_labels') or [] if os.path.isfile(os.path.join(folder, f))]})
    return out


def _courier_csv_choices(folder, label_map):
    """The job's courier CSVs in the project, newest first, for 'which file has the current services' on the Stitch
    page: [{'name', 'kind', 'saved', 'changed'}]. changed: saved after the labels were generated (edited since)."""
    files = label_map.get('files') or {}
    pdf = os.path.join(folder, os.path.basename(files.get('packing_labels_pdf') or ''))
    generated = os.path.getmtime(pdf) if os.path.isfile(pdf) else None
    out = []
    for name in newest_first(folder, [n for n in courier_csv_names(label_map) if os.path.isfile(os.path.join(folder, n))]):
        kind, _ = read_courier_csv(os.path.join(folder, name))
        if kind:
            when = os.path.getmtime(os.path.join(folder, name))
            out.append({'name': name, 'kind': CSV_KINDS[kind],
                        'saved': datetime.fromtimestamp(when).strftime('%d/%m/%Y %I:%M %p').replace(' 0', ' '),
                        'changed': generated is not None and when > generated + 60})
    return out


def _service_csv(folder, label_map):
    """The courier CSV whose services go on the labels: the one picked on the Stitch page (kept in the label map),
    else the newest of the job's."""
    names = [c['name'] for c in _courier_csv_choices(folder, label_map)]
    chosen = (label_map.get('files') or {}).get('service_csv')
    return chosen if chosen in names else (names[0] if names else None)


def _finalise(folder, label_map, pdfs, add_unmatched=True):
    """Stitches once: the plan (codes + saved hand matches) drawn into a new Complete Labels PDF, which replaces the
    earlier ones. Returns (output name, result, earlier files that were open elsewhere and stayed)."""
    plan, session = _plan(folder, label_map, pdfs, add_unmatched)
    project_label = (label_map.get('project') or {}).get('name') or re.sub(r'_\d{6}_\d{4}$', '', os.path.basename(folder))
    output_name = free_name(folder, numbered_name(folder, "Complete Labels", label_map.get('consignment_reference'), project_label,
                                                  datetime.now().strftime("%y%m%d_%H%M")))
    started = time.perf_counter()
    try:
        result = render_stitch(folder, label_map, pdfs, plan, os.path.join(folder, output_name), add_unmatched,
                               logos=_stitch_logos(folder, label_map))
    except Exception:
        log_stitch.exception("Stitch FAILED after %.2fs", time.perf_counter() - started)
        raise
    run = label_map.get('_run') or _current_run(folder)  # the run it started on, not one picked since
    files = label_map.setdefault('files', {})
    mine = set(files.get('complete_labels') or [])
    locked = _remove_old_labels(folder, keep=output_name, only=None if len(_label_maps(folder)) == 1 else mine)
    files['complete_labels'] = [output_name] + [f for f in locked if f in mine]
    if run:
        _write_json(os.path.join(folder, run), label_map)
    source = _service_csv(folder, label_map)
    moved = address_changes(folder, label_map, source) if source else []
    for m in moved:  # the courier label is printed from the CSV, the packing label from Generate
        log_stitch.warning("  #%s %s: the address in %s isn't the one on its packing labels. Labels: %s. CSV: %s",
                           m['number'], ', '.join(m['stores']), source, m['labels'], m['now'])
    result = {**result, 'service_csv': source, 'address_changes': moved}
    _record(folder, 'stitch', [output_name])
    session.update(output=output_name, added_unmatched=add_unmatched, finalised=True)
    _write_json(_run_file(folder, 'session', run), session)
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
    run_id = progress.detach()  # the overlay keeps following the stitch in its thread, not this request

    def work():
        progress.attach(run_id)
        try:
            output_name, result, locked = _finalise(folder, label_map, pdfs, add_unmatched)
            STITCH_JOBS[folder] = {'state': 'done', 'output': output_name, 'placed': len(result['placed']),
                                   'added_unmatched': add_unmatched, 'locked': locked, **extra,
                                   'unmatched': len(result['unmatched']),  # really matched nothing (hand matches aren't)
                                   'waiting': sum(1 for w in result['without_label'] if not w.endswith("(not in the courier CSV)")),
                                   # each packing label still without a courier label, with its page in the new file
                                   'without': [m for m in result.get('missing_pages', []) if not m['not_in_csv']],
                                   'not_in_csv': [m for m in result.get('missing_pages', []) if m['not_in_csv']]}
        except Exception as e:  # the page goes to the Logs page at it; the traceback is in the log (_finalise)
            STITCH_JOBS[folder] = {'state': 'error', 'error': str(e), 'log_url': _logs_for_error()}
        finally:
            progress.finish()
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
    picked = os.path.basename(request.args.get('run', ''))
    if picked:  # a run picked from the list (a project with labels from more than one Final file)
        here = os.path.join(app.config['UPLOAD_FOLDER'], os.path.basename(project_name))
        if (STITCH_JOBS.get(here) or {}).get('state') == 'running':
            return _show('stitch.html', project=os.path.basename(project_name),
                         error="Still stitching the other labels of this project: pick another run when it's done.")
        if picked in _label_maps(here):
            _write_json(os.path.join(here, CURRENT_RUN_FILE), {'label_map': picked})
        return redirect(url_for('stitch_page', project_name=os.path.basename(project_name)))
    found = _stitch_project(project_name)
    if not found:
        return _show('stitch.html', project=project_name, error="This project has no label map, so there's nothing to stitch to.")
    folder, map_name, label_map = found
    stitched = newest_first(folder, [f for f in os.listdir(folder) if f.startswith("Complete Labels") and f.endswith(".pdf")])
    own = {(label_map.get('files') or {}).get('packing_labels_pdf')} | set(stitched)
    candidates = newest_first(folder, [f for f in os.listdir(folder) if f.lower().endswith('.pdf') and f not in own])
    slots = packing_slots(label_map)
    session = _read_json(_run_file(folder, 'session'), None)
    info = {'packs': sum(1 for sl in slots if not sl['not_in_csv']),
            'boxed': sum(1 for sl in slots if sl['boxes'] > 1 and sl['box'] == 1),
            'with_code': sum(1 for pk in (p for ps in label_map.get('stores', {}).values() for p in ps) if pk.get('open360_code')),
            'left_out': len(label_map.get('not_in_courier_csv') or []),
            'latest': stitched[0] if stitched else None, 'stitched': stitched, 'candidates': candidates,
            'matching': bool(session and not session.get('complete') and (session.get('unmatched') or session.get('waiting'))
                             and not session.get('finalised')),
            'csvs': _courier_csv_choices(folder, label_map), 'service_csv': _service_csv(folder, label_map),
            'runs': _run_choices(folder, map_name)}
    if request.method == 'GET':
        return _show('stitch.html', project=os.path.basename(folder), info=info)

    if (STITCH_JOBS.get(folder) or {}).get('state') == 'running':  # one stitch of a project at a time
        return _show('stitch.html', project=os.path.basename(folder), info=info,
                     error="Still stitching this project in the background: stitch again when it's done.")
    csv_upload = request.files.get('courier_csv')
    if csv_upload and csv_upload.filename:
        if not secure_filename(csv_upload.filename).lower().endswith('.csv'):
            return _show('stitch.html', project=os.path.basename(folder), info=info,
                         error=f"{csv_upload.filename} isn't a CSV. Drop the Open360 or OpenFreight CSV of this job.")
        source, problem = _keep_courier_csv(folder, map_name, label_map, csv_upload)
        if problem:
            return _show('stitch.html', project=os.path.basename(folder), info=info, error=problem)
    else:  # the file picked on the page as the one with the current services
        source = os.path.basename(request.form.get('service_csv', ''))
        source = source if source in {c['name'] for c in info['csvs']} else info['service_csv']
    if source and source != (label_map.get('files') or {}).get('service_csv'):
        label_map.setdefault('files', {})['service_csv'] = source  # kept: Finalise on the matching page uses it too
        _write_json(os.path.join(folder, map_name), label_map)
        log_stitch.info("Stitch %s: services and addresses from %s", os.path.basename(folder), source)
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
    session = _read_json(_run_file(folder, 'session'), None)
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
    session = _read_json(_run_file(folder, 'session'), {})
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
    session = _read_json(_run_file(folder, 'session'), {})
    known_keys = {u['key'] for u in session.get('unmatched') or []}
    known_slots = {w['item_reference'] for w in session.get('waiting') or []}
    return {k: v for k, v in (matches or {}).items() if k in known_keys and v in known_slots}, session, known_keys


def _save_matches(folder, matches, known_keys):
    """The page's pairs replace what was saved for its courier labels (a pair undone there is dropped too)."""
    saved = {k: v for k, v in _read_json(_run_file(folder, 'matches'), {}).items() if k not in known_keys}
    saved.update(matches)
    _write_json(_run_file(folder, 'matches'), saved)
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
    session = _read_json(_run_file(folder, 'session'), {})
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


def _sheet_outputs(folder_path, files=None):
    """A Packing Sheets project's outputs for Generate Labels (horizontal) and the Label Shuffler: its Final files
    (the Final file, and the Stage - n - Final files Sub-Group makes), newest first, and its Signature Links file
    (one per project). Returns (finals, signature or None)."""
    files = files if files is not None else os.listdir(folder_path)
    finals = newest_first(folder_path, [f for f in files if f.lower().endswith('.xlsx') and not f.startswith('~$')
                                        and re.match(r'(Stage - \d+ - )?Final ', f)])
    signature = next((f for f in newest_first(folder_path, files)
                      if f.lower().startswith('signature links') and f.lower().endswith(('.xlsx', '.xlsm'))), None)
    return finals, signature


def _keep_used_copy(filepath, project_dir, filename, time_stamp, log):
    """After Generate into a Packing Sheets project: the page's copy of the Final file goes when it's the same as the
    project's own; when it was changed (Open Excel on the page), it's kept in the project as 'Used for labels - <Final
    file> - <time>.xlsx', so the edits aren't lost and the project shows what the labels came from. Returns that name,
    or None. 'Used ...' isn't read as a Final file (_sheet_outputs)."""
    if not os.path.exists(filepath):
        return None
    finals, _ = _sheet_outputs(project_dir)
    original = next((f for f in finals if secure_filename(f) == filename), None)
    if original and filecmp.cmp(filepath, os.path.join(project_dir, original), shallow=False):
        os.remove(filepath)
        return None
    name = free_name(project_dir, f"Used for labels - {os.path.splitext(original or filename)[0]} - {time_stamp}.xlsx")
    shutil.move(filepath, os.path.join(project_dir, name))
    log.info("  the Final file was changed on this page: kept as %s", name)
    return name


def _sheet_project(name):
    """name when it's one of the Packing Sheets projects with a Final file (_sheet_projects), else '': a project name
    sent by a form is only ever one of those, never '.', '..' or another folder, so Generate can't write into (or,
    on a failure, clear) anything else."""
    name = os.path.basename(str(name or '').strip())
    if not name or name.startswith('.'):
        return ''
    return name if any(p['name'] == name for p in _sheet_projects()) else ''


def _sheet_projects():
    """The Packing Sheets projects that have a Final file, newest first, for 'use a project' on Generate Labels:
    [{'name', 'label', 'finals', 'signature'}]."""
    out = []
    root = app.config['UPLOAD_FOLDER']
    for name in os.listdir(root) if os.path.isdir(root) else []:
        path = os.path.join(root, name)
        if not os.path.isdir(path) or ('_Job-' not in name and not any(f.startswith('Packing Sheet_') for f in os.listdir(path))):
            continue
        finals, signature = _sheet_outputs(path)
        if finals:
            project, _, job = re.sub(r'_\d{6}_\d{4}$', '', name).partition('_Job-')
            out.append({'name': name, 'label': f"{job} | {project}" if job else project,  # as on the dashboard
                        'finals': finals, 'signature': signature, 'time': os.path.getmtime(path)})
    return sorted(out, key=lambda p: -p['time'])


def _courier_csvs(folder_path, files):
    """The project's courier CSVs by kind ({'open360': name, 'openfreight': name}, read from their header rows) and
    its newest recipe (write_csv_recipe), or None."""
    import csv as csv_module
    kinds = {}
    for f in files:
        if f.lower().endswith('.csv'):
            kind = None
            # The app writes UTF-8; Excel's "CSV (Comma delimited)" saves in Windows' own code page, so a CSV
            # edited and saved in Excel is still known for what it is
            for encoding in ('utf-8-sig', 'cp1252'):
                try:
                    with open(os.path.join(folder_path, f), encoding=encoding, newline='') as fh:
                        kind = csv_kind(next(csv_module.reader(fh), []))
                    break
                except UnicodeDecodeError:
                    continue
                except OSError:
                    break
            if kind and kind not in kinds:
                kinds[kind] = f
    recipes = newest_first(folder_path, [f for f in files if f.endswith(RECIPE_SUFFIX)])
    recipe = recipes[0] if recipes else None
    # Several runs in a project: the newest run's own CSVs (named in its recipe), so another run's file isn't taken
    # for this run's
    made = (_read_json(os.path.join(folder_path, recipe), {}) or {}).get('made') if recipe else None
    if made:
        kinds = {k: n for k, n in made.items() if k in CSV_KINDS and n in files}
    return kinds, recipe


@app.route('/project/<folder_name>/courier-csv', methods=['POST'])
def make_other_csv(folder_name):
    """The courier CSV kind a project doesn't have yet, made from its recipe in the same folder (a few ms), then back
    to the dashboard with the project open at its courier CSVs and the new file marked."""
    folder = os.path.join(app.config['UPLOAD_FOLDER'], os.path.basename(folder_name))
    kind = request.form.get('kind', '')
    if not os.path.isdir(folder) or kind not in CSV_KINDS:
        abort(404)
    files = os.listdir(folder)
    kinds, recipe = _courier_csvs(folder, files)
    if not recipe:
        return "This project has no courier CSV details saved (it was made before they were kept). Generate it again.", 400
    if kind in kinds:
        return redirect(url_for('dashboard', open=os.path.basename(folder), new=kinds[kind]) + '#dashboard-section')
    # The new CSV is made from what Generate saved. A courier CSV changed since (edited in Excel, or replaced on the
    # Stitch page) has changes the new one won't: say so on the card first, and make it only when asked again
    recipe_time = os.path.getmtime(os.path.join(folder, recipe))
    changed = [f for f in kinds.values() if os.path.getmtime(os.path.join(folder, f)) > recipe_time + 2]
    if changed and request.form.get('confirm') != '1':
        return redirect(url_for('dashboard', open=os.path.basename(folder), stale=changed[0]) + '#dashboard-section')
    started = time.perf_counter()
    name = csv_from_recipe(folder, recipe, kind)
    # The label map lists every courier CSV, so Stitch Labels and the Open360 codes know the new one too
    maps = newest_first(folder, [f for f in files if f.endswith('.labelmap.json')])
    if maps:
        label_map = _read_json(os.path.join(folder, maps[0]), None)
        if label_map:
            lf = label_map.setdefault('files', {})
            lf['courier_csvs'] = list(dict.fromkeys(list(lf.get('courier_csvs') or [lf.get('courier_csv')]) + [name]))
            if kind == 'open360' and not lf.get('open360_csv'):
                lf['open360_csv'] = name
            _write_json(os.path.join(folder, maps[0]), label_map)
    _record(folder, 'courier_csv', [name], kind=kind)
    log_pack.info("%s CSV made from %s in %s: %s (%.0f ms)", CSV_KINDS[kind], recipe, os.path.basename(folder), name,
                  (time.perf_counter() - started) * 1000)
    return redirect(url_for('dashboard', open=os.path.basename(folder), new=name) + '#dashboard-section')


def _features_from_files(folder_name, files):
    """What a project from before records were kept was made with, read from its files' names."""
    pdfs = [f for f in files if f.lower().endswith('.pdf')]
    csvs = [f for f in files if f.lower().endswith('.csv')]
    has_map = any(f.endswith('.labelmap.json') for f in files)
    found = set()
    if '_Job-' in folder_name or any(f.startswith('Packing Sheet_') for f in files):
        found.add('packing_sheets')
    if any('_Shuffled' in f for f in pdfs):
        found.add('shuffler')
    if has_map or any(f.startswith(('Packing Labels', 'Complete Labels')) for f in pdfs):
        found.add('labels')
    if csvs and not pdfs and not has_map:
        found.add('courier_import')
    return found


def _record(folder, feature, files=(), **details):
    """project_record.record, starting a project's first record from what its files show (an older project)."""
    project_record.record(folder, feature, files, seed=lambda: _features_from_files(os.path.basename(folder), os.listdir(folder)),
                          **details)


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
                json_files = sorted([f for f in files if f.endswith('.json') and not f.endswith(('.labelmap.json', RECIPE_SUFFIX))
                                     and not _stitch_file(f) and f != project_record.RECORD_FILE], key=lambda x: os.path.getctime(os.path.join(folder_path, x)), reverse=True )

                
                excel_files = [f for f in display_files if f.lower().endswith(('.xlsx', '.xls'))]
                pdf_files = [f for f in display_files if f.lower().endswith('.pdf')]
                csv_files = [f for f in display_files if f.lower().endswith('.csv')]
                
                # Get human-readable date
                timestamp = os.path.getctime(folder_path)
                date_str = datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M')

                # The project name the user gave (the folder's name without the job and the time it was made), and
                # its job number: shown as 'J477130 | Macpac'
                display_name = re.sub(r'_\d{6}_\d{4}$', '', folder_name.split('_Job-')[0] if '_Job-' in folder_name else folder_name)
                job_id = _project_job(folder_path, files, display_name)
                        
                has_label_map = any(f.endswith('.labelmap.json') for f in files)
                csv_kinds, csv_recipe = _courier_csvs(folder_path, files) if csv_files else ({}, None)

                # What the project was made with: its record (project_record.json, written by each tool), else, for a
                # project from before records were kept, read from its files. It sets the card's colours (the same as
                # the tools' tiles on Home, sharing the band) and which buttons it offers.
                features = project_record.features(folder_path)
                if features is None:
                    features = _features_from_files(folder_name, files)
                kinds = [k for f, k in (('packing_sheets', 'sheets'), ('shuffler', 'shuffler'), ('labels', 'labels'),
                                        ('courier_import', 'import')) if f in features] or (['labels'] if pdf_files else ['sheets'])
                courier_only = 'courier_import' in features and 'labels' not in features and 'packing_sheets' not in features

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
                    # Only one kind of courier CSV and the details to make the other: offer it
                    "csv_missing": next((k for k in CSV_KINDS if k not in csv_kinds), None)
                                   if csv_recipe and len(csv_kinds) == 1 else None,
                    "json_files": json_files,    # Pass the grouped JSON files
                    # Packing Sheets projects: Generate Labels from a Final file, the Label Shuffler with its Signature Links
                    **dict(zip(('final_files', 'signature_file'),
                               _sheet_outputs(folder_path, files) if 'sheets' in kinds else ([], None))),
                })
                
    # Sort projects newest to oldest
    projects.sort(key=lambda x: x['timestamp'], reverse=True)
    
    return render_template('index.html', projects=projects, csv_kinds=CSV_KINDS,
                           open_project=request.args.get('open', ''), new_file=request.args.get('new', ''),
                           stale_csv=request.args.get('stale', ''))

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
        json_files = [f for f in os.listdir(project_dir) if f.endswith('.json') and not f.endswith(('.labelmap.json', RECIPE_SUFFIX))
                      and not _stitch_file(f) and f != project_record.RECORD_FILE]
        target_json = json_files[0] if json_files else f"{safe_project}.json"
        log_subgroup.debug("Sub-Group %s: no details file chosen, using %s", safe_project, target_json)
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
                
        started = time.perf_counter()
        log_subgroup.info("Sub-Group %s from %s: %s", safe_project, target_json,
                          "; ".join(f"tab {t} (item row {i['item_row']}): " +
                                    ", ".join(f"pack {p} items {' '.join(f'{a}-{b}' for a, b in r)}" for p, r in i['packs'].items())
                                    for t, i in subgroup_instructions.items()) or "nothing chosen")

        # 3. Trigger the Engine
        try:
            before = set(os.listdir(project_dir))
            execute_subgroups(project_dir, metadata, subgroup_instructions)
            made = sorted(set(os.listdir(project_dir)) - before)
            _record(project_dir, 'sub_group', made,
                                  tabs=sorted(subgroup_instructions) if isinstance(subgroup_instructions, dict) else None)
            log_subgroup.info("Sub-Group done in %.2fs: %s", time.perf_counter() - started, ", ".join(made) or "no new files")
        except SubgroupValidationError as e:
            log_subgroup.warning("Sub-Group stopped: %s", e)
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
            job_number = re.sub(r'\s+', '', request.form.get('job_number', '')).upper()
            if new_project and not JOB_NUMBER.match(job_number):
                # The dashboard shows each project's job number: a new project has nothing else to read it from
                return _show('pdf.html', **_shuffler_step1(), new_project=new_project, job_number=job_number,
                             error="Type the job number of this project, e.g. J476523.")
            time_stamp = datetime.now().strftime("%y%m%d_%H%M")
            safe_project_name = clean_file_name(new_project)
            final_folder_name = f"{safe_project_name}_{time_stamp}"

            
            project_name = final_folder_name.strip() if new_project and new_project.strip() else existing_project
            
            if not project_name:
                return "Please select or enter a Project Name.", 400
            
            # --- THE FIX: Point directly to the final Project Folder ---
            project_folder = os.path.join(PROJECTS_FOLDER, os.path.basename(project_name))
            os.makedirs(project_folder, exist_ok=True)
            if new_project:
                project_record.set_job(project_folder, job_series(job_number))
                log_shuffler.info("New project %s, job %s", os.path.basename(project_folder), job_series(job_number))

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
                        if sheet in NOT_STORE_TABS:
                            continue  # divider barcodes / the Label Plugin: not a tab of store codes
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

                projects_json = script_json(projects_info)

                log_shuffler.warning("Signature Links %s: duplicate store names: %s", os.path.basename(excel_path),
                                     "; ".join(map(str, duplicate_errors)))
                # Re-render Step 1 with duplicate errors, plus enough context for the
                # "Open Excel File" / "Recheck File" buttons to target the exact file.
                return _show('pdf.html',
                                       step=1,
                                       existing_projects=list(projects_info.keys()),
                                       projects_json=projects_json,
                                       duplicate_errors=duplicate_errors,
                                       duplicate_project_name=os.path.basename(project_folder),
                                       duplicate_excel_filename=os.path.basename(excel_path))
                
            log_shuffler.info("Signature Links %s for %s: %s", os.path.basename(excel_path), project_name,
                              {t: len(p) for t, p in tabs_data.items()})
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
            # Only the PDFs actually attached; each is a step of the overlay's bar (progress.py)
            uploads = [(k, f) for k, f in request.files.items() if k.startswith('pdf---') and f and f.filename]
            started = time.perf_counter()
            log_shuffler.info("Shuffle %s: %d PDF(s) with %s", os.path.basename(project_name), len(uploads),
                              os.path.basename(excel_path or ''))
            for done, (key, file) in enumerate(uploads):
                progress.stage(f"Shuffling {file.filename} ({done + 1} of {len(uploads)})",
                               5 + 90 * done / len(uploads), 5 + 90 * (done + 1) / len(uploads))
                if key.count('---') == 2:  # pdf---<tab>---<pack>
                    parts = key.split('---')
                    
                    if len(parts) == 3:
                        tab_name = parts[1]
                        pack_name = parts[2]
                        
                        df = all_sheets_data.get(tab_name)
                        
                        # Failsafe: Prevent crashes if the tab name is completely invalid
                        if df is None:
                            log_shuffler.warning("  %s: no tab '%s' in the Signature Links file: skipped", file.filename, tab_name)
                            continue 
                        
                        # 1. Build the mapping for THIS specific tab and pack
                        store_mapping = {}
                        
                        # Failsafe: Prevent KeyErrors if the pack name isn't found
                        if pack_name not in df.columns:
                            log_shuffler.warning("  %s: no pack '%s' in tab '%s' of the Signature Links file: skipped",
                                                 file.filename, pack_name, tab_name)
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
                        log_shuffler.info("  %s (tab %s, pack %s%s) -> %s", file.filename, tab_name, pack_name,
                                          ", divider sheets" if add_dividers_flag else "", final_filename)

            try:
                shutil.rmtree(temp_dir)
            except Exception:
                pass

            if generated_files:
                _record(project_folder, 'shuffler', generated_files)
            log_shuffler.info("Shuffle done in %.2fs: %d file(s) in %s", time.perf_counter() - started,
                              len(generated_files), os.path.basename(project_folder))
            return _show('pdf.html', step=3, project_name=project_name, generated_files=generated_files)

    chosen = os.path.basename(request.args.get('project', ''))  # from the dashboard's Label Shuffler button
    step1 = _shuffler_step1()
    return _show('pdf.html', **step1, chosen=chosen if chosen in step1['existing_projects'] else '')


def _shuffler_step1():
    """Step 1 of the Label Shuffler: the projects to pick from, with each one's Signature Links file."""
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
                
    return {'step': 1, 'existing_projects': list(projects_info.keys()), 'projects_json': script_json(projects_info)}


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


def _signature_path(filepath):
    """Where the Signature Links file uploaded with a horizontal distribution file is kept until Generate."""
    return filepath + '.signature.xlsx'


def _signature_for(filepath):
    """The Signature Links file uploaded with this distribution file (read_signature_file), or None."""
    path = _signature_path(filepath)
    if not os.path.isfile(path):
        return None
    try:
        return read_signature_file(path)
    except Exception as e:
        log_pack.warning("  Signature Links file %s couldn't be read: %s", os.path.basename(path), e)
        return None


def _read_tab(filepath, header_row, tab, columns=None, mode='labels', tab_no=1):
    """parse_packing_data() (Courier Import: read_allocation(); horizontal files: read_horizontal()), plus store
    names for files whose only address info is one combined cell. columns: the tab's Column mapping set by hand in
    the preview (_column_choices). tab_no: the tab's place among the selected tabs (T1, T2 … in references)."""
    if mode == 'horizontal':
        pack_groups, last_row, found_headers, warnings = read_horizontal(filepath, header_row, tab, columns, tab_no,
                                                                         _signature_for(filepath))
        return fill_store_names(pack_groups), last_row, found_headers, warnings
    read = read_allocation if mode == 'courier' else parse_packing_data
    pack_groups, last_row, found_headers, warnings = read(filepath, header_row, sheet_name=tab, columns=columns)
    return fill_store_names(pack_groups), last_row, found_headers, warnings


def _column_choices(form, tab, mode='labels'):
    """The preview's Column mapping for a tab: {field: column letter, or '-' for not used}; automatic ones left out.
    Sent as 'colmap::<tab>::<field>' (the page also puts back what was chosen for this file last time). Courier
    Import's fields are the address list's, and 'pack_<n>' for its pack columns."""
    choices = {}
    prefix = f'colmap::{tab}::'
    if mode in ('courier', 'horizontal'):
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
    if mode in ('courier', 'horizontal'):
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


def _receiver_groups(consignments, separate=()):
    """Consignments for the same receiver (a store, or an installer) at different addresses, put next to each other
    so they can be merged into one in the preview. Names count as the same when they're equal ignoring capitals,
    spacing and punctuation, or one small typo apart. Returns the consignments in the new order, each in a group
    carrying c['group'] = {'id', 'first', 'size', 'receiver', 'similar', 'members'}; 'similar' says whether every
    address looks like the same place written differently (similar_addresses) or some really differ.
    separate: the consignment sources the user chose to keep apart (Keep separate): a group of only those isn't one."""
    from difflib import SequenceMatcher
    separate = set(separate)

    def typo_apart(a, b):
        # One character changed, added or dropped (two in a long name): 'Macpac Chadstone' / 'Macpac Chadston'. Not
        # two different places that happen to look alike ('Macpac Perth' / 'Macpac Penrith' is two letters apart)
        changed = sum(max(i2 - i1, j2 - j1) for op, i1, i2, j1, j2 in SequenceMatcher(None, a, b).get_opcodes() if op != 'equal')
        return min(len(a), len(b)) > 5 and changed <= (2 if max(len(a), len(b)) >= 20 else 1)

    keys = []  # one per receiver name, in order of first appearance: (name, state of its first address)
    group_of = {}
    for c in consignments:
        name = _norm(c['destination']['receiver'])
        if not name:
            continue
        state = (c['destination'].get('state') or '').upper()
        # A name spelt slightly differently is the same receiver only in the same state (or where one has no address)
        key = next((k for k, st in keys if k == name or (typo_apart(k, name) and (not st or not state or st == state))), None)
        if key is None:
            keys.append((name, state))
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
        if len(members) > 1 and all(x['source_id'] in separate for m in members for x in m['cartons']):
            # kept apart on purpose: still next to each other, without the merge heading
            for m in members:
                m['kept_separate'] = True
        elif len(members) > 1:
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
    items = open360_items(cartons, consignments)
    export, _ = clean_export(_export_values(form))
    csv_files, made = [], {}
    wanted = [k for k in ('open360', 'openfreight') if k in csv_formats]
    # The overlay's bar (progress.py): each CSV a step of 20 % … 90 %, then the details for the other CSV
    step = lambda kind, n: progress.stage(f"Writing the {CSV_KINDS[kind]} CSV ({len(cartons)} carton(s), "
                                          f"{len(consignments)} consignment(s))", 20 + 70 * n / len(wanted), 20 + 70 * (n + 1) / len(wanted))
    if 'open360' in csv_formats:
        step('open360', wanted.index('open360'))
        name = unused_name(project_dir, f"{base_name} - Open360.csv")
        with app_log.step(log, f"  Open360 CSV {name}"):
            write_open360_csv(os.path.join(project_dir, name), items, reference, item_refs=refs)
        log.info("    %d row(s), Shipment Reference %s", len(refs), reference)
        csv_files.append(name)
        made['open360'] = name
    if 'openfreight' in csv_formats:
        step('openfreight', wanted.index('openfreight'))
        name = unused_name(project_dir, f"{base_name} - OpenFreight.csv")
        with app_log.step(log, f"  OpenFreight CSV {name}"):
            write_courier_csv(os.path.join(project_dir, name), cartons, consignments, reference, fixed, export=export)
            abroad = sum(1 for c in consignments if (c['destination'].get('country') or 'AU').upper() != 'AU')
            if abroad:
                log.info("    export details on %d consignment(s) outside Australia: %s", abroad, export)
        log.info("    %d row(s), Reference %s", len(cartons), reference)
        csv_files.append(name)
        made['openfreight'] = name
    # What both CSVs are made from, so the other one can be made later from the dashboard
    progress.stage("Saving the details for the other courier CSV", 90, 99)
    recipe = unused_name(project_dir, base_name + RECIPE_SUFFIX)
    write_csv_recipe(os.path.join(project_dir, recipe), base_name, reference,
                     reference, fixed, export, cartons, consignments, items, refs, None, made)
    log.info("  details for the other courier CSV: %s", recipe)
    return csv_files


def _carton_weights(form):
    """Carton weights typed in the preview's consignment table: {pack key: kg}, from a hidden JSON field."""
    try:
        data = json.loads(form.get('carton_weights') or '{}')
    except ValueError:
        return {}
    return {str(k): v for k, v in data.items()} if isinstance(data, dict) else {}


def _same_address(a, b):
    return all(str(a.get(k) or '').strip() == str(b.get(k) or '').strip() for k in EDITABLE_FIELDS) \
        and bool(a.get('authority_to_leave')) == bool(b.get('authority_to_leave'))


def _book_options(pack_groups, edits):
    """What the address book knows for each Excel address, offered in the preview (nothing is filled in by itself:
    the preview shows the Excel address, and the user picks an option with one click). {source id: [(entry, why)]}:
      - the address the book learned for this exact Excel spelling (a correction made before)
      - an address with no name in it: who the book has at that address
      - a receiver with no address in the file: the receiver's saved addresses (up to 3)
      - a receiver whose Excel address isn't what the book has for them: the receiver's saved addresses (up to 2,
        verified first), the same name or a near-identical spelling, so a store that moved or was spelt differently
        shows its saved address with what differs marked
    Excel addresses the user has already edited in this preview get none, nor do ones kept as is ('book_keep')."""
    sources = source_destinations(pack_groups)
    ask = [sid for sid in sources if not any(k in edits.get(sid, {}) for k in EDITABLE_FIELDS)
           and not (isinstance(edits.get(sid, {}).get('book_keep'), dict) and _same_place(edits[sid]['book_keep'], sources[sid]))]
    options = {}
    try:
        learned = ADDRESS_BOOK.lookup_aliases(ask)
        for sid in ask:
            dest = sources[sid]
            if sid in learned:
                options[sid] = [(learned[sid], "Saved for this Excel spelling")]
            elif dest['postcode'] and not dest['receiver']:
                here = ADDRESS_BOOK.at_address(dest)
                if here:
                    options[sid] = [({**dest, **{k: here[0][k] for k in ('receiver', 'contact')}}, "Saved at this address")]
            elif dest['postcode'] and dest['receiver']:
                # The receiver is saved, at another address (or this one spelt differently): offer theirs, unless
                # the book already holds this exact address for them
                saved = ADDRESS_BOOK.receiver_entries(dest['receiver'])
                if saved and not any(_same_place(e, dest) for e in saved):
                    seen, found = set(), []
                    for e in saved:
                        place = _address_key(one_line(e))
                        if place not in seen:
                            seen.add(place)
                            found.append((e, "Verified for this receiver" if e['verified_at'] else "Saved for this receiver"))
                    options[sid] = found[:2]
            elif not dest['postcode'] and dest['receiver']:
                seen, found = set(), []
                for e in ADDRESS_BOOK.find_by_receiver(dest['receiver']):  # most used first
                    place = _address_key(one_line(e))
                    if place not in seen:
                        seen.add(place)
                        found.append((e, "Saved for this receiver"))
                if found:
                    options[sid] = found[:3]
    except Exception as e:
        log_book.exception("Address book lookup failed: %s", e)
    log_book.info("Address book: %d Excel addresses, %d with an address-book option", len(sources), len(options))
    return options


app.add_template_filter(group_issues, 'group_issues')

# Courier checks shown in the To check list in their own way: per packing spec with a link, and per row
SPEC_WARNINGS = ("No size saved for packing spec", "Couldn't read the size from")
# Courier checks that mean the courier can't take a consignment as it is: listed with the major problems
MAJOR_COURIER_WARNINGS = ("No store / receiver name and no address", "No receiver name for", "Consignment ")


def _unfixed_preview_problems(consignments, refs):
    """The consignment preview's problems still there at Generate, as file problems for the PACK AUDIT pages:
    [(pack key, issue)], one per pack of the consignment, on that pack's Excel rows and store. The ones the preview
    ranks first: no address, a suburb / state / postcode missing, no receiver name, the official data disagreeing or
    a street not on record (not for an address saved in the book). refs: {consignment number: _reference_for(...)}."""
    out = []
    for c in consignments:
        d = c['destination']
        problems = []
        if c.get('no_address'):
            problems.append("No address: left out of the courier CSV")
        else:
            au = (d.get('country') or 'AU').upper() == 'AU'
            missing = [k for k in (('suburb', 'state', 'postcode') if au else ('suburb', 'postcode')) if not d.get(k)]
            if missing:
                problems.append("Address incomplete: " + ", ".join(f"no {k}" for k in missing))
            r = refs.get(c['number']) or {}
            if (r.get('level') == 'mismatch' or (r.get('level') == 'note' and not r.get('soft'))) \
                    and not (missing and str(r.get('message', '')).startswith('No ')):  # said already, just above
                problems.append(f"Official address data: {r.get('message', '')}")
        if not d.get('receiver'):
            problems.append("No receiver name: the courier needs one")
        for text in problems:
            for x in c['cartons']:
                rows = sorted(int(r) for r in x.get('excel_rows') or [] if str(r).isdigit())
                out.append((x['pack_key'], {
                    'rows': _row_ranges(rows) if rows else '', 'col': f"Consignment #{c['number']}",
                    'text': f"{text} (not fixed in the preview)", 'detail': '', 'major': True,
                    'stores': [[r, x['store']] for r in rows] if x.get('store') else []}))
    return out


def _review(previews, cartons, courier_warnings):
    """What the To check list shows besides the rows' own problems (script.js section 25 adds those):
      file     each tab's problems, grouped (group_issues), with the tab's name
      specs    packing specs with no size saved (or a size that can't be read), the consignments they're in and a
               link to put it in on the Packing Specs page: their dimensions are blank in the courier CSV
      courier  the other courier checks, major or not"""
    file = [{**g, 'tab': bp['sheet_name']} for bp in previews for g in group_issues(bp.get('warnings'))]
    specs = {}
    for x in cartons:
        if x['size_cm'] is None:
            s = specs.setdefault(x['packing_spec'], {'spec': x['packing_spec'], 'cartons': 0, 'consignments': [],
                                                     'unread': x['spec_kind'] in ('OB', 'CS', 'PALLET', 'FP')})
            s['cartons'] += 1
            if x['consignment'] not in s['consignments']:
                s['consignments'].append(x['consignment'])
    for s in specs.values():
        # A code the size couldn't be read from goes to Check a code; any other spec to its own row
        s['url'] = (url_for('packing_specs_page', code=s['spec']) + '#ps-try' if s['unread']
                    else url_for('packing_specs_page', spec=s['spec']))
    # The consignments a check names ('#4 …', 'Consignment 5 …') come with it, so the list can go to them. An
    # incomplete address is left to the rows' own 'missing' check (it follows edits)
    courier = []
    for w in courier_warnings:
        if w.startswith(SPEC_WARNINGS + ("No address for ",)) or re.match(r"Consignment \d+ \(.*\): incomplete address", w):
            continue
        numbers = list(dict.fromkeys(int(n) for n in re.findall(r"#(\d+)", w) + re.findall(r"^Consignment (\d+)", w)))
        courier.append({'text': w, 'major': w.startswith(MAJOR_COURIER_WARNINGS), 'numbers': numbers})
    return {'file': file, 'specs': list(specs.values()), 'courier': courier}


def _option_view(entry, why, current=None):
    """An address-book option as the preview shows it, with the fields one click puts in. 'parts': the address as
    [(text, changed)] against the current one, so the page can mark what differs ('Shop 2 *Woolworths* Curlewis Town
    Centre … VIC *3222*'): words of the street lines the current address doesn't have, whole fields otherwise."""
    fields = {k: str(entry.get(k) or '') for k in EDITABLE_FIELDS}
    fields['country'] = fields['country'] or 'AU'
    fields['authority_to_leave'] = bool(entry.get('authority_to_leave'))
    label = one_line(fields) + (f", {fields['country']}" if fields['country'] != 'AU' else '')
    parts, changed = [], []
    current = current or {}
    have_words = set(_norm(" ".join(str(current.get(k) or '') for k in ('line1', 'line2'))).split())
    order = ['line1', 'line2', 'suburb', 'state', 'postcode'] + (['country'] if fields['country'] != 'AU' else [])
    for k in order:
        value = fields[k]
        if not value:
            continue
        if parts:
            parts.append((', ', False))
        if k in ('line1', 'line2'):
            words = value.split(' ')
            for i, w in enumerate(words):
                diff = bool(current) and bool(_norm(w)) and _norm(w) not in have_words
                parts.append((w + (' ' if i < len(words) - 1 else ''), diff))
                if diff and k not in changed:
                    changed.append(k)
        else:
            diff = bool(current) and _norm(value) != _norm(current.get(k) or ('AU' if k == 'country' else ''))
            parts.append((value, diff))
            if diff:
                changed.append(k)
    return {'label': label, 'receiver': fields['receiver'], 'contact': fields['contact'], 'why': why,
            'verified': bool(entry.get('verified_at')), 'fields': fields, 'parts': parts, 'changed': changed}


LOCALITY_FIELD_NAMES = {'suburb': 'Suburb', 'state': 'State', 'postcode': 'Postcode', 'country': 'Country'}


def _locality_hints(dest, ref):
    """What the saved addresses say about this address's suburb, field by field: the state, postcode and country
    other saved (verified first) addresses in that suburb have, and the suburb other addresses with this postcode
    have. Offered only for a field that's empty, or that the official data can't check (New Zealand postcodes and
    regions, or no official data loaded); a field the official data checks is left to it. Each hint changes only its
    own field. [{'field', 'value', 'label', 'why', 'verified'}]"""
    country = (dest.get('country') or 'AU').upper()
    official = address_reference.available() and ref.get('level') not in ('none', None)
    checked_by_official = {'suburb', 'state', 'postcode'} if (official and country == 'AU') else ({'suburb'} if official else set())
    hints = []
    try:
        suburb, postcode = (dest.get('suburb') or '').strip(), (dest.get('postcode') or '').strip()
        if suburb:
            rows = [r for r in ADDRESS_BOOK.localities(suburb) if (r['country'] or 'AU') == country] or ADDRESS_BOOK.localities(suburb)
            best = rows[0] if rows else None
            # Supported well enough: verified, or saved more than once, and not contradicted by an equal alternative
            if best and (best['verified'] or best['entries'] > 1) and not (len(rows) > 1 and rows[1]['verified'] == best['verified']
                                                                          and rows[1]['entries'] == best['entries']):
                for field in ('state', 'postcode', 'country'):
                    value = (best[field] or ('AU' if field == 'country' else '')).strip()
                    have = (dest.get(field) or ('AU' if field == 'country' else '')).strip()
                    if not value or _norm(value) == _norm(have):
                        continue
                    if have and field in checked_by_official:
                        continue
                    if field == 'country' and official and ref.get('level') in ('ok', 'note', 'verified'):
                        continue  # the official data recognised it in this country
                    hints.append({'field': field, 'value': value,
                                  'label': f"{LOCALITY_FIELD_NAMES[field]} {value}",
                                  'why': f"{suburb.upper()} is {value} in {best['verified'] or best['entries']} "
                                         f"{'verified' if best['verified'] else 'saved'} address{'es' if (best['verified'] or best['entries']) != 1 else ''}"
                                         + (f" (this says {have})" if have else ''),
                                  'verified': bool(best['verified'])})
        elif postcode:
            rows = ADDRESS_BOOK.suburbs_at(postcode, country)
            if rows and (rows[0]['verified'] or rows[0]['entries'] > 1) and len({r['suburb'] for r in rows[:2]}) == 1:
                r = rows[0]
                hints.append({'field': 'suburb', 'value': r['suburb'], 'label': f"Suburb {r['suburb']}",
                              'why': f"{postcode} is {r['suburb']} in your {'verified' if r['verified'] else 'saved'} addresses",
                              'verified': bool(r['verified'])})
    except Exception as e:
        log_book.warning("Saved-address suburb lookup failed for %s: %s", one_line(dest), e)
    return hints





SUGGESTION_FIELDS = EDITABLE_FIELDS + ('authority_to_leave',)


def _closest(address, cache=None):
    """Closest saved addresses for the ✏️ form, best first, each with the fields a click fills in."""
    return [{**s, 'json': json.dumps({k: s[k] for k in SUGGESTION_FIELDS})}
            for s in closest_addresses(address, ADDRESS_BOOK.candidates(address, cache))]


def _same_place(a, b):
    """Two addresses are the same place as the address book compares them (capitals, spacing, St = Street …)."""
    pick = lambda x: {k: str((x or {}).get(k) or '') for k in ('line1', 'line2', 'suburb', 'state', 'postcode')}
    return _address_key(one_line(pick(a))) == _address_key(one_line(pick(b)))


def _reference_for(dest, edit=None, here=None):
    """What the official address data says about a consignment address, after the address book has had its say:
      verified  an address-book entry here was verified by the courier portal: nothing more is checked
      kept      the user chose to keep this exact address as is (in this preview, or before in the book)
      ok / note / mismatch / none   address_reference.assess(), with 'soft' when the address is saved in the book
                (not verified): the official data's view is then only a note, the book's address may well be right"""
    if not address_reference.available() or not any(dest.get(k) for k in ('postcode', 'suburb', 'line1', 'line2')):
        return {'level': 'none', 'message': '', 'fields': {}, 'options': []}
    try:
        here = ADDRESS_BOOK.at_address(dest) if here is None else here
    except Exception as e:
        log_book.warning("Address book read failed while checking %s: %s", one_line(dest), e)
        here = []
    if any(e.get('verified_at') for e in here):
        return {'level': 'verified', 'message': 'Verified by the courier portal', 'fields': {}, 'options': []}
    keep = (edit or {}).get('ref_keep')
    if any(e['ref']['status'] == 'kept' for e in here) or (isinstance(keep, dict) and _same_place(keep, dest)):
        return {'level': 'kept', 'message': 'Kept as is', 'fields': {}, 'options': []}
    try:
        r = address_reference.assess(dest)
    except Exception as e:
        log_book.exception("Official address check failed for %s: %s", one_line(dest), e)
        return {'level': 'none', 'message': '', 'fields': {}, 'options': []}
    # A suburb, state or postcode missing is never soft, saved in the book or not: the courier can't take it
    missing = any(str(r['fields'].get(k, {}).get('text', '')).startswith('No ') for k in ('suburb', 'state', 'postcode'))
    return {'level': r['level'], 'message': r['message'], 'fields': r['fields'], 'options': r['options'],
            'country': r.get('country', ''), 'soft': bool(here) and not missing}


def _check_against_book(consignments, edits, book_options=None):
    """Sets each consignment's 'book' status for the preview:
      book       filled in from the address book (a spelling it learned)
      matched    the book holds this address, under this receiver or another one (an address can have several
                 receivers / Attn names). A receiver or Attn new to it gets 'book_note': it's added on Generate.
      checked    not in the book, but the user has reviewed it in the edit form
      unverified not in the book: listed first, with 'suggestions' (closest saved addresses, best first)
      missing    no address at all (only a receiver name): listed first of all, with the receiver's saved
                 addresses as suggestions. Left out of the courier CSV until it has one."""
    cache = {}
    book_options = book_options or {}
    for c in consignments:
        d = c['destination']
        c['book'], c['book_note'], c['suggestions'] = 'unverified', '', []
        # The address book's options for this consignment's Excel address(es), unless they're what it already says
        offered = []
        for sid in dict.fromkeys([c['id']] + [x['source_id'] for x in c['cartons']]):
            for entry, why in book_options.get(sid, []):
                view = _option_view(entry, why, d)
                same = _same_place(view['fields'], d) and _norm(view['receiver']) == _norm(d['receiver'])
                if not same and view['label'] not in [o['label'] for o in offered]:
                    offered.append(view)
        c['book_options'] = offered[:3]
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
    # The official address data: book entries first (a verified one settles it), then the reference
    started = time.perf_counter()
    for c in consignments:
        c['ref'] = ({'level': 'none', 'message': '', 'fields': {}, 'options': []} if c['book'] == 'missing'
                    else _reference_for(c['destination'], edits.get(c['id'], {})))
        # What the saved addresses say about its suburb: only fields left empty or the official data can't check
        # (a row with no postcode included: its suburb may be all it needs to be completed)
        c['locality_hints'] = [] if c['ref']['level'] == 'verified' else _locality_hints(c['destination'], c['ref'])
    if address_reference.available():
        levels = Counter(c['ref']['level'] + ('-saved' if c['ref'].get('soft') else '') for c in consignments)
        log_book.info("  official address data: %s in %.2fs", ", ".join(f"{k} {v}" for k, v in sorted(levels.items())),
                      time.perf_counter() - started)
    # What would make a consignment go wrong first: no address (left out of the CSV), then the official data saying
    # a suburb, state, postcode, country or street doesn't exist there; then the address book's options and hints;
    # then addresses that are only new to the address book (they're saved on Generate); then the rest
    def rank(c):
        if c['book'] == 'missing':
            return 0
        if c['ref']['level'] in ('mismatch', 'note') and not c['ref'].get('soft'):
            return 1
        if c['ref']['level'] == 'mismatch' or c['book_options'] or c['locality_hints']:
            return 2
        return 3 if c['book'] == 'unverified' else 4
    return sorted(consignments, key=rank)


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


# --- BACK UP OR MOVE THE APP'S DATA (data_transfer.py) ---
DATA_DIR = os.path.join(BASE_DIR, 'data')
TRANSFER_UPLOADS = os.path.join(DATA_DIR, 'transfer_uploads')


# ---------- Logs: what the app is doing and has done, from its own log file (log_view.py) ----------
# Read-only GETs: nothing here changes anything.

@app.route('/logs')
def logs_page():
    return render_template('logs.html')


@app.route('/api/logs')
def api_logs():
    a = request.args
    if a.get('since') and a.get('since') == log_view.version(LOG_DIR):
        return {'unchanged': True}  # the page asks every few seconds: nothing new, nothing read
    before, limit = a.get('before', ''), a.get('limit', '')
    return log_view.view(LOG_DIR, area=a.get('area', ''), show=a.get('show', 'all'), q=a.get('q', '')[:200],
                         day=a.get('day', 'today'), limit=min(int(limit), 500) if limit.isdigit() else 150,
                         before=int(before) if before.isdigit() else None,
                         focus_at=a.get('at') or None, focus_text=a.get('find', '')[:200])


@app.route('/data-transfer')
def data_transfer_page():
    """Export any of the address book, packing specs and settings as one zip, or import such a zip."""
    kept = _kept_view()  # an import's next step or result, reopened on refresh
    if kept:
        return kept
    return render_template('data_transfer.html', here=data_transfer.what_is_here(DATA_DIR, ADDRESS_BOOK, SERVICE_STORE))


@app.route('/data-transfer/export', methods=['POST'])
def data_transfer_export():
    chosen = request.form.getlist('item')
    try:
        data, manifest = data_transfer.export_zip(DATA_DIR, ADDRESS_BOOK, chosen, SERVICE_STORE)
    except data_transfer.TransferError as e:
        return render_template('data_transfer.html', here=data_transfer.what_is_here(DATA_DIR, ADDRESS_BOOK, SERVICE_STORE), error=str(e))
    log_book.info("Data export: %s", ", ".join(manifest['items']))
    return send_file(io.BytesIO(data), as_attachment=True, mimetype='application/zip',
                     download_name=f"warehouse-data-{datetime.now():%Y%m%d-%H%M}.zip")


@app.route('/data-transfer/import', methods=['POST'])
def data_transfer_import():
    """Step 1: an exported zip is uploaded and its contents listed. Step 2 ('apply'): the parts chosen are taken."""
    here = data_transfer.what_is_here(DATA_DIR, ADDRESS_BOOK, SERVICE_STORE)
    os.makedirs(TRANSFER_UPLOADS, exist_ok=True)
    if 'apply' not in request.form:
        upload = request.files.get('file')
        if not upload or not upload.filename:
            return _show('data_transfer.html', _at=url_for('data_transfer_page'), here=here, error="Choose the exported zip (warehouse-data-….zip).")
        data = upload.read()
        try:
            manifest, items = data_transfer.describe_zip(data)
        except data_transfer.TransferError as e:
            return _show('data_transfer.html', _at=url_for('data_transfer_page'), here=here, error=str(e))
        token = secrets.token_urlsafe(10)
        with open(os.path.join(TRANSFER_UPLOADS, f"{token}.zip"), 'wb') as f:
            f.write(data)
        # Uploads not imported within a day go
        for name in os.listdir(TRANSFER_UPLOADS):
            path = os.path.join(TRANSFER_UPLOADS, name)
            if time.time() - os.path.getmtime(path) > 86400:
                os.remove(path)
        return _show('data_transfer.html', _at=url_for('data_transfer_page'), here=here, incoming=items, token=token, filename=upload.filename,
                     exported_at=manifest.get('exported_at', '').replace('T', ' '))
    token = request.form.get('token', '')
    path = os.path.join(TRANSFER_UPLOADS, f"{os.path.basename(token)}.zip")
    if not re.fullmatch(r'[\w-]{8,40}', token) or not os.path.exists(path):
        return _show('data_transfer.html', _at=url_for('data_transfer_page'), here=here, error="That upload is no longer available: choose the zip again.")
    # Each chosen item's mode; an item with only one way to import (settings: replace) needs no choice
    choices = {k: request.form.get(f'mode_{k}') or data_transfer.ITEMS[k][3][0]
               for k in request.form.getlist('item') if k in data_transfer.ITEMS}
    try:
        with open(path, 'rb') as f:
            result = data_transfer.import_zip(f.read(), choices, DATA_DIR, ADDRESS_BOOK, SPEC_STORE, SERVICE_STORE)
    except data_transfer.TransferError as e:
        return _show('data_transfer.html', _at=url_for('data_transfer_page'), here=here, error=str(e))
    except Exception as e:
        log_book.exception("Data import failed")
        return _show('data_transfer.html', _at=url_for('data_transfer_page'), here=here, error=f"The import failed: {e}. Nothing after the failure was changed; "
                                                           f"a copy of the data from before is in data/backups.")
    os.remove(path)
    log_book.info("Data import: %s (backup %s)", result['done'], result['backup'])
    recheck_book_async('import')  # imported addresses checked against the official data
    return _show('data_transfer.html', _at=url_for('data_transfer_page'), here=data_transfer.what_is_here(DATA_DIR, ADDRESS_BOOK, SERVICE_STORE), done=result['done'],
                 backup=os.path.relpath(result['backup'], BASE_DIR))


@app.route('/address-data')
def address_data_page():
    """Official address data (G-NAF for Australia, LINZ for New Zealand): what's loaded, updates, a lookup tester."""
    return render_template('address_data.html', au=address_reference.loaded_info('AU'), nz=address_reference.loaded_info('NZ'),
                           due=address_reference.reminder(), settings=address_reference.load_settings(),
                           job=address_reference.job_status(), remind_days=address_reference.REMIND_DAYS,
                           linz_layer=address_reference.linz_layer_info())


@app.route('/api/address-data/key', methods=['POST'])
def api_address_data_key():
    """Saves the LINZ API key (kept in data/address_reference.json, on this computer only)."""
    key = str((request.get_json(silent=True) or {}).get('linz_key', '')).strip()
    address_reference.save_settings({'linz_key': key})
    log_book.info("Address data: LINZ API key %s", "saved" if key else "removed")
    return {'saved': bool(key)}


@app.route('/api/address-data/layer', methods=['POST'])
def api_address_data_layer():
    """Sets NZ Addresses' LINZ layer number by hand, for when LINZ republishes it under a new one; blank goes back
    to looking it up in LINZ's catalogue."""
    layer = str((request.get_json(silent=True) or {}).get('linz_layer', '')).strip()
    if layer and not (layer.isdigit() and len(layer) <= 9):
        return {'error': "A layer number is digits only, e.g. 123113 (from the layer's page address on data.linz.govt.nz)."}, 400
    address_reference.save_settings({'linz_layer': layer})
    log_book.info("Address data: NZ Addresses layer %s", f"set by hand to {layer}" if layer else "back to automatic")
    return {'layer': int(layer) if layer else None}


@app.route('/api/address-data/updates')
def api_address_data_updates():
    """What the sources have now against what's loaded (only when the user asks)."""
    return address_reference.check_updates()


@app.route('/api/address-data/load', methods=['POST'])
def api_address_data_load():
    """Downloads and loads a country's data in the background ({'country': 'AU' or 'NZ'})."""
    country = str((request.get_json(silent=True) or {}).get('country', '')).upper()
    if country not in ('AU', 'NZ'):
        return {'error': 'Choose AU or NZ.'}, 400
    address_reference.close_all()
    error = address_reference.start(country, log_book, on_done=lambda: recheck_book_async(f'{country} data updated'))
    log_book.info("Address data: load %s %s", country, error or "started")
    return ({'error': error}, 409) if error else {'started': country}


@app.route('/api/address-data/status')
def api_address_data_status():
    job = address_reference.job_status()
    if job.get('state') in ('done', 'failed'):
        job['au'], job['nz'] = address_reference.loaded_info('AU'), address_reference.loaded_info('NZ')
    return job


@app.route('/api/address-data/lookup')
def api_address_data_lookup():
    """Checks one address against the official data: suburb + postcode + state, street, number."""
    return address_reference.check_address({k: request.args.get(k, '') for k in
                                            ('line1', 'line2', 'suburb', 'state', 'postcode', 'country')})


@app.route('/api/address-data/check-book', methods=['POST'])
def api_address_data_check_book():
    """Every address-book entry checked against the official data: counts per result and each entry's result."""
    started = time.perf_counter()
    rows, offset = [], 0
    while True:
        page, more = ADDRESS_BOOK.search('', 200, offset)
        rows += page
        offset += len(page)
        if not more or not page:
            break
    results = []
    for e in rows:
        r = address_reference.check_address(e)
        ADDRESS_BOOK.set_reference(e['id'], address_reference.assess(e))  # the book's own check, kept up to date
        results.append({'id': e['id'], 'receiver': e['receiver'], 'address': address_one_line(e),
                        'status': r['status'], 'country': r['country'], 'ms': r['ms'],
                        'checks': r['checks'], 'suggestions': r['suggestions']})
    counts = Counter(r['status'] for r in results)
    took = time.perf_counter() - started
    log_book.info("Address data: checked %d address-book entries in %.2fs: %s", len(results), took, dict(counts))
    return {'results': results, 'counts': counts, 'seconds': round(took, 2),
            'average_ms': round(sum(r['ms'] for r in results) / len(results), 1) if results else 0}


RECHECK = {'running': False, 'done': 0, 'reason': ''}
_RECHECK_LOCK = threading.Lock()


def recheck_book_async(reason):
    """Checks, in the background, the address-book entries not yet checked against the current official data (or
    changed since). Runs once after each update of the official data (every entry), and after a Generate or an import
    (only the entries those just saved). Never on a timer, never just because a page was opened."""
    if not address_reference.available():
        return
    with _RECHECK_LOCK:
        if RECHECK['running']:
            return
        RECHECK.update(running=True, done=0, reason=reason)
    book = ADDRESS_BOOK

    def work():
        started, seen = time.perf_counter(), set()
        try:
            version = address_reference.data_version()
            while True:
                batch = [e for e in book.stale_for_reference(version, 500) if e['id'] not in seen]
                if not batch:
                    break
                for e in batch:
                    seen.add(e['id'])
                    book.set_reference(e['id'], address_reference.assess(e))
                    RECHECK['done'] += 1
            if RECHECK['done']:
                log_book.info("Address book checked against the official data (%s): %d entries in %.1fs, %d need a look",
                              reason, RECHECK['done'], time.perf_counter() - started, book.needs_look_count())
        except Exception as e:
            log_book.exception("Address book check against the official data failed: %s", e)
        finally:
            RECHECK['running'] = False
    threading.Thread(target=work, daemon=True).start()


def _check_entry(row):
    """An entry just added or changed: checked against the official data straight away."""
    if address_reference.available() and isinstance(row, dict) and row.get('id'):
        try:
            ADDRESS_BOOK.set_reference(row['id'], address_reference.assess(row))
            return ADDRESS_BOOK.get(row['id'])
        except Exception as e:
            log_book.exception("Official address check failed for #%s: %s", row.get('id'), e)
    return row


@app.route('/api/addresses')
def api_addresses():
    query, offset = request.args.get('q', ''), request.args.get('offset', 0, type=int)
    rows, has_more = ADDRESS_BOOK.search(query, request.args.get('limit', 50, type=int), offset,
                                         needs_look=request.args.get('needs') == '1')
    # Nothing is checked from here: the book is checked once after each update of the official data
    return {'rows': rows, 'has_more': has_more, 'needs_look': ADDRESS_BOOK.needs_look_count(),
            'checking': RECHECK['running'], 'reference': address_reference.available()}


@app.route('/api/addresses/<int:address_id>/keep', methods=['POST'])
def api_address_keep(address_id):
    """Keep this address as is: the official data's mismatch stops being flagged until the address changes."""
    if not ADDRESS_BOOK.keep_as_is(address_id):
        return {'error': 'That address no longer exists.'}, 404
    log_book.info("Address #%s kept as is despite the official data", address_id)
    return ADDRESS_BOOK.get(address_id)


@app.route('/api/address-data/assess', methods=['POST'])
def api_address_data_assess():
    """One consignment address checked again after an edit in the preview ({'address', 'keep'})."""
    data = request.get_json(silent=True) or {}
    address = {k: str((data.get('address') or {}).get(k) or '') for k in ('line1', 'line2', 'suburb', 'state', 'postcode', 'country')}
    return _reference_for(address, {'ref_keep': data.get('keep')})


@app.route('/api/address-data/complete')
def api_address_data_complete():
    """Suggestions for one address field as it's typed: saved addresses the courier verified first, then other saved
    addresses (both as whole addresses), then places from the saved addresses, then the official data for that field.
    In the suburb and postcode fields a whole address is offered only when it's this receiver's own (receiver= the
    name in the form): anyone else's would put their street in; places fill just the suburb, state and postcode."""
    from engine.addresses.address_match import name_score
    field = request.args.get('field', '')
    value = request.args.get('value', '').strip()
    receiver = request.args.get('receiver', '').strip()
    ctx = {k: request.args.get(k, '') for k in ('line1', 'line2', 'suburb', 'state', 'postcode', 'country')}
    out = []
    if field in ('line1', 'line2', 'suburb', 'postcode') and len(value) >= 2:
        try:
            rows = ADDRESS_BOOK.search(value, 30)[0]
        except Exception:
            rows = []
        typed = _norm(value)
        rows = [r for r in rows if typed in _norm(r.get(field)) or _norm(r.get(field)).startswith(typed)]
        if field in ('suburb', 'postcode'):
            rows = [r for r in rows if receiver and name_score(receiver, r['receiver']) >= 0.85]
        seen = set()
        for verified in (True, False):
            n = 0
            for r in rows:
                key = _address_key(r['address'])
                if bool(r.get('verified_at')) != verified or key in seen or n >= 3:
                    continue
                seen.add(key)
                n += 1
                out.append({'source': 'verified' if verified else 'saved', 'label': r['address'] + (f", {r['country']}" if r['country'] != 'AU' else ''),
                            'detail': f"saved for {r['receiver']}" if r['receiver'] else 'saved address',
                            'fields': {k: r[k] for k in ('line1', 'line2', 'suburb', 'state', 'postcode', 'country')}})
    # Places from the saved addresses (verified first): a suburb with its state, postcode and country, or the suburbs
    # of a postcode. They come before the official data, and are the only source for NZ postcodes and regions.
    try:
        places = []
        if field == 'suburb':
            places = ADDRESS_BOOK.localities_starting(value)
        elif field == 'postcode' and re.fullmatch(r'\d{4}', value):
            places = [{**r, 'postcode': value} for r in ADDRESS_BOOK.suburbs_at(value)]
        shown = {o['label'] for o in out}
        for p in places[:4]:
            label = " ".join(x for x in (p['suburb'], p['state'], p['postcode']) if x) + (f", {p['country']}" if p['country'] != 'AU' else '')
            if label in shown:
                continue
            shown.add(label)
            n = p['verified'] or p['entries']
            out.append({'source': 'verified' if p['verified'] else 'saved', 'label': label,
                        'detail': f"in {n} {'verified' if p['verified'] else 'saved'} address{'es' if n != 1 else ''}",
                        'fields': {'suburb': p['suburb'], 'state': p['state'], 'postcode': p['postcode'], 'country': p['country'] or 'AU'}})
    except Exception as e:
        log_book.warning("Saved-place suggestions failed for %s=%r: %s", field, value, e)
    try:
        out += [{**o, 'source': 'official'} for o in address_reference.complete(field, value, ctx)]
    except Exception as e:
        log_book.warning("Official suggestions failed for %s=%r: %s", field, value, e)
    return {'options': out[:10]}


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
        return _check_entry(row), 201
    except AddressBookError as e:
        log_book.warning("Address not added: %s", e)
        return {'error': str(e)}, 400


@app.route('/api/addresses/<int:address_id>', methods=['PUT'])
def api_address_update(address_id):
    try:
        row = ADDRESS_BOOK.update(address_id, request.get_json(silent=True) or {})
        log_book.info("Address #%s updated: %s", address_id, row.get('receiver') if isinstance(row, dict) else row)
        return _check_entry(row)
    except AddressBookError as e:
        log_book.warning("Address #%s not updated: %s", address_id, e)
        return {'error': str(e)}, 400


@app.route('/api/addresses/delete', methods=['POST'])
def api_addresses_delete():
    """Deletes the addresses picked on the Address Book page ({ids: [...]}), all in one go."""
    try:
        ids = sorted({int(i) for i in (request.get_json(silent=True) or {}).get('ids') or []})
    except (TypeError, ValueError):
        return {'error': 'Those addresses could not be read. Refresh the page and try again.'}, 400
    if not ids:
        return {'error': 'Pick the addresses to delete first.'}, 400
    n = ADDRESS_BOOK.delete_many(ids)
    log_book.info("Deleted %d address(es) at once: %s", n, ", ".join(f"#{i}" for i in ids[:50]) + (" …" if len(ids) > 50 else ''))
    return {'deleted': n}


@app.route('/api/addresses/<int:address_id>', methods=['DELETE'])
def api_address_delete(address_id):
    if ADDRESS_BOOK.delete(address_id):
        log_book.info("Address #%s deleted", address_id)
        return '', 204
    return {'error': 'That address no longer exists.'}, 404


# ---------- courier services (Address Book page, and "+ Add service" in the consignment preview) ----------
# Every change is a POST, so same_origin_only guards it.

def _services_payload():
    groups = SERVICE_STORE.load()['groups']
    for g in groups:
        g['logo_url'] = url_for('courier_logo', group_id=g['id'], v=os.path.getmtime(p)) if (p := SERVICE_STORE.logo_path(g)) else ''
    return {'groups': groups}


def _service_change(fn, what):
    try:
        result = fn()
        log_book.info("Courier services: %s", what)
        return {'ok': True, 'result': result, **_services_payload()}
    except ServiceError as e:
        log_book.warning("Courier services: %s not saved: %s", what, e)
        return {'error': str(e)}, 400


@app.route('/api/courier-services')
def api_courier_services():
    return _services_payload()


@app.route('/api/courier-services', methods=['POST'])
def api_courier_service_add():
    d = request.get_json(silent=True) or {}
    # Added in a group's card: a code already in another group moves into this one (the preview's + Add service
    # doesn't move anything: a code that's there already is an error)
    return _service_change(lambda: SERVICE_STORE.add_service(d.get('code'), d.get('name'), d.get('group') or None,
                                                             d.get('new_group') or '', move=bool(d.get('move'))),
                           f"added {d.get('code')}")


@app.route('/api/courier-services/<code>/update', methods=['POST'])
def api_courier_service_update(code):
    d = request.get_json(silent=True) or {}
    return _service_change(lambda: SERVICE_STORE.update_service(code, d.get('code'), d.get('name'), d.get('group') or None,
                                                                d.get('new_group') or ''), f"updated {code}")


@app.route('/api/courier-services/<code>/delete', methods=['POST'])
def api_courier_service_delete(code):
    return _service_change(lambda: SERVICE_STORE.delete_service(code), f"deleted {code}")


@app.route('/api/courier-groups', methods=['POST'])
def api_courier_group_add():
    name = (request.get_json(silent=True) or {}).get('name')
    return _service_change(lambda: SERVICE_STORE.add_group(name), f"group {name!r} added")


@app.route('/api/courier-groups/<group_id>/rename', methods=['POST'])
def api_courier_group_rename(group_id):
    name = (request.get_json(silent=True) or {}).get('name')
    return _service_change(lambda: SERVICE_STORE.rename_group(group_id, name), f"group {group_id} renamed {name!r}")


@app.route('/api/courier-groups/<group_id>/delete', methods=['POST'])
def api_courier_group_delete(group_id):
    return _service_change(lambda: SERVICE_STORE.delete_group(group_id), f"group {group_id} deleted")


@app.route('/api/courier-groups/<group_id>/logo', methods=['POST'])
def api_courier_group_logo(group_id):
    f = request.files.get('logo')
    if not f or not f.filename or os.path.splitext(secure_filename(f.filename))[1].lower() not in ('.png', '.jpg', '.jpeg', '.webp'):
        return {'error': 'Drop a PNG or JPG logo.'}, 400
    raw = f.read(LOGO_MAX_BYTES + 1)
    return _service_change(lambda: SERVICE_STORE.set_logo(group_id, raw), f"logo for {group_id}")


@app.route('/api/courier-groups/<group_id>/logo/delete', methods=['POST'])
def api_courier_group_logo_delete(group_id):
    return _service_change(lambda: SERVICE_STORE.delete_logo(group_id), f"logo for {group_id} removed")


@app.route('/courier-logo/<group_id>.png')
def courier_logo(group_id):
    group = next((g for g in SERVICE_STORE.load()['groups'] if g['id'] == group_id), None)
    path = SERVICE_STORE.logo_path(group)
    if not path:
        abort(404)
    return send_file(path, mimetype='image/png', max_age=0)


@app.route('/packing-labels', methods=['GET', 'POST'])
def create_packing_labels():
    """Packing labels (PDF) and courier CSVs from a packing list (a Packing Spec column, rows merged per box)."""
    return _distribution_page('labels')


@app.route('/packing-labels/horizontal', methods=['GET', 'POST'])
def create_horizontal_labels():
    """Packing labels from a horizontal distribution file (a row per store, a column per item, packs under merged
    headers), in packing-sheet order, optionally with the Signature Links file of its Packing Sheets run."""
    return _distribution_page('horizontal')


@app.route('/courier-import', methods=['GET', 'POST'])
def courier_import():
    """Courier Import: only the courier CSV, from an address list with a column per pack (courier_import.py)."""
    return _distribution_page('courier')


def _signature_note(signature):
    """What the preview says about the Signature Links file of a horizontal job."""
    if not signature:
        return {'loaded': False}
    return {'loaded': True, 'plugin': sorted(signature.get('plugin') or {}), 'codes': sorted(signature.get('codes') or {}),
            'barcodes': len(signature.get('barcodes') or {})}


def _keep_signature_upload(upload, filepath, tabs):
    """Keeps the Signature Links file sent with a horizontal distribution file next to it (an older one is
    removed, so a re-scan without one doesn't use it). upload: the uploaded file, or a project's file path. With a Label Plugin tab, the tabs it describes come
    ticked with their header rows filled in. Returns (user_inputs or None, error text or None)."""
    path = _signature_path(filepath)
    if os.path.exists(path):
        os.remove(path)
    if isinstance(upload, str):  # a project's own Signature Links file: copied, the project keeps its own
        shutil.copy2(upload, path)
        shown = os.path.basename(upload)
    elif not upload or not upload.filename:
        return None, None
    elif os.path.splitext(upload.filename)[1].lower() not in ('.xlsx', '.xlsm'):
        return None, "The Signature Links file must be an .xlsx file (the one Packing Sheets made)."
    else:
        upload.save(path)
        shown = upload.filename
    try:
        signature = read_signature_file(path)
    except Exception as e:
        os.remove(path)
        return None, f"Could not read the Signature Links file: {e}"
    plugin = signature.get('plugin') or {}
    log_pack.info("  Signature Links %s: codes for %d tab(s), Label Plugin for %s", secure_filename(shown),
                  len(signature['codes']), sorted(plugin) or 'none')
    if not plugin:
        return None, None
    inputs = {}
    for tab in tabs:
        t = plugin.get(tab)
        above = (t['first_row'] - 1 if t and t.get('first_row') else
                 max([r for r in (t.get('pack_group_row'), t.get('job_id_row')) if r] or [1]) if t else 1)
        inputs[tab] = {'selected': bool(t), 'header_row': str(above)}
    return inputs, None


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
    horizontal_mode = mode == 'horizontal'
    log = log_import if courier_mode else log_pack
    # Started from a Packing Sheets project: every step keeps it, and Generate saves into it
    into = _sheet_project(request.form.get('from_project')) if horizontal_mode and request.method == 'POST' else ''
    into_label = (lambda p, _, j: f"{j} | {p}" if j else p)(*re.sub(r'_\d{6}_\d{4}$', '', into).partition('_Job-')) if into else ''

    def show(**context):
        if context.get('page_error'):  # what the run did before it stopped, and where: shown under the message
            context.setdefault('run_log', progress.lines_now())
            context.setdefault('run_failure', next((f for f in (log_view.stopped_at(x.get('more')) for x in context['run_log']) if f), None))
        context.setdefault('from_project', into)
        context.setdefault('from_project_label', into_label if context['from_project'] == into else '')
        return _show('packing_labels.html', mode=mode, **context)
    if request.method == 'GET':
        return show(sheet_projects=_sheet_projects() if horizontal_mode else [])

    # STEP 1b: a Packing Sheets project's Final file and its Signature Links (the dashboard's Generate Labels button,
    # or 'Use a Packing Sheets project' on this page): copied in as the upload, so the project itself stays as it is
    if (horizontal_mode and request.form.get('from_project') and not request.form.get('filename')
            and not (request.files.get('file') and request.files['file'].filename)):
        project = _sheet_project(request.form['from_project'])
        if not project:
            return show(page_error="That Packing Sheets project isn't on the dashboard any more: pick one from the list.",
                        sheet_projects=_sheet_projects())
        folder = os.path.join(app.config['UPLOAD_FOLDER'], project)
        finals, signature = _sheet_outputs(folder)
        final = os.path.basename(request.form.get('final_file', '')) or (finals[0] if len(finals) == 1 else '')
        if final not in finals:
            return show(page_error="Choose the project's Final file to make the labels from.", sheet_projects=_sheet_projects())
        filename = secure_filename(final)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        try:
            close_if_open_elsewhere(filepath)
            shutil.copy2(os.path.join(folder, final), filepath)
            wb = openpyxl.load_workbook(filepath, read_only=True)
            try:
                tabs = wb.sheetnames
            finally:
                wb.close()
            log.info("From project %s: %s (%d tab(s)), Signature Links %s", project, final, len(tabs), signature or 'none')
            inputs, sig_error = _keep_signature_upload(os.path.join(folder, signature) if signature else None, filepath, tabs)
            if sig_error:
                return show(page_error=sig_error, sheet_projects=_sheet_projects())
            return show(tabs=tabs, filename=filename, user_inputs=inputs, from_project=project)
        except Exception as e:
            log.exception("Could not load %s from project %s", final, project)
            return show(page_error=f"Could not read '{final}' from the project: {e}", sheet_projects=_sheet_projects())

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
            inputs = None
            if horizontal_mode:
                inputs, sig_error = _keep_signature_upload(request.files.get('signature_file'), filepath, tabs)
                if sig_error:
                    return show(page_error=sig_error)
            return show(tabs=tabs, filename=filename, user_inputs=inputs)
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
        first_tab_groups = None  # the first readable tab's packs: its first pack is the live cell preview
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
                pack_groups, last_row, found_headers, warnings = _read_tab(filepath, header_row, tab, choices, mode,
                                                                           selected_tabs.index(tab) + 1)
                
                total_packs = len(pack_groups)
                total_stores = len(set(g['store_name'] for g in pack_groups.values()))
                log.info("  tab %s (header row %d): %d packs, %d stores, last row %s, %d warning(s)",
                              tab, header_row, total_packs, total_stores, last_row, len(warnings))
                for w in warnings:  # each on its own line, where it is and what it is, for the Logs page
                    (log.warning if w.get('major') else log.info)("    needs a look: %s", issue_text(w))
                if first_tab_groups is None and pack_groups:
                    first_tab_groups = pack_groups

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
                log.warning("  tab %s: %d problem(s) stop it from generating; fix them in the Excel file, then Update Previews",
                            tab, len(e.issues))
                for i in e.issues:  # every one, where it is: the Logs page lists them under the preview
                    log.warning("    %s", issue_text(i))
                for w in e.warnings:
                    log.info("    needs a look: %s", issue_text(w))
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
                groups, _, _, _ = _read_tab(filepath, header_rows[tab], tab, _column_choices(request.form, tab, mode), mode,
                                            selected_tabs.index(tab) + 1)
                combined.update({f"{tab} - {k}": v for k, v in groups.items()})
                if courier_mode:
                    _import_refs(groups, selected_tabs.index(tab) + 1)
            # The Excel address as written; what the address book knows is offered in the row, never filled in by itself
            edits = _consignment_edits(request.form)
            edits.pop('__offered__', None)
            book_options = _book_options(combined, edits)
            services = SERVICE_STORE.services()
            options = service_options(load_service_usage(SERVICE_USAGE_FILE, {x['code'] for x in services}), services)
            most_used = options[0]['code'] if options and options[0]['used'] else ''
            fixed = {**DEFAULT_FIXED, 'service_code': most_used,
                     **{k: request.form[k] for k in ('who_pays', 'charge_account', 'service_code', 'reference') if k in request.form}}
            weights = _carton_weights(request.form)
            consignments, cartons, courier_warnings = build_consignments(combined, edits, fixed['service_code'],
                                                                         SPEC_STORE.resolver(), weights)
            if courier_mode:
                _import_item_refs(cartons, combined)
            kept_apart = {sid for sid, e in edits.items() if isinstance(e, dict) and e.get('separate')}
            consignments = _receiver_groups(_check_against_book(consignments, edits, book_options), kept_apart)  # one receiver's next to each other
            # A store kept at its different addresses on purpose: the file's warning about its addresses goes too
            apart = {_norm(c['destination']['receiver']) for c in consignments if c.get('kept_separate')}
            if apart:
                for bp in previews:
                    bp['warnings'] = [w for w in bp.get('warnings') or []
                                      if not any(_norm(w.get('text', '')).startswith(f"{n} store address") or
                                                 _norm(w.get('text', '')).startswith(f"{n} installer addresses") for n in apart)]
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
                'ref_mismatch': sum(c['ref']['level'] == 'mismatch' and not c['ref'].get('soft') for c in consignments),
                'ref_loaded': address_reference.available(),
                'missing': sum(c['book'] == 'missing' for c in consignments),
                'warnings': courier_warnings,
                'review': _review(previews, cartons, courier_warnings),
                'fixed': {'reference': series, **fixed},
                'edits_json': json.dumps(edits),
                'book_option_rows': sum(1 for c in consignments if c['book_options']),
                # Export details for deliveries outside Australia (OpenFreight): as typed so far, else the defaults
                'export': _export_values(request.form),
                'international': sum(1 for c in consignments if (c['destination'].get('country') or 'AU').upper() != 'AU'),
                'service_options': options,
                'service_groups': [{'id': g['id'], 'name': g['name']} for g in SERVICE_STORE.load()['groups']],
                # Horizontal labels: grouped by allocation or not, dividers or not, as chosen (kept through a preview)
                'label_grouping': 'store' if request.form.get('label_grouping') == 'store' else 'code',
                'code_dividers': request.form.get('code_dividers', '1') == '1',
            }

        preview_cells = None
        if not courier_mode and first_tab_groups:
            try:
                preview_cells = cell_preview(first_tab_groups)
            except Exception:
                log.exception("  live cell preview could not be built")

        signature = _signature_for(filepath) if horizontal_mode else None
        return show(tabs=all_tabs, filename=filename, user_inputs=user_inputs, previews=previews, courier=courier,
                               file_note=file_note, cell_preview=preview_cells,
                               signature_note=_signature_note(signature) if horizontal_mode else None)

    # Action: Generate Final PDF & Save to Project
    if 'generate' in request.form:
        combined_pack_groups = {}
        
        # Grab the sorted layout order from the frontend, and filter out 'empty' slots
        attribute_order_str = request.form.get('attribute_order', 'thumbnail,desc,dimension,job_no,barcode,qty')
        attribute_order = [attr for attr in attribute_order_str.split(',') if attr and attr != 'empty']
        
        missing = [label for key, label in (('who_pays', 'Who Pays'), ('service_code', 'Service Code'), ('reference', 'Consignment Reference'),
                                            ('project_name', 'Project Name'))
                   if not request.form.get(key, '').strip() and not (key == 'project_name' and into)]  # into: its own project
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
        project_had = None  # into a Packing Sheets project: the files it had, so a failure takes back only this run's
        started = time.perf_counter()
        tab_of = {}  # pack key -> its tab's number among the selected tabs, for the Item Reference ('01T2QWXK Store')
        pack_notes = {}  # horizontal: (tab number, pack number) -> the warnings about that pack, for its audit page
        tab_notes = {}   # vertical: tab number -> its warnings, for the tab's audit page
        try:
            # 1. Parse Data
            sheet_warnings = []
            progress.stage("Reading the Excel file", 2, 15)  # the loading overlay's bar (progress.py)
            for tab in selected_tabs:
                progress.tick(selected_tabs.index(tab), len(selected_tabs), f"Reading tab {tab}")
                header_row = int(user_inputs[tab]['header_row'])
                pack_groups, _, _, tab_warnings = _read_tab(filepath, header_row, tab, _column_choices(request.form, tab, mode), mode,
                                                            selected_tabs.index(tab) + 1)
                if horizontal_mode:
                    for g in pack_groups.values():
                        key = (g['tab_no'], g['pack_no'])
                        pack_notes.setdefault(key, [w for w in tab_warnings
                                                    if f"pack {g['pack_name']}".lower() in f"{w.get('col', '')} {w.get('text', '')}".lower()])
                if courier_mode:
                    _import_refs(pack_groups, selected_tabs.index(tab) + 1)
                tab_notes[selected_tabs.index(tab) + 1] = list(tab_warnings)
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
            if horizontal_mode:  # '03T1P2': the store's number in its tab, T + the tab, P + the pack
                serials = {key: g['serial'] for key, g in combined_pack_groups.items()}

            # 2. Courier consignments, checked before anything is written or moved
            progress.stage("Checking the consignments and addresses", 15, 20)
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
            unknown = sorted({c['service_code_used'] for c in consignments} - SERVICE_STORE.codes())
            if unknown:
                raise ValueError(f"Unknown service code {', '.join(repr(u) for u in unknown)}. Pick a service from the list.")

            # 3. Setup Project Folder structure
            raw_project_name = _project_name(request.form.get('project_name'), 'Courier_Import_Job' if courier_mode else
                                             'Horizontal_Labels_Job' if horizontal_mode else 'Packing_Labels_Job')
            safe_project_name = clean_file_name(raw_project_name)
            time_stamp = datetime.now().strftime("%y%m%d_%H%M")
            signature_used = _signature_for(filepath) if horizontal_mode else None
            if into:
                # From a Packing Sheets project: everything goes into that project, which already has the Final file and
                # its Signature Links. Nothing of the project is changed until the run has worked (see the end), and a
                # failure takes back only the files this run wrote: the project itself is never removed.
                final_folder_name = into
                project_dir = os.path.join(app.config['UPLOAD_FOLDER'], into)
                project_had = set(os.listdir(project_dir))
                safe_project_name = clean_file_name(re.sub(r'_\d{6}_\d{4}$', '', into).split('_Job-')[0])
                log.info("  into project %s", into)
            else:
                final_folder_name = f"{safe_project_name}_{time_stamp}"
                project_dir = os.path.join(app.config['UPLOAD_FOLDER'], final_folder_name)
                os.makedirs(project_dir, exist_ok=True)

                # 4. Move the Excel File into the Project
                new_filepath = os.path.join(project_dir, filename)
                if signature_used:  # the project keeps the Signature Links file it was made with (not named 'Signature links…',
                    # so the Label Shuffler doesn't take this project for one of its own)
                    shutil.copy(_signature_path(filepath), os.path.join(project_dir, f"Used Signature Links - {os.path.splitext(filename)[0]}.xlsx"))
                shutil.move(filepath, new_filepath)
                log.info("  project folder %s", final_folder_name)

            if courier_mode:
                csv_files = _import_csvs(project_dir, reference, safe_project_name, csv_formats, cartons, consignments,
                                         fixed, request.form, log)
                output_files = csv_files
            else:
                # 5. Generate the PDF with Custom Layout
                # 'Packing Labels Only 1 - J477161 - Lux Test 30 - 261005_1732.pdf' (stitched: 'Complete Labels 1 - ...')
                output_pdf_name = unused_name(project_dir, numbered_name(project_dir, "Packing Labels Only", reference, safe_project_name, time_stamp))
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
                ref_issues = []
                edits_now = _consignment_edits(request.form)
                refs = {c['number']: _reference_for(c['destination'], edits_now.get(c['id'], {}))
                        for c in all_consignments if not c.get('no_address')}
                for c in consignments:
                    r = refs[c['number']]
                    # A street not on record counts too (unless the address is saved in the book): the preview
                    # ranks it with the major problems
                    if r['level'] == 'mismatch' or (r['level'] == 'note' and not r.get('soft')):
                        ref_issues.append(f"#{c['number']} {c['destination']['receiver']} - {one_line(c['destination'])}: {r['message']}")
                # The consignment preview's problems not fixed before Generate go on the audit pages too: each pack's
                # (horizontal) or each tab's (vertical), with the rows and stores they're on
                for key, issue in _unfixed_preview_problems(all_consignments, refs):
                    if horizontal_mode:
                        g = combined_pack_groups[key]
                        pack_notes.setdefault((g['tab_no'], g['pack_no']), []).append(issue)
                    elif key in tab_of:
                        tab_notes.setdefault(tab_of[key], []).append(issue)
                report = _generation_report(reference, safe_project_name, shipment_reference, sheet_warnings, courier_warnings,
                                            all_cartons, cartons, consignments, by_number, fixed['service_code'], sender, csv_formats,
                                            ref_issues)
                progress.stage("Drawing the packing labels", 20, 90)
                with app_log.step(log, f"  packing labels PDF {output_pdf_name}"):
                    if horizontal_mode:  # packing-sheet order: an audit page per pack, a divider per code
                        page_info = generate_horizontal_labels(combined_pack_groups, output_pdf, attribute_order, label_addresses,
                                                               report, serials, add_dividers=request.form.get('code_dividers') == '1',
                                                               group_by_code=request.form.get('label_grouping', 'code') != 'store',
                                                               notes=pack_notes, on_progress=progress.tick)
                    else:
                        page_info = generate_packing_labels(combined_pack_groups, output_pdf, attribute_order, label_addresses, report,
                                                            serials, {k: (n, selected_tabs[n - 1]) for k, n in tab_of.items()},
                                                            tab_notes=tab_notes, on_progress=progress.tick)
                log.info("    %d page(s), %d report page(s)", len(fitz.open(output_pdf)), page_info.get('report_pages', 0))

                # 6. The courier CSVs chosen (TIG Open360's bulk upload and / or OpenFreight's), and the label map used to
                # match courier labels back to packing label pages. Both files carry the same Item References
                # ('01T1QWXK Store': serial + the row's own code), so courier labels booked from either stitch by code.
                progress.stage("Writing the courier CSVs and the label map", 90, 99)
                base_name = f"{reference} - {safe_project_name}" if reference else safe_project_name
                map_name = unused_name(project_dir, f"{base_name}.labelmap.json")  # never over an earlier run's
                items = open360_items(cartons, consignments, tab_of, serials)
                if horizontal_mode:
                    # '03T1P2QWXK Chadstone': the serial printed on the packing label (store number, T + tab, P + pack)
                    # and a random code of this Generate, as vertical files have: serial + code is what Stitch Labels
                    # matches, so a courier label of another job (whose serials repeat) can't land on this one
                    codes = unique_codes(len(cartons))
                    open360_refs = [(f"{serials[x['pack_key']]}{code} {x['store']}", f"{serials[x['pack_key']]}{code}")
                                    for x, code in zip(cartons, codes)]
                else:
                    open360_refs = open360_item_references(items)
                ref_of = {x['item_reference']: r[0] for x, r in zip(cartons, open360_refs)}
                csv_files = []
                export, export_problems = clean_export(_export_values(request.form))
                if 'open360' in csv_formats:
                    open360_name = unused_name(project_dir, f"{base_name} - Open360.csv")
                    with app_log.step(log, f"  Open360 CSV {open360_name}"):
                        write_open360_csv(os.path.join(project_dir, open360_name), items, shipment_reference, item_refs=open360_refs)
                    csv_files.append(open360_name)
                    log.info("    %d row(s), Shipment Reference %s", len(open360_refs), shipment_reference)
                if 'openfreight' in csv_formats:
                    openfreight_name = unused_name(project_dir, f"{base_name} - OpenFreight.csv")
                    with app_log.step(log, f"  OpenFreight CSV {openfreight_name}"):
                        write_courier_csv(os.path.join(project_dir, openfreight_name), cartons, consignments, reference, fixed,
                                          item_refs=ref_of, export=export)
                        abroad = sum(1 for c in consignments if (c['destination'].get('country') or 'AU').upper() != 'AU')
                        if abroad:
                            log.info("    export details on %d consignment(s) outside Australia: %s", abroad, export)
                    csv_files.append(openfreight_name)
                for x, r in zip(cartons, open360_refs):
                    log.debug("    %s -> %s", x['item_reference'], r)
                # What both CSVs are made from, so the other one can be made later from the dashboard
                write_csv_recipe(os.path.join(project_dir, unused_name(project_dir, base_name + RECIPE_SUFFIX)), base_name,
                                 reference, shipment_reference, fixed, export, cartons, consignments, items, open360_refs,
                                 ref_of, {k: n for k, n in zip([f for f in CSV_FORMATS if f in csv_formats], csv_files)})
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
            used_copy = None
            if into:
                # It worked: a new label map starts a new stitching session, and the copy of the Final file the labels
                # were made from stays with them when it was changed (edited with Open Excel on this page), so the
                # project shows what the labels came from; an unchanged copy goes
                _keep_earlier_session(project_dir, map_name)
                used_copy = _keep_used_copy(filepath, project_dir, filename, time_stamp, log)
                if used_copy:
                    output_files.append(used_copy)
            _record(project_dir, 'courier_import' if courier_mode else 'labels',
                                  output_files + ([map_name] if not courier_mode else []) + ([] if into else [filename]),
                                  **({} if courier_mode else {'layout': 'horizontal' if horizontal_mode else 'vertical',
                                                              'from_final_file': bool(into)}))
            record_service_usage(SERVICE_USAGE_FILE, consignments, SERVICE_STORE.codes())
            try:
                ADDRESS_BOOK.record_used([(c['destination'], [x['source_id'] for x in c['cartons']]) for c in consignments])
                log_book.info("Address book: %d consignment address(es) saved / marked used", len(consignments))
                # Kept as is in the preview despite the official data: the book remembers, so it isn't flagged again
                edits = _consignment_edits(request.form)
                kept = [c for c in consignments if isinstance(edits.get(c['id'], {}).get('ref_keep'), dict)
                        and _same_place(edits[c['id']]['ref_keep'], c['destination'])]
                for c in kept:
                    ADDRESS_BOOK.keep_address_as_is(c['destination'])
                if kept:
                    log_book.info("Address book: %d address(es) kept as is despite the official data", len(kept))
                recheck_book_async('generate')
            except Exception as e:
                log_book.exception("Could not update the address book after Generate: %s", e)

            if horizontal_mode and os.path.exists(_signature_path(filepath)):
                os.remove(_signature_path(filepath))
            log.info("Generate done in %.2fs: %s", time.perf_counter() - started, final_folder_name)
            left_out = [x for x in all_cartons if x not in cartons]
            summary = {'items': len(cartons), 'consignments': len(consignments), 'left_out': [
                f"{x['item_reference']} ({x['packing_spec']}) - {x['store']}" for x in left_out]} if courier_mode else None
            return show(generation_complete=True, project_folder=final_folder_name, generated_files=output_files,
                        summary=summary)

        except Exception as e:
            log.exception("Generate FAILED after %.2fs (project folder undone)", time.perf_counter() - started)
            for i in getattr(e, 'issues', None) or []:  # a check that stopped it: each problem, where it is
                log.warning("    %s", issue_text(i))
            failed_at = app_log.last_error()
            # Undo a half-made project so the upload is back where it was and Generate can be retried
            if project_had is not None and os.path.isdir(project_dir):
                # An existing Packing Sheets project: only what this run wrote goes; the project and the copy being
                # worked on stay, so Generate can be retried
                for name in set(os.listdir(project_dir)) - project_had:
                    try:
                        os.remove(os.path.join(project_dir, name))
                    except OSError:
                        log.warning("  could not take back %s from %s", name, into)
            elif project_dir and os.path.isdir(project_dir):
                moved = os.path.join(project_dir, filename)
                if os.path.exists(moved) and not os.path.exists(filepath):
                    shutil.move(moved, filepath)
                shutil.rmtree(project_dir, ignore_errors=True)
            page = show(tabs=all_tabs, filename=filename, user_inputs=user_inputs, page_error=str(e))
            if isinstance(e, (PackCheckError, ValueError)):
                return page  # a check of the file or the choices (a service code, a column): fixed on this page
            # Anything else lands on the Logs page at the failure; its Back link is this page as it was, to try again
            return redirect(_logs_for_error(back=getattr(page, 'location', None), error=failed_at), code=303)

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
    # Old projects are cleared when the app starts (python app.py, the .bat, the .exe), never when another script
    # imports app.py: tests and tools that import it must not delete the user's projects
    clean_old_projects()
    # In debug mode Flask runs this file twice: a watcher that restarts the server on code changes, and the
    # server itself (WERKZEUG_RUN_MAIN=true). Only the first may clear the port, or the server would stop its own watcher.
    if os.environ.get('WERKZEUG_RUN_MAIN') != 'true':
        free_port(PORT)
    # The Flask debugger runs any code typed into its error page: only on request (set WAREHOUSE_DEBUG=1), never in
    # the .exe, and never reachable from another computer (no host= here: 127.0.0.1 only)
    debug = os.environ.get('WAREHOUSE_DEBUG') == '1' and not getattr(sys, 'frozen', False)
    app.run(debug=debug, port=PORT)

    
