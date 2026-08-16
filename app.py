import time
from datetime import datetime
from flask import Flask, render_template, request, redirect, session, url_for, send_file
import os
import shutil
import json
import sys
import stat
import pprint
import pandas as pd
from werkzeug.utils import secure_filename
from pdf_engine import process_and_shuffle_pdf
from matrix_engine import clean_file_name, scan_excel_tabs, generate_tab_map, generate_all_outputs
from core_math import clean_file_name, get_available_project_files, close_if_open_elsewhere
from subgroup_engine import execute_subgroups, SubgroupValidationError



if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

        

PROJECTS_FOLDER = os.path.join(BASE_DIR, 'projects')
os.makedirs(PROJECTS_FOLDER, exist_ok=True)
temp_dir = os.path.join(BASE_DIR, 'temp_pdf_engine')
os.makedirs(temp_dir, exist_ok=True)

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
            if (current_time - creation_time) > seven_days_in_seconds:
                try:
                    shutil.rmtree(folder_path, ignore_errors=True)
                    print(f"Cleaned up old project: {folder_name}")
                except Exception as e:
                    print(f"Could not delete {folder_name}: {e}")

clean_old_projects()  # Retry cleanup if deletion fails
app = Flask(__name__, 
            template_folder=os.path.join(BASE_DIR, 'templates'),
            static_folder=os.path.join(BASE_DIR, 'static'))
app.config['UPLOAD_FOLDER'] = PROJECTS_FOLDER
app.config['TEMP_FOLDER'] = temp_dir
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['TEMP_FOLDER'], exist_ok=True)

# --- PART 1: MATRIX ENGINE ---
@app.route('/matrix', methods=['GET', 'POST'])
def matrix():
    tabs = None
    filename = None
    
    if request.method == 'POST':
        if 'file' not in request.files:
            return "No file part"
            
        file = request.files['file']
        
        if file.filename != '':
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            tabs = scan_excel_tabs(filepath)
            
    return render_template('matrix.html', tabs=tabs, filename=filename)

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
            
    return render_template('matrix.html', tabs=all_tabs, previews=previews, filename=filename, user_inputs=user_inputs, blueprints_json=blueprints_json)

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
    
   # 1. Grab the user's custom project name from the form
    raw_project_name = request.form.get('project_name', 'Untitled_Project')
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
    file1_name, file2_name, file3_name = generate_all_outputs(
        new_filepath, filename, selected_tabs, user_inputs, blueprints, project_dir
    )
    raw_files = [file1_name, file2_name, file3_name]

    # Filter the list to ONLY keep actual file names (removes None/Empty strings)
    files_to_download = [f for f in raw_files if f]

    # 3. Pass the clean list to the template
    return render_template('matrix.html',
                           generation_complete=True,
                           project_folder=final_folder_name,
                           generated_files=files_to_download)

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
                display_files = [f for f in files if not f.endswith('.json')]
                # Sort JSON files by creation time (newest first)
                json_files = sorted([f for f in files if f.endswith('.json')], key=lambda x: os.path.getctime(os.path.join(folder_path, x)), reverse=True )

                
                excel_files = [f for f in display_files if f.lower().endswith(('.xlsx', '.xls'))]
                pdf_files = [f for f in display_files if f.lower().endswith('.pdf')]
                
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
                
                projects.append({
                    "name": folder_name,          
                    "display_name": display_name, 
                    "date": date_str,
                    "timestamp": timestamp,
                    "job_id": job_id,
                    "excel_files": excel_files,  # Pass the grouped Excel files
                    "pdf_files": pdf_files,       # Pass the grouped PDFs
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
        json_files = [f for f in os.listdir(project_dir) if f.endswith('.json')]
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
    os.makedirs(PROJECTS_FOLDER, exist_ok=True)
    
    if request.method == 'POST':
        step = request.form.get('step')
        
       # STEP 1: Process Excel and Project Name
        if step == '1':
            excel_file = request.files.get('excel_file')
            existing_project = request.form.get('existing_project')
            new_project = request.form.get('new_project')
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
            try:
                with pd.ExcelFile(excel_path) as xls:
                    for sheet in xls.sheet_names:
                        df = pd.read_excel(xls, sheet_name=sheet)
                        packs = df.columns[1:].tolist()
                        tabs_data[sheet] = packs
                        # --- CHECK FOR DUPLICATE STORE NAMES (COLUMN 0) ---
                        if not df.empty:
                            # Store names live in column 0
                            store_col = df.iloc[:, 0].dropna().astype(str).str.strip()
                            store_col = store_col[store_col.str.lower() != 'nan']  # Remove string 'nan'
                            
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
                return render_template('pdf.html',
                                       step=1,
                                       existing_projects=list(projects_info.keys()),
                                       projects_json=projects_json,
                                       duplicate_errors=duplicate_errors,
                                       duplicate_project_name=os.path.basename(project_folder),
                                       duplicate_excel_filename=os.path.basename(excel_path))
                
            return render_template('pdf.html', step=2, tabs_data=tabs_data, excel_path=excel_path, project_name=project_name)
            
        # STEP 2: Process the PDFs and Go to Success Screen
        elif step == '2':
            excel_path = request.form.get('excel_path')
            project_name = request.form.get('project_name') 
            temp_dir = os.path.join(BASE_DIR, 'temp_pdf_engine')
            os.makedirs(temp_dir, exist_ok=True)
            # ---> THE FIX: Define project_folder HERE, before any loops!
            project_folder = os.path.join(PROJECTS_FOLDER, os.path.basename(project_name))
            os.makedirs(project_folder, exist_ok=True)

            all_sheets_data = {}
            
            try:
                with pd.ExcelFile(excel_path) as xls:
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
                            add_dividers=add_dividers_flag
                        )
                        
                        # Log it for the download screen
                        generated_files.append(final_filename)

            try:
                shutil.rmtree(temp_dir)
            except Exception:
                pass

            return render_template('pdf.html', step=3, project_name=project_name, generated_files=generated_files)

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
    return render_template('pdf.html', step=1, existing_projects=list(projects_info.keys()), projects_json=projects_json)

if __name__ == '__main__':
    if getattr(sys, 'frozen', False):
        app.run(debug=False, port=5000)
    else:
        app.run(debug=True, port=5000)

    
