import time
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, send_file
import os
import shutil
import json
import sys
import stat
from werkzeug.utils import secure_filename
from matrix_engine import clean_file_name, scan_excel_tabs, generate_tab_map, generate_all_outputs

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PROJECTS_FOLDER = os.path.join(BASE_DIR, 'projects')
os.makedirs(PROJECTS_FOLDER, exist_ok=True)

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
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

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
                
                # Get human-readable date
                timestamp = os.path.getctime(folder_path)
                date_str = datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M')

                # Extract the Job ID dynamically from File 2's name
                job_id = "N/A"
                for f in files:
                    if f.startswith("Packing Sheet_"):
                        job_id = f.replace("Packing Sheet_", "").rsplit(".", 1)[0]
                        break
                        
                # --- CLEAN DISPLAY NAME ---
                # Chops off "_Job-" and the timestamp for the UI header
                display_name = folder_name.split('_Job-')[0] if '_Job-' in folder_name else folder_name
                
                projects.append({
                    "name": folder_name,          # Keep full name for backend links
                    "display_name": display_name, # Send clean name for the frontend
                    "date": date_str,
                    "timestamp": timestamp,
                    "job_id": job_id,
                    "files": files
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
@app.route('/pdf')
def pdf_shuffler():
    return "<h1 style='font-family:sans-serif; text-align:center; margin-top:50px;'>PDF Shuffler Engine Coming Soon!</h1><center><a href='/'>Go Back</a></center>"

if __name__ == '__main__':
    app.run(debug=True)