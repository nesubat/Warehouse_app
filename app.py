from flask import Flask, render_template, request, redirect, url_for, send_file
import os
import shutil
from werkzeug.utils import secure_filename
from matrix_engine import scan_excel_tabs, generate_tab_map, generate_all_outputs
from datetime import datetime, timedelta
import re
from flask import send_from_directory

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'input_files'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def cleanup_old_folders(uploads_dir):
    """Deletes any project folder older than 14 days."""
    if not os.path.exists(uploads_dir):
        return
    
    # Set the threshold to exactly 14 days ago
    cutoff_date = datetime.now() - timedelta(days=14)
    
    for folder_name in os.listdir(uploads_dir):
        folder_path = os.path.join(uploads_dir, folder_name)
        
        if os.path.isdir(folder_path):
            try:
                # We pull the YYYYMMDD_HHMMSS part from the folder name
                date_str = folder_name[:15]
                folder_date = datetime.strptime(date_str, '%Y%m%d_%H%M%S')
                
                # If the folder is too old, drop the wrecking ball
                if folder_date < cutoff_date:
                    shutil.rmtree(folder_path)
            except ValueError:
                # If the folder name doesn't match our timestamp format, ignore it safely
                pass

# --- THE MAIN MENU ---
@app.route('/')
def index():
    return render_template('index.html')

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
    
    for tab in all_tabs:
        start_cell = request.form.get(f"start_{tab}", "B8")
        job_id_cell = request.form.get(f"job_{tab}", "E1")
        store_col = request.form.get(f"store_{tab}", "A")
        
        user_inputs[tab] = {
            "start": start_cell,
            "job": job_id_cell,
            "store": store_col,
            "selected": tab in selected_tabs
        }
        
        if tab in selected_tabs:
            try:
                blueprint = generate_tab_map(filepath, tab, start_cell, job_id_cell, store_col)
                previews.append(blueprint)
            except Exception as e:
                previews.append({
                    "sheet_name": tab,
                    "error": f"Failed to map. Check your coordinates! ({str(e)})"
                })
            
    return render_template('matrix.html', tabs=all_tabs, previews=previews, filename=filename, user_inputs=user_inputs)


@app.route('/generate', methods=['POST'])
def generate():
    filename = request.form.get('filename')
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    
    all_tabs = request.form.getlist('all_tabs')
    selected_tabs = request.form.getlist('selected_tabs')
    
    user_inputs = {}
    for tab in all_tabs:
        user_inputs[tab] = {
            "start": request.form.get(f"start_{tab}"), #capturing th start cell for each tab
            "job": request.form.get(f"job_{tab}"), #capturing the job id cell for each tab
            "store": request.form.get(f"store_{tab}"), #capturing the store column for each tab
            "selected_packs": request.form.getlist(f"packs_{tab}")  #capturing the selected packages for each tab
        }
    
        
        # 2. Grab the Campaign Name and create a unique folder name
        raw_campaign = request.form.get("campaign_name", "Untitled_Project")
        clean_campaign = re.sub(r'[^A-Za-z0-9 _-]', '', raw_campaign).strip()
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        folder_name = f"{timestamp}_{clean_campaign}"
        
        # Create the new folder inside the uploads directory
        new_project_dir = os.path.join(app.config['UPLOAD_FOLDER'], folder_name)
        os.makedirs(new_project_dir, exist_ok=True)
        
        # 3. Move the uploaded file from the main directory into its new unique folder
        old_filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        new_filepath = os.path.join(new_project_dir, filename)
        
        if os.path.exists(old_filepath):
            try:
                # Try to move it cleanly
                shutil.move(old_filepath, new_filepath)
            except PermissionError:
                # If Excel is secretly holding the file open in the background, just copy it!
                shutil.copy(old_filepath, new_filepath)
        else:
            # If the file is already gone (e.g., button clicked twice), just proceed if it's already in the new folder
            if not os.path.exists(new_filepath):
                return "Error: Could not find the uploaded file.", 404
        
        # 4. Run the Engine using the NEW filepath inside the isolated folder!
        file1_name, file2_name, file3_name = generate_all_outputs(new_filepath, filename, selected_tabs, user_inputs)
        
        # Render your success template (We will update this to redirect to the Dashboard next)
        return render_template('matrix.html', 
                               generation_complete=True,
                               file1=file1_name,
                               file2=file2_name,
                               file3=file3_name,
                               folder_name=folder_name) # Pass the folder name to HTML
    
@app.route('/dashboard')
def dashboard():
    projects = []
    if os.path.exists(app.config['UPLOAD_FOLDER']):
        for folder_name in os.listdir(app.config['UPLOAD_FOLDER']):
            folder_path = os.path.join(app.config['UPLOAD_FOLDER'], folder_name)
            
            if os.path.isdir(folder_path):
                try:
                    # Parse the folder name: YYYYMMDD_HHMMSS_CampaignName
                    date_str = folder_name[:15]
                    folder_date = datetime.strptime(date_str, '%Y%m%d_%H%M%S')
                    campaign_name = folder_name[16:].replace('_', ' ')
                    
                    # Scan the files inside the folder
                    files = os.listdir(folder_path)
                    file1 = next((f for f in files if f.startswith('Final')), None)
                    file2 = next((f for f in files if f.startswith('Packing Sheet')), None)
                    file3 = next((f for f in files if f.startswith('Signature links')), None)
                    
                    projects.append({
                        'folder_name': folder_name,
                        'campaign_name': campaign_name,
                        'date': folder_date.strftime('%B %d, %Y at %I:%M %p'),
                        'sort_date': folder_date,
                        'file1': file1,
                        'file2': file2,
                        'file3': file3
                    })
                except ValueError:
                    pass # Ignore folders that don't match our exact timestamp format

    # Sort so the newest projects appear at the top left
    projects.sort(key=lambda x: x['sort_date'], reverse=True)
    return render_template('dashboard.html', projects=projects)

    


@app.route('/download/<folder_name>/<filename>')
def download_file(folder_name, filename):
    """Secure endpoint allowing users to pull individual output sheets from specific project folders."""
    
    # Secure the inputs to prevent directory traversal attacks
    secure_folder = secure_filename(folder_name)
    secure_file = secure_filename(filename)
    
    # Path is now: UPLOAD_FOLDER -> Specific Campaign Folder -> File
    folder_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_folder)
    file_path = os.path.join(folder_path, secure_file)
    
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
        
    return "File Not Found", 404

@app.route('/delete/<folder_name>', methods=['POST'])
def delete_project(folder_name):
    """Manually deletes a specific project folder when the user clicks Delete."""
    # Use basename to prevent directory traversal WITHOUT mangling spaces
    safe_folder = os.path.basename(folder_name)
    folder_path = os.path.join(app.config['UPLOAD_FOLDER'], safe_folder)
    
    # If the folder exists, drop the wrecking ball
    if os.path.exists(folder_path) and os.path.isdir(folder_path):
        try:
            # Attempt to remove the folder and its contents
            shutil.rmtree(folder_path)
        except PermissionError:
            # Windows blocked it! Return a friendly error instead of crashing.
            return """
            <div style="font-family: sans-serif; text-align: center; margin-top: 50px;">
                <h2 style="color: #e74c3c;">❌ Access Denied</h2>
                <p>Windows blocked the deletion because the folder (or an Excel file inside it) is currently open.</p>
                <p>Please close any File Explorer windows, ensure Excel is fully closed, and try again.</p>
                <a href='/dashboard' style="display: inline-block; margin-top: 20px; padding: 10px 20px; background-color: #3498db; color: white; text-decoration: none; border-radius: 5px;">Return to Dashboard</a>
            </div>
            """, 403
        
    # Send them back to the dashboard to see the updated list
    return redirect(url_for('dashboard'))


# --- PART 2: PDF SHUFFLER (Coming Soon) ---
@app.route('/pdf')
def pdf_shuffler():
    return "<h1 style='font-family:sans-serif; text-align:center; margin-top:50px;'>PDF Shuffler Engine Coming Soon!</h1><center><a href='/'>Go Back</a></center>"

if __name__ == '__main__':
    # 1. Ensure the uploads folder exists before the janitor runs
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    
    # 2. Run the janitor instantly when the app boots up
    cleanup_old_folders(app.config['UPLOAD_FOLDER'])

    # 3. Start the local server
    app.run(debug=True)