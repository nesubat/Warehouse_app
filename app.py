from flask import Flask, render_template, request, redirect, url_for, send_file
import os
import shutil
import json
from werkzeug.utils import secure_filename
from matrix_engine import scan_excel_tabs, generate_tab_map, generate_all_outputs

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'input_files'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

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
    
    # Notice we now pass blueprints directly to the engine
    file1_name, file2_name, file3_name = generate_all_outputs(filepath, filename, selected_tabs, user_inputs, blueprints)
    
    return render_template('matrix.html', 
                           generation_complete=True,
                           file1=file1_name,
                           file2=file2_name,
                           file3=file3_name)

@app.route('/download/<filename>')
def download_file(filename):
    secure_file = secure_filename(filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_file)
    
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
        
    return "File Not Found", 404

@app.route('/pdf')
def pdf_shuffler():
    return "<h1 style='font-family:sans-serif; text-align:center; margin-top:50px;'>PDF Shuffler Engine Coming Soon!</h1><center><a href='/'>Go Back</a></center>"

if __name__ == '__main__':
    app.run(debug=True)