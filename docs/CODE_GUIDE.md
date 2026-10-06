# 🧭 Code Guide — Warehouse Automation Suite

This document explains **every file in this project, piece by piece**, in plain language. It's written for someone with only basic coding knowledge (you understand variables, loops, `if`/`else`, and functions, but you don't need to remember Flask, Excel-automation, or PDF-library details — they're explained here).

It covers all four tools — the Distribution Mapper (Matrix Engine), Sub-Groups, the Label Shuffler and Packing Labels with its courier CSV — and, for Packing Labels, also lists the rules the tool applies, so this one file is both the "how it works" and the "what it does" reference.

> **Note on diagrams:** This file uses [Mermaid](https://mermaid.js.org/) diagrams (the ` ```mermaid ` code blocks). They render automatically on GitHub/GitLab. In VS Code, install the **"Markdown Preview Mermaid Support"** extension and open Preview (`Ctrl+Shift+V`) to see them as pictures instead of code.

---

## Table of Contents

1. [The Big Picture](#1-the-big-picture)
2. [Key Vocabulary](#2-key-vocabulary)
3. [File Map & Folder Structure](#3-file-map--folder-structure)
4. [`app.py` — The Web Server](#4-apppy--the-web-server)
5. [`core_math.py` — Shared Math & Formatting](#5-core_mathpy--shared-math--formatting)
6. [`matrix_engine.py` — The Matrix Engine](#6-matrix_enginepy--the-matrix-engine)
7. [`subgroup_engine.py` — The Sub-Group Engine](#7-subgroup_enginepy--the-sub-group-engine)
8. [`pdf_engine.py` — The Label Shuffler](#8-pdf_enginepy--the-label-shuffler)
9. [`packing_label_generator.py` — Packing Labels](#9-packing_label_generatorpy--packing-labels)
10. [`courier_export.py` — Addresses, Consignments & the Courier CSV](#10-courier_exportpy--addresses-consignments--the-courier-csv)
11. [`templates/*.html` — The Web Pages](#11-templateshtml--the-web-pages)
12. [`static/script.js` — The Frontend Brain](#12-staticscriptjs--the-frontend-brain)
13. [`static/styles.css` — The Look & Feel](#13-staticstylescss--the-look--feel)
14. [End-to-End Journeys](#14-end-to-end-journeys)
15. [Quick Reference: "I Want to Change X"](#15-quick-reference-i-want-to-change-x)

---

## 1. The Big Picture

This app is a **Flask web application** — a Python program that runs a small local web server. You open a browser, go to `http://127.0.0.1:5001`, and click buttons; behind the scenes, Python code reads/writes real Excel, PDF and CSV files on your hard drive.

There is no database. Everything the app "remembers" is stored as **actual files on disk**, inside a `projects/` folder — Excel files, PDF files, courier CSVs, and small `.json` text files that act as the app's memory of what it did. One extra file, `data/service_code_usage.json`, remembers how often each courier service code has been used.

```mermaid
flowchart LR
    Browser["🖥️ Your Browser"] -->|"HTTP (clicks, uploads, downloads)"| Flask["app.py\n(Flask routes)"]
    Flask --> Browser
    Flask --> Matrix["matrix_engine.py"]
    Flask --> Sub["subgroup_engine.py"]
    Flask --> PDF["pdf_engine.py"]
    Flask --> PL["packing_label_generator.py"]
    Flask --> CE["courier_export.py"]
    CE --> PL
    Matrix --> Core["core_math.py\n(shared helpers)"]
    Sub --> Core
    Matrix -.controls.-> Excel[("Microsoft Excel\n(via xlwings)")]
    Sub -.controls.-> Excel
    PDF -.processes.-> Fitz[("PDF files\n(via PyMuPDF/fitz)")]
    PL -.draws.-> Fitz
    Flask --> FS[("projects/ folder\non your hard drive")]
    Flask -->|renders| Templates["templates/*.html"]
    Templates --> JS["static/script.js"]
    Templates --> CSS["static/styles.css"]
```

Four independent "engines" do the heavy lifting, and `app.py` is the traffic cop that receives clicks from the browser and calls the right engine at the right time:

| Engine | File | What it does |
|---|---|---|
| **Matrix Engine** | `matrix_engine.py` | Turns a raw distribution Excel file into a "Packing Sheet" with calculated signature codes. |
| **Sub-Group Engine** | `subgroup_engine.py` | Takes an existing Packing Sheet and breaks a pack into smaller item-number ranges (repeatable — Stage 1, Stage 2, Stage 3...). |
| **PDF Label Shuffler** | `pdf_engine.py` | Takes raw PDF shipping labels and re-orders/re-stacks their pages to match the signature codes calculated above. |
| **Packing Labels** | `packing_label_generator.py` + `courier_export.py` | Turns a packing-spec spreadsheet into one printable packing label per box (PDF), plus a courier consignment CSV and a label map for matching courier labels to packing labels. |

`core_math.py` is a toolbox of small functions (cleaning text, calculating "signatures", formatting Excel cells) that both the Matrix Engine and the Sub-Group Engine call, so the same logic isn't written twice. `courier_export.py` likewise borrows the text-comparison helpers from `packing_label_generator.py`.

---

## 2. Key Vocabulary

These words appear everywhere in the code. If you're ever confused reading a function, come back to this table.

**Distribution Mapper, Sub-Groups and Label Shuffler**

| Term | Meaning |
|---|---|
| **Tab** | A sheet inside an Excel workbook (e.g., "Store A", "Store B"). The app processes one or more tabs independently. |
| **Pack** | A group of columns in a tab that represents one "package type" (e.g., "Banner Pack", "Signage Pack"). Detected either from a merged header cell, or a single column with a header. |
| **Store** | A row in a tab — one retail location receiving an allocation. |
| **Signature / Signature Code** | A short letter code (`A`, `B`, ... `Z`, `AA`, `AB`...) assigned to a **unique combination of quantities** across a pack's columns. Two stores with the exact same order quantities get the same letter. This is the core trick of the whole app — instead of printing every store's exact numbers, you print one letter, and a lookup table tells the warehouse floor what that letter means. |
| **Item Number** | A number identifying one column/item inside a pack (used only by the Sub-Group Engine to slice a pack into smaller ranges, e.g. items 1–5 vs 6–10). |
| **Stage** | A version of the output Excel file. `Stage - 1 - ...xlsx` is the first sub-grouped version, `Stage - 2 - ...xlsx` the next, etc. Each stage builds on the previous one without destroying it. |
| **Blueprint** | The in-memory map the Matrix Engine builds after scanning a tab: where the stores are, where the packs are, where the Job ID is. Shown to the user as a "preview" before generating anything. |
| **Metadata (`project_metadata.json` / `Stage - N - ....json`)** | A JSON file that remembers exact column numbers for every pack/sub-group in a project, so the app can find its own past work later (e.g. when you come back to sub-group a project a second time). |

**Packing Labels and the courier CSV** — note that here a "pack" means something different: one physical box.

| Term | Meaning |
|---|---|
| **Pack group** | The rows covered by one merged **Packing Spec** cell. One pack group = one box = one packing label = one courier carton = one CSV row. |
| **Packing Spec** | The box type, e.g. `OB1370170170` (a 137 × 17 × 17 cm box), `P7 Jiffy Bag`, `A4 Box`, `Pallet…`. |
| **Installer pack** | A pack whose **Install** cell says `Y`/`Yes`/`True`: it goes to an installer, not to the store. |
| **Label X of Y** | Each store's pack groups numbered in file order. Printed on the packing label and used in the courier item reference. |
| **Consignment** | All packs going to the **same delivery address**, whichever store they belong to. The courier treats them as one shipment. |
| **Consignment reference** | Pre-filled from the distribution file's Job Number column (`job_series()` / `detect_series()`): strictly **J + 6 digits** at the start (`J477161-54` → `J477161`). Job numbers in any other format are ignored; if none match, the field starts empty and must be typed. When the file has several jobs, the most common one, with a warning. Editable in the preview; it is also the Open360 **Shipment Reference** on every row. |
| **Item reference** | `Label <X> - <Store>`, the last CSV column. Printed on the courier label; it's the key that links a courier label to its packing label. |
| **Label map** | `….labelmap.json`, saved next to the PDF: which PDF page each label (each box) is on, where it's headed, its CSV row and its code (serial + job code, `01T1QWXK`). Stitch Labels uses it to put each courier label on its packing label. |
| **Comparison key** | A simplified copy of a text value (lower-case, no punctuation, Street = St…) used to decide whether two spellings mean the same thing. |
| **Service code** | The courier service for a consignment (e.g. `STEROAD · STARTRACK ROAD EXPRESS`), picked from a list. |
| **Address book** | `data/address_book.db`: every clean delivery address the app has sent, plus the raw Excel spellings that led to each one (**aliases**), so they're filled in automatically next time. |

**Libraries and general terms**

| Term | Meaning |
|---|---|
| **Project Folder** | One folder per job, named like `CampaignName_Job-12345_260803_1430`, holding every file (Excel, PDF, CSV, JSON) generated for that job. |
| **`xlwings`** | A Python library that remote-controls a real, invisible copy of Microsoft Excel. Used because it can insert/merge columns and preserve formatting exactly like a human using Excel would — something the faster libraries (openpyxl/pandas) can't do well. |
| **`openpyxl`** | Reads `.xlsx` files directly (no Excel needed). Packing Labels uses it to read cells, merged ranges and pictures. |
| **`fitz` (PyMuPDF)** | A Python library for reading and rewriting PDF files page-by-page — used to detect text on a label, to cut/paste/reorder pages, and to draw packing labels from scratch. The PyPI/pip package is still called `PyMuPDF`, but the *import name* `fitz` is a deprecated legacy alias upstream; the code imports it as `import pymupdf as fitz` so the rest of the file can keep using the familiar `fitz.` prefix without triggering the deprecation warning. |
| **Point (pt) / EMU** | PDF drawing units: 1 pt = 1/72 inch, 1 mm = 2.83465 pt. Excel positions pictures in EMU (English Metric Units): 12,700 per point, 9,525 per pixel. |

---

## 3. File Map & Folder Structure

```mermaid
graph TD
    app["app.py"] --> matrix["matrix_engine.py"]
    app --> sub["subgroup_engine.py"]
    app --> pdf["pdf_engine.py"]
    app --> core["core_math.py"]
    app --> plg["packing_label_generator.py"]
    app --> ce["courier_export.py"]
    matrix --> core
    sub --> core
    ce --> plg
```

Only `app.py` talks to Flask/the browser. The engines never render a web page themselves — they take plain Python data in, and return plain Python data (or files on disk) out. This separation means you could, in theory, delete the entire website and still call `generate_all_outputs()` or `generate_packing_labels()` from a Python script.

### What's on disk after you use the app

```
Warehouse_app/
├── app.py, core_math.py, matrix_engine.py, pdf_engine.py, subgroup_engine.py
├── packing_label_generator.py, courier_export.py, address_book.py, address_import.py, address_match.py, packing_specs.py, label_stitcher.py
├── templates/            → the 5 HTML pages Flask renders
├── static/               → script.js + styles.css (shared by all pages)
├── docs/CODE_GUIDE.md    → this guide
├── data/                  → service_code_usage.json (how often each courier service was picked)
│                             address_book.db (the address book; -wal/-shm files appear next to it while the app runs)
├── temp_pdf_engine/       → scratch space, wiped after every PDF run
└── projects/              → EVERY job you've ever created lives here
    ├── MyCampaign_Job-12345_260804_1000/        ← Distribution Mapper project
    │   ├── OriginalUpload.xlsx              ← your original file (matrix flow)
    │   ├── Final OriginalUpload.xlsx        ← "File 1": working copy with signature columns
    │   ├── Packing Sheet_12345.xlsx         ← "File 2": master summary book (this is what floor staff read)
    │   ├── Signature links_12345.xlsx       ← "File 3": clean store→code lookup table (used by the PDF shuffler)
    │   ├── Final OriginalUpload.json        ← metadata memory for this project
    │   ├── Stage - 1 - Final OriginalUpload.xlsx   ← after your 1st sub-group run
    │   ├── Stage - 1 - Final OriginalUpload.json
    │   └── SomeLabelFile_Shuffled.pdf       ← output of the PDF shuffler
    └── LUX_261004_0152/                      ← Packing Labels project
        ├── PACKING_SPECS_….xlsx             ← your original file
        ├── Labels_PACKING_SPECS_….pdf       ← one packing label per box
        ├── J477161 - LUX.csv                ← courier consignment CSV
        └── J477161 - LUX.labelmap.json      ← label map (kept for stitching, never shown in the app)
```

Notice the **three-file pattern** from the Matrix Engine: **File 1** (working copy, has the new signature columns baked in), **File 2** (the human-readable Packing Sheet, columns removed/summarized), **File 3** (a clean, minimal store→code table meant to feed the PDF engine). All three are built from the same uploaded Excel, at the same time, inside `generate_all_outputs()` (see [Section 6](#6-matrix_enginepy--the-matrix-engine)).

---

## 4. `app.py` — The Web Server

**What Flask is, in one paragraph:** Flask lets you write a normal Python function and attach it to a URL with `@app.route('/some/url')`. When a browser visits that URL, Flask calls your function and sends whatever it `return`s back as the web page. `methods=['GET', 'POST']` means the function handles both "just looking at the page" (GET) and "submitting a form" (POST) — you check `request.method` inside the function to tell which one happened.

### 4.1 Startup

```python
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
```
This figures out "where am I actually running from." `sys.frozen` is `True` only when the app has been bundled into a standalone `.exe` by PyInstaller (see the README's `pyinstaller --onefile` step) — in that case, the base folder is wherever the `.exe` sits, not wherever the raw `.py` source happens to be.

```python
def clean_old_projects():
    ...
clean_old_projects()
```
Runs **once, immediately, the moment the file is imported** (note it's called at module level, not inside a route). It loops through every folder in `projects/`, checks its creation timestamp, and deletes anything older than 7 days with `shutil.rmtree`. This is the "Zombie Sweeper" mentioned in the README — it keeps old test jobs from piling up forever.

The same startup sweep now also removes **abandoned uploads**: loose `.xlsx`/`.xls` files left in the `projects/` root, and project folders that hold nothing but spreadsheets, once they're more than 1 day old (see [4.4](#44-project--file-management-delete-delete_file-open_local-open_upload) for how they're normally cleaned up straight away).

**Freeing port 5001 (`free_port`).** When `app.py` is started directly (`python app.py` or the `.exe`), it first stops **any process already listening on port 5001** — typically an earlier copy of the app left running in another terminal — so the new copy never fails with "address already in use". It reads Windows' `netstat -ano` (a listening socket is one whose foreign address ends in `:0`, which works whatever language Windows is in), stops each process found with `taskkill /T /F` (its child processes too), prints `Stopped python.exe (PID …), which was using port 5001.`, and waits until the port is free. In debug mode Flask runs `app.py` twice — a watcher that restarts the server whenever code changes, plus the server itself, which gets the environment variable `WERKZEUG_RUN_MAIN=true` — so the check runs only when that variable is **not** set; otherwise every code-change reload would stop the app's own watcher.

There's no secret key and no `.env` file: the app never uses Flask's `session` or `flash()`. Error messages are passed straight into the page as template variables instead (e.g. `page_error` on the Packing Labels page).

The rest of this section just creates the Flask `app` object and tells it where `templates/` and `static/` live, plus creates the `projects/` and `temp_pdf_engine/` folders if they don't exist yet.

### 4.2 The Matrix Engine routes: `/matrix` → `/preview` → `/generate` → `/download`

This is a **3-step wizard**, and each step is its own route because each one is a full page reload (the form `POST`s to the next URL).

```mermaid
sequenceDiagram
    participant U as Browser
    participant M as /matrix (GET+POST)
    participant P as /preview (POST)
    participant G as /generate (POST)
    participant D as /download (GET)

    U->>M: Load page (GET)
    M-->>U: Empty upload form
    U->>M: Upload Excel file (POST)
    M->>M: scan_excel_tabs() → list of sheet names
    M-->>U: Show one card per tab + coordinate fields (Start Cell, Job ID, Store Col)
    U->>P: Submit coordinates for the tabs you checked
    P->>P: generate_tab_map() for each selected tab
    P-->>U: Show "Blueprint" preview cards (pack list, warnings, checkboxes)
    U->>G: Submit selected packs + a Project Name
    G->>G: generate_all_outputs() → writes File 1/2/3 + metadata JSON
    G-->>U: Show download links
    U->>D: Click a download link
    D-->>U: send_file() streams the Excel file back
```

**`/matrix`.** On GET, does nothing but render the empty upload form. On POST:
```python
if 'file' not in request.files:
    return "No file part"
file = request.files['file']
if file.filename != '':
    filename = secure_filename(file.filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)
    tabs = scan_excel_tabs(filepath)
return render_template('matrix.html', tabs=tabs, filename=filename)
```
`request.files['file']` is the uploaded file object (from the `<input type="file" name="file">` in `matrix.html`). `secure_filename()` strips out anything dangerous from the filename (e.g. `../../etc/passwd` style path tricks) before it's used to build a path on disk. The file is saved straight into `projects/` (not into a project sub-folder yet — that only happens at `/generate`, once we know the Job ID). `scan_excel_tabs()` (from `matrix_engine.py`) just opens the file and returns the list of sheet names, e.g. `["Store A", "Store B"]`. `render_template` hands the list of tab names to `matrix.html`, which then draws one card per tab (see [Section 11](#11-templateshtml--the-web-pages)).

**`/preview`.** The user has now typed a Start Cell / Job ID Cell / Store Column for every tab they checked. For **every** tab that exists (not just the checked ones — so unchecked tabs still remember their typed values if the user re-visits), this route:
1. Reads the 3 coordinate fields from the form, defaulting to `B8`, `E1`, `A` if blank.
2. Stores them in `user_inputs[tab]` (this dict gets round-tripped back into hidden form fields, so if the user clicks "Update Previews" again, their typed values survive).
3. **Only for checked tabs**, calls `generate_tab_map()` (see [Section 6.2](#62-generate_tab_map--reading-a-blueprint-from-a-tab)) inside a `try/except`. If the coordinates are wrong (e.g. "E1" isn't actually where the Job ID lives), `openpyxl` throws an exception, which gets caught and turned into a friendly `{"error": "..."}` dictionary instead of crashing the whole page.
4. All the successfully-mapped blueprints get serialized to JSON (`blueprints_json`) and stuffed into a **hidden form field** — this is important: Flask does not remember anything between requests by itself (no session used here), so the *entire* blueprint has to be smuggled through the HTML back to the browser and then submitted right back to the server at `/generate`.

**`/generate`.** This is where a **Project Folder actually gets created**:
```python
raw_project_name = request.form.get('project_name', 'Untitled_Project')
safe_project_name = clean_file_name(raw_project_name)
job_id = clean_file_name(blueprints[first_tab].get("raw_job_id", "UNKNOWN"))
time_stamp = datetime.now().strftime("%y%m%d_%H%M")
final_folder_name = f"{safe_project_name}_Job-{job_id}_{time_stamp}"
```
So a folder like `SummerCampaign_Job-12345_260804_1430` is born. Right before the originally-uploaded file is `shutil.move()`d from the loose `projects/` folder into this new sub-folder, `close_if_open_elsewhere(filepath)` ([Section 5](#5-core_mathpy--shared-math--formatting)) is called — if the user opened that exact file via the "Open Excel File" button in `matrix.html` to check it (or fix a duplicate store name) and never closed it, Excel's file lock would otherwise make `shutil.move()` crash with a `PermissionError`. This force-closes that copy first (discarding any unsaved edits in it — by this point the user has already saved what they meant to keep). Then `generate_all_outputs()` (the real engine, [Section 6.3](#63-generate_all_outputs--the-core-generator)) is called to build File 1/2/3. The three resulting filenames are filtered to drop any that are `None` (File 1 and File 3 are only created if at least one pack was selected for signature-code generation — see the `any_packs_selected` logic below) and shown to the user as download buttons.

**`/download/<folder_name>/<filename>`.** A very small, security-conscious route: `os.path.basename()` is applied to both URL parts so a mischievous URL like `/download/../../secrets/file.txt` can't escape the `projects/` folder. If the file exists, `send_file(..., as_attachment=True)` streams it to the browser as a download.

### 4.3 The Dashboard: `/`

This route has no form to process — it just **scans the filesystem** every time it's loaded and builds a list of "project" dictionaries to show on the homepage.

```mermaid
flowchart TD
    A[Loop every folder in projects/] --> B{Folder empty?}
    B -- yes --> C["🧹 Zombie Sweeper:\ntry to delete it, then skip"]
    B -- no --> D[Split files into: Excel / PDF / JSON groups]
    D --> E[Find Job ID by looking for a file\nstarting with 'Packing Sheet_']
    E --> F[Build a dict: name, date, job_id, file lists]
    F --> G[Sort newest-first by folder creation time]
    G --> H[render index.html]
```

The "Zombie Sweeper" comment refers to folders that are empty because Windows/OneDrive sometimes holds a file lock a few seconds after Excel closes, so a delete attempt made moments earlier (e.g. in `/delete`) can silently fail, leaving an empty ghost folder. Every dashboard load gives it another chance to clean itself up via `shutil.rmdir`.

`json_files` are **sorted newest-first by creation time** (`os.path.getctime`) — this feeds the "Create Sub-Group" popup dropdown in `index.html`, so the most recent Stage is offered first as the baseline.

The dashboard also groups **CSV** files (the courier CSV from Packing Labels) into their own "🚚 Courier CSV" list, with the same Open / Download / Delete buttons. `.json` files are never listed, and `….labelmap.json` files are also kept **out** of `json_files`, so they never show up as a Sub-Group baseline. A project that contains a CSV shows a **🧵 Stitch Labels** button (a placeholder until stitching is built) instead of "Create Sub-Group".

### 4.4 Project & File Management: `/delete`, `/delete_file`, `/open_local`, `/open_upload`

- **`/delete/<folder_name>`** — deletes an entire project. Before calling `shutil.rmtree`, it manually walks every file with `os.walk` and calls `os.chmod(file_path, stat.S_IWRITE)` to strip any read-only flag first. This matters because files that were just closed by Excel/xlwings can briefly retain a read-only attribute that would otherwise make `rmtree` fail.
- **`/delete_file/<folder_name>/<filename>`** — same idea, but for a single file (used by the little 🗑️ Delete button next to each file row in the dashboard).
- **`/open_local/<folder_name>/<filename>`** — calls Windows' own `os.startfile(file_path)`, which is the exact same as double-clicking the file in File Explorer — it opens in whatever program Windows has associated with that extension (Excel, Adobe Reader, etc.). Returns an empty `204 No Content` response so the browser's JavaScript `fetch()` call (see [Section 12.2.D](#12-staticscriptjs--the-frontend-brain)) doesn't navigate away from the dashboard.
- **`/open_upload/<filename>`** — the exact same `os.startfile()` trick, but for a file that's still sitting loose in the `projects/` root rather than inside a project sub-folder yet. This exists because `matrix.html` lets the user open the raw distribution file (to inspect it, or fix a duplicate store name) at a point in the wizard — right after upload, before "Generate" — when no project folder has been created yet, so `/open_local`'s `folder_name` segment doesn't apply. Once "Generate" actually runs, the file moves into its project folder and `/open_local` takes over for anything after that point.
- **`/discard_upload` and `/discard_project`** — called by the browser, not by a link, when you **leave a page without generating** (see [Section 12.2](#122-section-by-section-walkthrough), section 0). `/discard_upload` deletes a scanned-but-unused `.xlsx`/`.xls` sitting loose in `projects/`. `/discard_project` deletes a Label Shuffler project folder, but **only if it holds nothing but spreadsheets** (`is_abandoned_project()`): every real project has a generated `.pdf` or `.json`, so picking an existing project and then leaving never deletes it. Both use `os.path.basename()` and reject `.`/`..`, so they can't reach outside `projects/`.
- **`force_delete_upload(file_path)`** — the deletion helper behind both: it calls `close_if_open_elsewhere()` ([Section 5](#5-core_mathpy--shared-math--formatting)) to close the file in Excel **without saving**, retries the delete up to 10 times 0.3 s apart while Windows releases its lock, then removes Excel's `~$` lock file.

| Tool | Left behind before Generate | Discarded on leave |
|---|---|---|
| Packing Labels | the scanned Excel file in `projects/` | the file |
| Distribution Mapper (Packing Sheet) | the scanned Excel file in `projects/` | the file |
| Label Shuffler | a project folder holding the Signature links file (created at step 1) | the whole folder, only if it holds nothing but spreadsheets |

Submitting the page's own forms (Preview, Generate, Recheck) and "Open Excel File" don't count as leaving; pressing Back keeps the file. The startup sweep in [4.1](#41-startup) catches anything the browser couldn't report (a crash, a power cut).

### 4.5 The Sub-Group route: `/subgroup/<project_name>`

```mermaid
flowchart TD
    A[GET: user opened the modal on the dashboard] --> B[Load metadata JSON\ndefault to newest file if none specified]
    B --> C[Render sub-group.html\n JS builds the form dynamically from metadata]
    C --> D[POST: user submits chosen tabs/packs/item ranges]
    D --> E["Build subgroup_instructions dict\n{tab: {item_row, packs: {pack: [[start,end],...]}}}"]
    E --> F[execute_subgroups]
    F --> G{SubgroupValidationError?}
    G -- yes --> H[Re-render sub-group.html\nwith error banner, NOTHING saved]
    G -- no --> I[redirect to dashboard]
```

The `target_json` query parameter (set by the dropdown in the dashboard's "Create Sub-Group" modal — see `index.html`) tells this route **which metadata file to treat as the baseline** — e.g. the original `project_metadata.json`, or an existing `Stage - 1 - ...json` if you're layering a second round of sub-grouping on top of the first. If none is specified, it just grabs the first `.json` file it finds in the folder.

On POST, the form data is parsed into a nested dictionary shaped exactly like what `subgroup_engine.execute_subgroups()` expects:
```python
subgroup_instructions = {
    "Store A": {
        "item_row": 9,
        "packs": {
            "Banner Pack": [[1, 5], [6, 10]]   # two sub-groups: items 1-5, and 6-10
        }
    }
}
```
This is exactly the *very same shape* used inside `subgroup_engine.py` — worth remembering if you ever need to trace a bug between the two files.

```python
try:
    execute_subgroups(project_dir, metadata, subgroup_instructions)
except SubgroupValidationError as e:
    return render_template('sub-group.html', ..., error=str(e))
return redirect(url_for('dashboard'))
```
This `try/except` is the fix we added recently: if the item numbers the user typed don't actually exist in the sheet (a typo, or duplicate item numbers), `execute_subgroups` raises `SubgroupValidationError` **before saving anything**, and the route shows the error message right there on the page instead of silently redirecting to the dashboard as if it had worked. See [Section 7](#7-subgroup_enginepy--the-sub-group-engine) for the full mechanics.

### 4.6 The PDF Label Shuffler route: `/pdf`

This is the biggest single route in the file because it handles **three different steps of one wizard inside one function**, distinguished by a hidden `step` form field.

```mermaid
flowchart TD
    A["GET /pdf"] --> B[Scan projects/ for folders that already have\na 'Signature links...' Excel file]
    B --> C[Render Step 1: pick existing project OR type a new name]

    C --> D["POST step=1"]
    D --> E[Create/locate the project folder]
    E --> F[Read every tab of the Signature Links Excel with pandas]
    F --> G{Duplicate store names\nin column A of any tab?}
    G -- yes --> H["Nothing is deleted.\nRe-render Step 1 with the error modal:\nOpen Excel File / Recheck File / Close"]
    H -.->|"user fixes + saves,\nclicks Recheck File"| D
    G -- no --> I[Render Step 2: one dropzone per Tab × Pack]

    I --> J["POST step=2"]
    J --> K[Loop every uploaded PDF file field]
    K --> L[Build a store→code mapping for that Tab/Pack\nfrom the Excel columns]
    L --> M["process_and_shuffle_pdf() (pdf_engine.py)"]
    M --> N[Render Step 3: download links]
```

**Step 1.** Two ways to pick a project: the `existing_project` dropdown, or typing a `new_project` name (a timestamp gets appended to new names to keep them unique). Whichever Excel file is in play — either freshly uploaded, or an existing `"Signature Links..."` file already sitting in that project's folder (found by `.startswith("signature links")`, case-insensitive) — gets opened with `pd.ExcelFile`. For **every sheet**, `df.columns[1:]` (everything except the first column, which holds store names) becomes the list of "Packs" shown in Step 2.

The duplicate-store check here matters a lot for the PDF engine: `process_and_shuffle_pdf` matches a label to a store purely by **searching for the store's name inside the label's text** — if two stores in the same tab share a name, the shuffler can't reliably tell them apart, so this route refuses to continue.

**Nothing gets deleted on a duplicate.** Earlier versions of this route `shutil.rmtree`'d the project folder the moment a duplicate was found — which, for an *existing* project, meant a bad re-upload could wipe out everything already generated for that job. Now the folder and the uploaded Excel are left exactly where they are, and the duplicate modal (`⚠️ Duplicate Store Names Found` — see [Section 11.3](#113-pdfhtml)) gives the user two extra buttons instead of just a dead end:
- **"Open Excel File"** — an `<a>` styled with `class="btn-file open"` pointing at `/open_local/<project>/<filename>`, intercepted by the same generic fetch-based opener described in [Section 12.2.D](#12-staticscriptjs--the-frontend-brain) — no page navigation, the file just pops open in Excel.
- **"Recheck File"** — a tiny second `<form>` inside the modal that resubmits `step=1` with the *same* `existing_project` and a `resume_filename` hidden field carrying the exact filename that was flagged. On the server side, Scenario B of Step 1 checks `resume_filename` first (an exact match against a file already in that project folder) before falling back to its usual `"signature links"`-prefix search — so recheck works even for a brand-new project whose file was never named that way. If the duplicates are gone, Step 1 falls straight through to Step 2 with zero extra clicks; if not, the same modal reappears.

This only catches *exact* duplicate names, though. Two *different but similar* names (e.g. `"Northlands"` and `"Northlands NZ"`) pass this check fine, but can still confuse the substring matcher at PDF-processing time — see the collision detection described in [Section 8.4](#84-build_audit_report-and-build_divider_sheet), and the separate same-store-matched-twice detection in [Section 8.2](#82-process_standard_layout).

**Step 2.** The form field names carry structured information inside their *name attribute itself*, using `---` as a separator (chosen specifically because pack/tab names might contain underscores):
```html
<input type="file" name="pdf---{{ tab }}---{{ pack }}">
<input type="checkbox" name="divider---{{ tab }}---{{ pack }}">
```
So the Python side loops `request.files.items()`, and for every key that starts with `pdf---`, splits it back apart:
```python
parts = key.split('---')            # ["pdf", "Store A", "Banner Pack"]
tab_name, pack_name = parts[1], parts[2]
```
For that Tab/Pack, it rebuilds a `store_mapping` dict (`{"Store 12": "A", "Store 47": "B", ...}`) straight from the Excel columns using `df.iterrows()`, then calls `process_and_shuffle_pdf()` — the actual PDF engine ([Section 8](#8-pdf_enginepy--the-label-shuffler)) — once per uploaded PDF. Every processed file is renamed with a `_Shuffled` suffix and saved directly into the project folder.

**Step 3** is just a static "success" screen listing the generated filenames as download links.

**GET (no step)** is what runs when you first land on `/pdf` (or click "Reset") — it rebuilds the `existing_projects` dropdown by scanning `projects/` for folders that already contain a Signature Links file.

### 4.7 The Packing Labels route: `/packing-labels`

One route, like `/pdf`, handling several steps of a wizard inside one function — here distinguished by which **submit button** was pressed (`'preview' in request.form` / `'generate' in request.form`) rather than a hidden `step` field.

```mermaid
flowchart TD
    A["GET /packing-labels"] --> B[Upload form]
    B --> C["POST with a file"]
    C --> D[Save to projects/, convert .xls, list tabs]
    D --> E["POST preview"]
    E --> F["parse_packing_data() per tab
(runs every check)"]
    F --> G{PackCheckError?}
    G -- yes --> H[Issue table per tab
Generate locked]
    G -- no --> I["build_consignments()
consignment table + service codes"]
    I --> J["POST generate"]
    J --> K["Validate fields + service codes,
build consignments FIRST"]
    K --> L[Create project folder, move Excel]
    L --> M["generate_packing_labels() → PDF"]
    M --> N["write_open360_csv() + write_label_map()
record_service_usage()"]
    N --> O[Success page: PDF + CSV downloads]
```

- **Preview** parses each selected tab. A `PackCheckError` becomes that tab's issue table ([Section 9.5](#95-checks-errors-and-warnings)); any other exception becomes a one-line issue. Only when every tab is clean are the tabs combined (keys prefixed with the tab name) and passed to `build_consignments()` for the consignment table ([Section 10.6](#106-the-consignment-table-service-codes-and-edits)).
- **Generate** follows "validate first, write last": required fields, edits and service codes are all checked **before** a project folder is created. If anything fails after that, the Excel file is moved back and the folder removed ([Section 10.8](#108-order-of-work-in-generate)).
- If a reload arrives after the upload was already discarded, the route answers "The uploaded file is no longer available. Scan it again." instead of crashing.
- `_consignment_edits(form)` safely decodes the hidden `consignment_edits` JSON; anything that isn't a dictionary is ignored. The same field also carries `__offered__`, the Excel addresses the address book has already been asked about ([10.10](#1010-the-address-book)).

### 4.8 The Address Book routes: `/address-book` and `/api/addresses`

| Route | Does |
|---|---|
| `GET /address-book` | The Address Book page (rows are loaded by JavaScript) |
| `GET /api/addresses?q=&limit=&offset=` | Search; empty `q` lists the most recently used. Returns `{rows, has_more}` |
| `POST /api/addresses` | Add an address (JSON body) |
| `PUT /api/addresses/<id>` | Edit an address |
| `DELETE /api/addresses/<id>` | Delete an address and the spellings learned for it |
| `GET/POST /address-book/update` | **Update in bulk**: update the book from the courier portal's CSV of verified addresses — upload → review → apply ([10.11](#1011-updating-the-book-from-the-courier-portal)) |

Problems the user can fix (missing fields, a duplicate) come back as `400 {"error": "..."}` and are shown in the edit row.

---

## 5. `core_math.py` — Shared Math & Formatting

This file has **no Flask, no Excel-opening code, no file I/O for the workbooks themselves** — it's pure logic, which is exactly why both `matrix_engine.py` and `subgroup_engine.py` import from it instead of duplicating this code.

### `clean_file_name(raw_string)`
Turns messy free-text (a Job ID cell value, a user-typed project name) into something safe to use as a filename or folder name: collapses repeated whitespace, then strips out every character except letters, digits, spaces, underscores and hyphens.
```python
clean_file_name("Job #12345 / Campaign!!") → "Job 12345  Campaign"
```

### `sanitize_cell(val)`
Reads one raw Excel cell value and decides what it "really" means for signature-matching purposes:
- `None` → `0`
- already a number → returned as-is
- text like `"-"`, `"."`, `"0"`, `""` → treated as `0` (these are common "nothing ordered" placeholders in warehouse spreadsheets)
- text that *looks* like a number (`"12.0"`) → converted to `int`/`float`
- anything else (e.g. `"J468791-01"`) → kept as the raw text string

This function is the reason the README says *"Dashes (-) and periods (.) are safely ignored as zero."*

### `sort_key(sig)`
A **signature** is a tuple of cell values across one store's row, e.g. `(5, 0, "Banners")`. To sort a list of these consistently (numbers before text, zeros last) without Python crashing on "can't compare `int` to `str`," every value is converted into a 3-part tuple:
```python
0        → (3, 0, "")            # zeros always sort last
5        → (1, 5, "")            # numbers sort by their value
"Banner" → (2, 0, "banner")      # text sorts alphabetically, after numbers
```
`unique_sigs.sort(key=sort_key)` then works because Python can always compare tuples of the same shape.

### `generate_pack_signatures(raw_values, store_rows, p_start, p_end)` — the heart of the whole app
This is what turns raw quantities into letter codes. Given a 2D grid of cell values (`raw_values`, straight from Excel), a range of rows (`store_rows`), and a column range (`p_start` to `p_end`):

```mermaid
flowchart TD
    A[For every store row] --> B["Build a tuple of sanitized cell values\n across this pack's columns"]
    B --> C{All values are 0?}
    C -- yes --> D["Signature = None\n(this store ordered nothing from this pack)"]
    C -- no --> E["Signature = the tuple itself"]
    D --> F[Collect all signatures]
    E --> F
    F --> G["Find the SET of unique signatures\n(duplicates collapse to one)"]
    G --> H[Sort them consistently]
    H --> I["Assign letters: A, B, C ... Z, AA, AB ..."]
    I --> J["Count how many stores share each signature"]
    J --> K[Return: per-row signatures, unique list, letter map, counts, ordered letter-per-row list]
```

The letter-assignment line is worth reading closely:
```python
letter = chr(65 + index) if index < 26 else chr(65 + (index // 26) - 1) + chr(65 + (index % 26))
```
`chr(65)` is `'A'`. For the first 26 unique signatures (`index` 0–25) it's a single letter `A`...`Z`. Past that, it builds a two-letter code the same way spreadsheet columns go `Z, AA, AB, ...` — `index=26` gives `AA`, `index=27` gives `AB`, and so on.

The function returns **five things at once** (Python lets a function return a tuple and the caller unpacks it):
1. `row_signatures` — `[(row_number, signature_or_None), ...]` for every store
2. `unique_sigs` — the sorted list of distinct signatures found
3. `sig_to_letter` — `{signature: "A", ...}`
4. `summary_counts` — `{signature: how_many_stores_have_it, ...}`
5. `ordered_codes` — just the letters, in store order (`""` where a store ordered nothing) — this list is exactly what later gets exported into "File 3" for the PDF engine to consume.

### `format_file1(...)` / `format_file2(...)`
Pure Excel-styling helpers using `xlwings` range objects. They don't calculate anything — they set colors, borders, fonts, and alignment on cells that have *already* been written. `format_file1` styles the single newly-inserted signature column inside the working copy; `format_file2` styles the whole summary block (quantities + count + letter) with alternating row-band colors (`r_offset % 2 == 1` → light blue) in the Packing Sheet. `-4108` is not a typo — it's the literal integer value of Excel's `xlCenter` constant, used because `xlwings` calls straight into Excel's underlying COM API, which doesn't know Python's friendlier constant names.

### `build_initial_metadata(...)`
Builds the **very first** metadata block for one tab, right after the Matrix Engine runs. Its trickiest part is tracking column shift as it scans packs left-to-right:
```python
current_shift = 0
for pack in tab_info["pack_ranges"]:
    if any_packs_selected and is_selected:
        final_start = pack["start"] + current_shift + 1   # +1 for the new code column about to be inserted
        current_shift += 1
    else:
        final_start = pack["start"] + current_shift        # unaffected packs still drift right if an earlier pack got a column inserted
```
Every time a pack to the *left* gets a new signature column inserted, every pack to its *right* shifts one column further right in the real spreadsheet — this loop keeps the metadata's column numbers in sync with that shift so future code (like the Sub-Group Engine) can trust the numbers.

### `update_metadata_for_subgroup(...)`
The same shifting idea, but triggered by the *Sub-Group Engine* inserting a column. It walks every existing pack (and every sub-group already nested inside a pack) and nudges anything that sits at-or-after the insertion point one column to the right, then registers the brand-new sub-group's own coordinates under its parent pack's `"sub_groups"` dictionary. This is what makes sub-grouping **infinitely repeatable** — the metadata always reflects the spreadsheet's true current shape.

### `get_next_stage_filenames(...)` / `get_available_project_files(...)`
Small filesystem helpers. The first parses a filename like `"Stage - 2 - Final Book1.xlsx"` with a regex to produce `"Stage - 3 - Final Book1.xlsx"` — note that `subgroup_engine.py` actually contains its own (near-identical) inline version of this numbering logic and doesn't end up calling this particular function; it's kept here as a small reusable utility. The second just lists `.json` files in a project folder, sorted alphabetically (which conveniently also sorts them "Stage 1, Stage 2, Stage 3..." since that's how the filenames are constructed).

### `close_if_open_elsewhere(file_path)`
Called by **every Generate** — the Distribution Mapper's `/generate`, the Label Shuffler's Step 2, Packing Labels' Generate (first thing, before the sheet is even read) — plus `subgroup_engine.py` and the abandoned-upload cleanup: anywhere the app is about to read, grab exclusive `xlwings` control of, move or delete a file that the user might have open in Excel (e.g. via one of the "Open Excel File" buttons).

It finds **every copy of that exact file open anywhere** and closes it **without saving**:

```mermaid
flowchart TD
    A["_excel_windows(): walk every Excel window
(visible or hidden, every Excel process)"] --> B{"Protected View window?
(has SourcePath/SourceName)"}
    B -- yes --> C{"Same file?"}
    C -- yes --> D["window.Close()"]
    B -- no --> E["window.Application → that Excel instance
(each instance handled once)"]
    E --> F["Every open workbook with the same full path
→ Close(SaveChanges=False)"]
    E --> G["Every Protected View copy of it
→ Close()"]
    D & F & G --> H["_wait_until_released(): wait (up to 5 s)
until Windows lets go of the file"]
```

- **Same file only.** Paths are compared in full (`os.path.normcase(os.path.normpath(...))`), so other workbooks — including one with the same *name* in a different folder — are left open.
- **Protected View** (the yellow "Enable Editing" bar Excel shows for downloaded files) isn't a normal workbook, so it's handled separately. This matters for more than closing: xlwings' own list of Excel instances (`xw.apps`) looks at one window per Excel process and **crashes** if that window is a Protected View window. That's why `_excel_windows()` walks the windows itself (borrowing xlwings' `accessible_object_from_window` to reach each one's COM object) and skips anything it can't read instead of failing.
- **Busy Excel.** Excel refuses COM calls while it's busy, e.g. while someone is typing in a cell. `_close_with_retry()` tries up to 6 times, 0.5 s apart.
- **Lock release.** Windows can hold the file for a moment after Excel closes it, so `_wait_until_released()` waits until the file can be opened for writing before returning — the move or delete that follows then works first time.
- Returns how many copies it closed; with nothing open it simply returns 0.

This is deliberately blunt: by the time it runs, the user has either already saved what they meant to keep, or is about to have the app work on the saved file anyway, so a silently-discarded stray edit is a smaller risk than the alternative (Excel's file lock making the whole operation crash with a `PermissionError`, or the app reading a stale copy).

---

## 6. `matrix_engine.py` — The Matrix Engine

### 6.1 `scan_excel_tabs(file_path)`
Three lines. Opens the file with pandas just long enough to read `.sheet_names`, then closes it immediately. This is deliberately the *cheapest possible* way to answer "what tabs does this file have?" — no cell data is read at all.

### 6.2 `generate_tab_map(...)` — reading a "Blueprint" from a tab

This function is what the `/preview` route calls. Given a tab and the 3 coordinates the user typed (Start Cell, Job ID Cell, Store Column), it has to **reverse-engineer the entire layout** of a spreadsheet it's never seen before, using only those 3 anchor points.

```mermaid
flowchart TD
    A["Start Cell (e.g. B8)"] --> B["Column letter → column number.\nRow number → this is the 'Pack Group Row'"]
    C["Job ID Cell (e.g. E1)"] --> D["Column number → 'Pack Start Column'.\nRow number → 'Job ID Row'"]
    E["Store Column (e.g. A)"] --> F[Read raw Job ID value at that cell]
    B & D --> G["Scan right from Job ID row\nuntil an empty cell → true last column"]
    G --> H["Scan up from the bottom in the Store column\nuntil a non-empty cell → last store row"]
    H --> I{"Last store row says 'total'?"}
    I -- yes --> J[Step back one more row]
    I -- no --> K[Keep as-is]
    J & K --> L["Walk every store name → detect duplicates"]
    L --> M["Walk columns at the Pack Group Row,\ndetecting MERGED cells → one Pack per merge,\nor a single column → one Pack per column"]
    M --> N["Return a 'blueprint' dict:\nstocks_map, packages_map, warnings, + raw backend_data"]
```

A few specific tricks worth calling out:
- **Finding the true last column** doesn't trust `sheet.max_column` (Excel/openpyxl can report phantom extra columns from old formatting) — it walks backward from the reported max column until it finds one with an actual value in the Job ID row.
- **Detecting "Total" rows:** if the very last non-blank cell in the Store column contains the word "total" (case-insensitive), that row is excluded from the store count — it's a summary row, not a real store.
- **Detecting packs** checks `sheet.merged_cells.ranges` first. If the current column falls inside a merged range at the pack-group row, that whole merged block becomes one pack (its name comes from the merge's top-left cell) and the scan jumps past it. If it's *not* merged, it falls back to treating a single non-empty header cell as a one-column pack.
- **Duplicate store detection** here is purely informational (shown as a yellow warning banner in `matrix.html`) — unlike the harder duplicate check in the `/pdf` route (Step 1), it does **not** block generation. It's a heads-up, not a hard stop.

### 6.3 `generate_all_outputs(...)` — the core generator

This is the single most important function in the Matrix Engine — the one that actually writes files to disk. It happens in three phases.

```mermaid
flowchart TD
    subgraph Setup
        A[Copy the uploaded file → File1 path + File2 path]
        B["any_packs_selected? (checked across ALL tabs)"]
    end
    subgraph "Phase 1 & 2 — same loop, same xlwings session"
        C[Open File1 + File2 invisibly in Excel]
        D[For each selected tab]
        D --> E["build_initial_metadata() → project_metadata"]
        E --> F["For each pack, RIGHT TO LEFT:\ngenerate_pack_signatures()"]
        F --> G{Pack is selected?}
        G -- yes --> H["Insert 1 new column in File1\n+ batch-write letter codes\n+ format_file1()"]
        G -- no --> I[Nothing written to File1]
        H & I --> J["Build the Count+Code summary block\nfor File2 (all packs, selected or not)"]
        J --> JN["Add a 'Job numbers' tab to File2\n(write_job_numbers_sheet)"]
        JN --> K[Add a 'Packaging Stocks' summary tab to File2]
        K --> L[Save + close File1 and File2]
    end
    subgraph "Phase 3 — pandas only, no Excel needed"
        M{any_packs_selected?}
        M -- yes --> N["Rebuild a MINIMAL workbook (File3):\njust Store Name + one 'Code for X' column per selected pack"]
        M -- no --> O[Skip File3 entirely]
    end
    N & O --> P[Write project_metadata.json to disk]
```

**Why right-to-left?** (`for pack in reversed(tab_info["pack_ranges"])`) Inserting a new column shifts every column *after* it one position to the right. If you inserted left-to-right, every subsequent pack's remembered column numbers would instantly go stale. Processing packs from the rightmost one backward means each insertion only ever affects columns you've *already finished with* — nothing still-to-be-processed ever moves.

**File 1's store names get cleaned too, not just File 3's.** Right after `sheet1_xw` is opened for a tab (before any pack-column insertions happen), the store-name column is read as a batch, run through `clean_store_name()` (`core_math.py` — strips a trailing period, collapses double-spaces), and written straight back:
```python
raw_store_values = store_range.value
store_range.value = [[clean_store_name(v)] for v in raw_store_values]
```
Historically only Phase 3 (File 3) applied this cleaning, since it rebuilds its "Store Name" column from scratch anyway — File 1 just carried over whatever raw text was in the original spreadsheet untouched. That meant the same store could read differently between the two output files (e.g. a trailing "." surviving in File 1 but not File 3). Cleaning happens over the identical row range/column File 3 later uses, so both stay in sync. This write happens safely *before* any pack columns are inserted for that tab, since insertions only ever happen to columns at-or-right-of a pack's start — never left of the store column — so `store_col_idx` is still valid at this point.

**Why `any_packs_selected` matters so much:** if the user didn't check *any* pack checkboxes in the preview screen, there's nothing to calculate signatures for — File 1 (the working copy with new signature columns) and File 3 (the signature lookup table) become pointless, so both are skipped (`file1_name = None`, `file3_path = None`). Only File 2, the plain Packing Sheet, still gets produced in that case — that's a legitimate use case (someone who just wants a cleaned-up copy of their distribution list).

**The "batch write" pattern**, seen here and repeated in `subgroup_engine.py`, is a deliberate performance choice:
```python
batch_data = []
for r, sig in row_signatures:
    batch_data.append([sig_to_letter[sig]] if sig is not None else [None])
sheet1_xw.range(f"{col_letter_start}{start_r}").value = batch_data
```
Writing to Excel cell-by-cell through `xlwings` is slow (each write is a round-trip to the real Excel application). Building the whole column as a Python list first, then writing it in **one** `.value = batch_data` assignment, can turn thousands of slow round-trips into a single fast one.

**The "Job numbers" tab.** While each tab's `raw_values` is read (before File 2's rows/columns are touched), every pack's columns are scanned along the **Job ID row**, and the job numbers found are kept in first-seen order with repeats inside that same pack dropped. Each pack becomes one entry in `master_job_data` (`{"header": "Tab | Pack X", "jobs": [...]}`). After all tabs are done, `write_job_numbers_sheet()` writes one block per pack (bold header, then the job numbers centered underneath), with a thin black divider row between blocks, the same idea as Packaging Stocks. Inside each block, job numbers are grouped by **series**: a job number is `series-index` (e.g. `J476523-01` → series `J476523`, index `01`), split on the first `-` by `split_job_number()`. A job number with no `-` is series-only. `group_jobs_by_series()` keeps the series in the order they first appear and sorts each series' jobs by index: series-only first, then numeric indexes in number order. Each series gets a grey, italic `Series J476523` sub-heading. The full job number stays in its own cell on each row so that cell can carry a link later. It also builds `job_to_groups` (`{job_number: [every pack header it appears in]}`). Any job number that appears in **more than one** pack group, whether in the same tab or a different one, is filled light red, and column B next to it lists the other group(s): `⚠ Shared with: Other Tab | Pack 2`. The same check, `find_shared_job_numbers()`, is also returned as the 4th value of `generate_all_outputs()` so `/generate` can list those job numbers in a warning on the results page. Generation still goes ahead.

**Installer tabs and the "Divider Barcodes" tab (File 3).** A tab ticked as "sent to an installer" (`user_inputs[tab]["installer"]`, set in `/generate`, which also selects every pack in that tab) gets one row per job number per code in an extra `Divider Barcodes` tab of the Signature Links file: `Tab | Pack | Code | Job Number | Kind | Qty | Barcode`. `tab_column_jobs` records which job number sits in each pack column. `assign_job_kinds()` numbers the kinds: a job number found in more than one column (same pack, another pack or another tab) gets Kind 1, 2, 3... in reading order, first tab to last and left to right within each tab. A job number in only one column has no kind. Kinds are counted across every selected tab, installer or not. `build_divider_barcode_rows()` then walks each code's signature: every column with a quantity and a job number becomes a row, and `Qty` is that store's quantity. `Barcode` is the exact text the divider's barcode encodes: the job number, plus ` Kind N` when it has more than one kind (e.g. `J476699-17 K1`), so a scan tells kinds apart. The tab is only written when at least one tab is an installer tab, so File 3 is otherwise unchanged.

**File 2's "delete rows below the data" step** (`sheet2.range(f"{start_del}:1048576").api.EntireRow.Delete()`) clears out anything left below the actual store list (old totals, stray notes) before the new Count/Code summary columns get written — otherwise leftover junk rows could visually collide with the new summary block.

**Phase 3's "reversed" pack summaries** — `for p_sum in reversed(tab_summaries.get(tab_name, []))` — exists purely so the resulting Signature Links columns appear in the same left-to-right pack order as the original sheet (since the summaries list itself was built during the reversed right-to-left insertion loop in Phase 1/2, reversing it a second time un-does that and restores original order).

Finally, `project_metadata` — built up tab-by-tab across this whole function — gets written once at the very end as `<File1 name>.json` inside the project folder. This is the file that `subgroup_engine.py` will later read back in.

---

## 7. `subgroup_engine.py` — The Sub-Group Engine

**Purpose:** take an *existing* Packing Sheet + working file (already produced by the Matrix Engine, or by a previous sub-group run) and slice one pack into smaller **item-number ranges**, e.g. turn "Banner Pack (items 1–20)" into "items 1–5" and "items 6–20" as two separately-coded sub-groups — without destroying the original file (it always writes a **new** `Stage - N -` file).

### 7.1 The validation guard (this is the bug fix from earlier in this project)

Before this existed, the code silently trusted whatever "Item Number Row" and Start/End item numbers the user typed. If a typo meant an item number didn't exist in that row, the old code just printed a `[WARNING]` to the terminal and quietly skipped that sub-group — the tool would report success even though part of the work never happened. Two small pieces fix that:

```python
class SubgroupValidationError(Exception):
    """Raised when the user-provided item row/numbers can't be safely resolved against the sheet."""
```
A **custom exception type**. This matters because `app.py` needs to tell the difference between *"the user's input was bad, show a friendly message"* and *"something genuinely broke, let it crash with a real error."* Catching a specific class (`except SubgroupValidationError`) instead of a generic `Exception` means only the deliberate validation failures get the friendly treatment.

```python
def _map_item_columns(row_data, tab_name, item_row):
    col_map = {}
    duplicates = {}
    for c_idx, val in enumerate(row_data):
        ...
        if item_num in col_map:
            duplicates.setdefault(item_num, [col_map[item_num]]).append(col)
        else:
            col_map[item_num] = col
    if not col_map:
        raise SubgroupValidationError(f"...no item numbers were found in row {item_row}...")
    if duplicates:
        raise SubgroupValidationError(f"...duplicate item numbers found...")
    return col_map
```
This function replaces the old plain dictionary-building loop. Two things now stop the whole process cold, with a message that reaches the browser:
1. **The row has zero item numbers at all** — almost certainly means the user pointed at the wrong row entirely.
2. **The same item number appears twice** in that row — item numbers are supposed to be unique per column, so a duplicate is a strong signal of a typo somewhere in the sheet. `duplicates.setdefault(item_num, [first_col]).append(second_col)` is a compact way to build `{105: [col_3, col_7]}` the first time a repeat is spotted.

### 7.2 The three stages

```mermaid
flowchart TD
    A["User's request:\n{tab: {item_row, packs: {pack: [[start,end], ...]}}}"] --> B["Figure out the next free 'Stage - N -' filename"]
    B --> C["Open File1 (old stage) + File2 (Packing Sheet) in Excel"]
    C --> D["STAGE 1 — for each tab:\nmap item_row → column via _map_item_columns()"]
    D --> E{"Row empty or has duplicates?"}
    E -- yes --> Z["❌ SubgroupValidationError\nNOTHING is saved — files are just closed unsaved"]
    E -- no --> F["For each requested start/end range,\nlook up columns + generate_pack_signatures()"]
    F --> G{"Start/end item # not found?"}
    G -- yes --> Z
    G -- no --> H["STAGE 2 — insert new column(s) into File1,\nRIGHT-TO-LEFT by pack, batch-write the letter codes,\nre-merge + color the pack header"]
    H --> I["STAGE 3 — re-map the item row (columns shifted after Stage 2's inserts),\nbuild a side-by-side matrix in that tab's own SG companion tab of File2"]
    I --> J{"Shifted columns not found?"}
    J -- yes --> Z
    J -- no --> K["Save File1 as the NEW Stage - N file\nSave File2 in place\nWrite the new Stage - N .json metadata"]
```

**Why re-map the item row a second time in Stage 3?** Stage 2 just inserted brand-new columns into File1 (the sub-group's code columns). Every item number that used to sit at, say, column 10 might now be at column 12 if two new columns were inserted to its left. The Stage-1 `item_col_map` is now stale, so Stage 3 re-reads the item row fresh off the *already-modified* worksheet before it can correctly locate where each sub-group's data actually lives for building the File 2 matrix.

**Why nothing gets corrupted when validation fails:** look at the very end of the function —
```python
try:
    ...all three stages...
    wb1.save(new_file1_path)   # ← the ONLY place File1 is saved
    wb2.save()                  # ← the ONLY place File2 is saved
except Exception as e:
    ...
    raise e
finally:
    app.quit()   # closes Excel WITHOUT saving whatever's still open
```
Both `.save()` calls sit at the very bottom of the `try` block, *after* every stage of every tab has finished successfully. If a `SubgroupValidationError` (or anything else) is raised partway through — even mid-way through tab 3 of 5 — execution jumps straight to `except`, which re-raises, and `finally` quits the invisible Excel instance **without ever calling `.save()`**. Whatever partial edits existed only in that temporary Excel session simply vanish; the real files on disk are untouched. This is exactly what makes it safe for `app.py` to say "abort the whole thing" on any validation error.

---

## 8. `pdf_engine.py` — The Label Shuffler

**Purpose:** take a raw PDF full of shipping labels (one per store, in whatever random order they were printed) and rebuild it so labels are grouped by their **signature code**, in the same letter-order the Packing Sheet uses — so a warehouse worker holding the printed stack for "Code C" can hand out labels in the exact sequence the packing sheet lists them.

### 8.1 Layout detection — the entry point

```python
first_page = doc[0]
is_split_layout = page_width < page_height  # taller than wide?
```
The whole engine branches on one measurement: is the PDF page **portrait** (taller than wide)? If yes, it assumes this is a "2 labels stacked on one sheet, meant to be cut in half" layout (`process_split_layout`). If the page is wider than tall, it assumes **one full label per page** (`process_standard_layout`).

```mermaid
flowchart TD
    A[Open the PDF, measure page 1] --> B{"page is taller than wide?"}
    B -- yes --> C["process_split_layout()\ntreat page as TOP half + BOTTOM half"]
    B -- no --> D["process_standard_layout()\ntreat page as ONE label, possibly spanning multiple pages\n+ detect same-store-matched-twice"]
    C --> E[build_audit_report]
    D --> E
    E --> F{add_dividers?}
    F -- yes --> G["Insert a divider sheet before each code group\n(green border if the store count checks out, red if not)\nplus one before the Unmatched Pages group"]
    F -- no --> H[Skip]
    G --> I[Save final shuffled PDF]
    H --> I
```

### 8.2 `process_standard_layout(...)`
For each page, it reads the text inside a specific rectangle (`fitz.Rect(0, 0, 317.5, page_height)` — roughly the left portion of the page, where a store name typically sits) and checks whether **any** known store name appears inside that text (case-insensitive substring search). If a page has *no* text at all, it's filed as "blank." If it has text but no store name matched, it's "unmatched."

The trickiest bit handles **multi-page labels** — a single shipping label that spans, say, "Page 2 / 3":
```python
page_info = re.search(r"page\s+(\d+)\s*/\s*(\d+)", text)
if page_info:
    expected_extra_pages = int(page_info.group(2)) - int(page_info.group(1))
elif expected_extra_pages > 0 and current_store:
    expected_extra_pages -= 1   # this page belongs to the store from the previous page
```
When a matched page also contains text like `"Page 1/3"`, the code knows 2 more pages are coming that *won't* have the store name printed on them again, but should still count as belonging to that same store — so it keeps assigning the current store to pages until that countdown reaches zero.

**Same store matched twice = a real problem, and this is where it's caught.** Every store should have exactly one label instance. A store's label can legitimately span several physical pages (the "Page 1/3, 2/3, 3/3" case above), and its header may even repeat on each of those pages — that's still *one* instance. But a match on the same store **outside** that continuation window is a second, separate instance, and a sign something's wrong (a genuine duplicate print, or a label that actually belongs to a similarly-named store getting mis-attributed):
```python
is_continuation = (match == current_store and expected_extra_pages > 0)
...
if not is_continuation:
    store_instance_count[match] += 1
```
Pages aren't committed to a code bucket as they're read — they're staged first into `store_pages[store]`, and only sorted into `code_buckets` (kept) vs. `unmatched_pages` (discarded) in a second pass, once the full instance count for every store is known:
```python
for store, page_nums in store_pages.items():
    if store_instance_count[store] > 1:
        for p_num in page_nums:
            unmatched_pages.append(doc[p_num])   # ALL of that store's pages, not just the extras
    else:
        confirmed_stores.add(store)
        code_buckets.setdefault(store_mapping[store], []).extend(page_nums)
        code_final_store_counts[store_mapping[store]] += 1
```
If a store shows up as 2+ separate instances, **every** page ever attributed to it is pulled into the normal `unmatched_pages` pile — the exact same pile (and divider) used for pages that never matched anything at all, deliberately, so there's no separate "needs review" concept to learn. Since none of that store's pages can be trusted to be the single "true" copy, `confirmed_stores` (not the raw `found_stores` set) is what feeds the missing-stores list and the collision check below — so a duplicated-and-removed store correctly shows up as missing too, exactly as if it had never been found at all, instead of quietly still counting as "matched" while all its actual pages sit in Unmatched.

Pages are bucketed by their signature code, sorted `by (length of code, code)` — this ordering, `sorted(code_buckets.keys(), key=lambda x: (len(x), x))`, is what makes single letters come before double letters (`A, B, ... Z, AA, AB...`) instead of plain alphabetical sort putting `AA` before `B`. Finally everything is stitched together in this fixed order: **audit report → matched pages (grouped by code, dividers optional) → unmatched pages (with their own divider, if enabled) → blank pages.**

`code_buckets` still holds raw **page numbers**, not stores — a store's legitimate multi-page label contributes several entries to the same bucket. `code_final_store_counts` (a separate `Counter`, built alongside `code_buckets` in the loop above) is the accurate "how many *distinct*, non-duplicated stores actually matched here" figure fed to `build_divider_sheet` — deliberately **not** `analyze_matches`'s own `matched_count`, which is blind to the duplicate-removal step and would still count a since-removed store as matched.

**This entire same-store-twice check is deliberately standard-layout-only.** `process_split_layout` does *not* have any of this logic — a repeat match there is left completely alone, matched normally, same as before this feature existed.

### 8.3 `process_split_layout(...)`
Same idea, but every physical page is treated as **two independent halves** (top/bottom), each matched against store names independently, since two different stores' labels might be stacked on one printed sheet. The stitching step at the end is the more complex part: it builds a `master_stack` of "halves" (some real, some from the audit report, some blank divider pages) and then re-pairs them two-at-a-time into brand new output pages:
```python
halfway_point = math.ceil(total_halves / 2)
top_stack = master_stack[:halfway_point]
bottom_stack = master_stack[halfway_point:]
```
i.e. the first half of the master list becomes everyone's *top* half, the second half of the list becomes everyone's *bottom* half — so pairing `top_stack[i]` with `bottom_stack[i]` for each new page reconstitutes full sheets, but now filled with re-ordered content instead of the original random order. `new_page.show_pdf_page(top_rect, doc, ..., clip=src_clip)` is the actual "paste this rectangle of content from the source PDF onto this rectangle of the new page" operation — the workhorse of the whole rebuilding process.

### 8.4 `build_audit_report(...)` and `build_divider_sheet(...)`
A handful of small "report generator" functions that draw brand-new PDF pages from scratch using `fitz`'s drawing primitives (`draw_rect`, `insert_text`, `insert_textbox`, `draw_line`, `draw_circle`) — no source PDF involved. `build_audit_report` is always inserted as the very first page(s) of the output, and turns red if any stores were missing/unmatched/blank, or green if everything matched perfectly — a floor worker can glance at just page 1 to know if the batch is trustworthy. Its sections, in order, are: header/metadata → Unmatched & Blank counts → **Missing Stores** (always last, described below).

**`group_missing_stores_by_code(store_mapping, found_stores)`** builds the Missing Stores data: every store not in `found_stores` (or, from `process_standard_layout`, not in `confirmed_stores` — see [8.2](#82-process_standard_layout)), bucketed by signature code and ordered with the same `(len(code), code)` key used everywhere else, with stores **within** a code group kept in their original spreadsheet row order (since `store_mapping` is a plain dict and dict iteration order in Python 3.7+ is insertion order, which is the order `app.py`'s `/pdf` Step 2 built it in via `df.iterrows()`).

**Missing Stores is rendered as auto-width columns, not a single vertical list.** With enough missing stores this used to spread confusingly across the "cut & stack" split-layout pagination, interleaved with unrelated divider content. Now it flows like a newspaper: each code group gets its own sub-header (`"CODE A - 19 missing"`) followed by its store names; when a column runs out of vertical room, a new column starts immediately to its right on the *same* page (only spilling to a genuinely new page once no more columns fit the page's width); each column's width is sized independently from only the text actually inside it (measured with `fitz.get_text_length`); and if a code group gets cut off mid-list by a column or page break, its header repeats at the top of the next column as `"CODE A (continued)"` so it's never ambiguous whether a name belongs to the group above or is a fresh one. A small guard (`page_top_y != 40 and ... > bottom_limit`) specifically prevents a pathological case where a page is so full already that a column would fit a header but not even one store name under it — rather than drawing an endless chain of empty "(continued)" headers, it jumps straight to a fresh page.

**Catching the "two similarly-named stores collapse into one" bug.** Because matching is plain substring search, a store like `"Northlands"` can silently swallow labels that actually belong to `"Northlands NZ"` if the printed text doesn't include the distinguishing suffix — both labels get matched to `"Northlands"`, and `"Northlands NZ"` shows up as a false "missing store." Two helper functions guard against this:
- `find_name_collisions(store_names)` — scans every pair of store names in the signature-links mapping and flags any pair where one name is fully contained inside another (case-insensitive), e.g. `"Northlands"` inside `"Northlands NZ"`.
- `analyze_matches(store_mapping, sorted_stores, found_stores)` — for each such colliding pair, checks whether *one* was matched while the *other* was not (both-found or both-missing isn't suspicious). When that happens it just marks both codes' `collision_note` flag `True` — **no store names or explanation reach the audit report at all anymore.** The reasoning: the missing store already shows up on the Missing Stores list, and any mis-attributed labels already show up in Unmatched, so a named callout would just repeat information that's already there in more detail, and risked confusing warehouse floor staff who might see the printed page (name-redundancy checking is the tool operator's job). The only visible effect of a collision now is that affected code's divider sheet flipping to a red border with a generic disclaimer. This same function also computes, per code, how many *distinct stores* actually matched versus how many the signature links say should be there.

`build_divider_sheet(page_width, page_height, signature_code, matched_count, expected_count, collision_note)` draws the separator sheet inserted before each code group when `add_dividers=True` is checked in the UI. `matched_count`/`expected_count` are **store** counts (for `process_standard_layout`, the accurate post-dedup `code_final_store_counts`, not `analyze_matches`'s own figure — see [8.2](#82-process_standard_layout)); `collision_note` is just a bool now, not text:
- **Green border** — `matched_count == expected_count` and `collision_note` is falsy. One line, no fraction: `"Matched Labels: {matched_count}"`.
- **Red border** — counts don't match (adds `"matched/expected"` to the same line, plus a "Short by N store(s)" / "N extra store(s)" disclaimer), **or** `collision_note` is `True` (generic "Possible mismatch - please verify" disclaimer, no names) — even if the count happens to check out, because the "extra" label absorbed into this bucket might be hiding behind an otherwise-correct-looking number. No missing-store names or collision reasoning are ever drawn on a divider — that detail lives only on the audit report (or, for the same-store-twice case in standard layout, isn't shown as text anywhere at all — see [8.2](#82-process_standard_layout)).

**Installer barcodes.** When `/pdf` Step 2 finds a `Divider Barcodes` tab in the Signature Links file (read by `core_math.read_divider_barcodes()`), it passes that pack's `{code: [entries]}` to `process_and_shuffle_pdf(divider_barcodes=...)`. `build_divider_sheet(..., barcode_entries=...)` then hands off to `_build_barcode_divider_sheet()`: same border, code and matched count, with a Code 128 barcode per job number underneath, captioned `J476699-09 K1 x 1` (`barcode_caption()`). `code128_modules()` is a small built-in Code 128 (code set B) encoder, so no extra package is needed, and `_draw_code128()` draws the bars as vector rectangles. `_plan_barcode_grid()` picks the column count that keeps the bars widest, and `_paginate_barcodes()` moves barcodes onto extra divider pages rather than shrinking them below a scannable size (`BARCODE_MIN_MODULE`). The split layout adds one half per divider page. With no barcode entries, `build_divider_sheet()` is exactly as before.

A sibling function, `build_unmatched_divider_sheet(page_width, page_height, count)`, draws the same style of separator (always red) right before the unmatched-pages group, headed `"UNMATCHED PAGES / Total Pages: {count}"`.

Since `add_dividers` pages can be as short as half a physical page (split layout) or as narrow as a small label, all of the text above is drawn through a shared `_fit_textbox()` helper that shrinks the font until it actually fits the box — PyMuPDF silently draws *nothing* if a fixed font size doesn't fit, so this is what keeps a divider page from ever rendering blank on an unusually small page size.

### 8.5 `process_and_shuffle_pdf(...)` — the master orchestrator
The only function `app.py` actually calls. It measures the page, decides which of the two engines above to run, then does a final cleanup: `final_shuffled_doc.set_page_labels([])` strips any leftover page-numbering metadata from the source PDF (so the new document doesn't confusingly display the *original* file's page numbers), then saves to `output_pdf_path`.

---

## 9. `packing_label_generator.py` — Packing Labels

Shown in the app as **📦 Packing Labels** — the *Label Maker for Vertical Distribution*.

**Purpose:** read a packing-spec spreadsheet, check it thoroughly, and draw one A4-landscape packing label per box. Each label has a space for the courier label, the box's Packing Spec, the store and address, and one rounded box per item (image, description, job number, quantity…).

### 9.1 What the user does

1. **Scan File** — upload an `.xlsx` or `.xls` file. The tool lists its tabs.
2. **Auto-Map & Update Previews** — choose tabs and each tab's header row. The tool reads the data and runs every check in [9.5](#95-checks-errors-and-warnings). Errors lock generation; warnings are shown but don't block it. If everything is clean, the consignment table ([10.6](#106-the-consignment-table-service-codes-and-edits)) appears.
3. **Customise Cell Layout** — drag blocks to set their order inside each item box. Blocks for columns the file doesn't have show as empty slots and take no space on the label.
4. **Create Project & Generate PDF** — saves the Excel file, the PDF, the courier CSV and the label map into a new project folder.

### 9.2 Columns it recognises

`_header_field()` lower-cases each header cell in the chosen row and tests it against keywords **in this order**; the first match decides the column:

| Data | Header contains | Required |
|---|---|---|
| Packing Spec | `packing spec` | **Yes**, on every row |
| Job Number | `job` | **Yes**, on every row |
| Quantity | `qty` or `quantity` | No (defaults to 1) |
| Description | `desc` | No |
| Thumbnail | `thumb`, `image`, `picture`, or the word `art` / `artwork` (not "Start Date", "Department") | No |
| Dimensions | `dimension`, or `width` / `height` (`w` / `h`) | No |
| Address line 1 / 2 | `address` or `street`; line 2 when it says `line 2` / `address 2` | No |
| Suburb, State, Postcode, Country | `suburb` / `town` / `city`, the word `state`, `postcode` / `post code` / `zip`, `country` | No |
| Material | `material` | No |
| Notes | `note` | No |
| Install | `install` (so `Installer` works too) | **Yes** — `Y`/`Yes`/`True` = installer; `N`, blank or anything else = store |
| Store Name (the receiver) | `store`, `retailer`, `receiver`, `consignee`, `ship to`, `deliver to` | No |

The address parts come **before** the store/receiver words, so `Receiver Address Line 1`, `Store Address`, `Receiver Suburb` or `Store State` are address columns and `Receiver Name` / `Store Name` / `Retailer` / `Consignee` is the store. A header about phone, email, contact or Attn is never taken as the receiver or an address. When two columns match the same field the later one wins, except that a `Store` / `Retailer` column beats a `Receiver`-type one.

Dimensions: a `Dimension` column wins; otherwise `Width x Height`; otherwise whichever exists.

**Column mapping by hand** (in the Data Previews). The columns found by their header names (`detect_columns()`) can be changed in each tab's preview card, in its **🧭 Column mapping** panel. Fields that have a column are **chips**: **green ✓** when the rules found it, **blue ✎** when it was set by hand, grey ⊘ when set to "Not used". Hovering a chip shows the Excel header it's matched to and that header's cell (`Excel header: "Installer?" · cell H2`); clicking it turns it into the list of columns to pick another. Only the fields **still to map** are listed as rows (field → column → cell, three to a line), the ones Generate needs first and in red. Each list offers *Auto* (what the rules found, or "not found"), *Not used*, and every header in the header row (`header_columns()`, read without loading the sheet); **Auto-map all** resets the tab. The choices go with Update Previews and Generate as `colmap::<tab>::<field>` (a column letter, or `-`), `_column_choices()` reads them and `apply_column_choices()` puts them on top of the found columns, so the preview, the checks and Generate all read the file the same way. When a required column can't be found, the error card says "Pick it in the Column mapping below", with the mapping right there. **They survive a refresh:** the page keeps them in this browser per file name and tab (localStorage, `script.js` section 18), and sends them with the first preview when the same file is scanned again.

**Serials and tab dividers on the labels.** Generate passes each pack's serial (`serials`, e.g. `01T1`) and tab (`tabs`) to `generate_packing_labels()`. The serial goes before the store name on every page of the pack, continuation pages and every box of a multi-box pack included (`01T1. Blue Star Moorabbin`). Between the last label of one tab and the first of the next comes a **divider page** ("DIVIDER - NOT A LABEL", `TAB 2`, the tab's name and its serials `01T2 - 14T2`); there's none before the first tab. Dividers stay in the Complete Labels PDF, without a courier label.

### 9.3 Pack groups and merged cells

`parse_packing_data(excel_path, header_row, sheet_name)` walks down from `header_row + 1` until the first empty Packing Spec.

- **Merged cells.** `get_cell_info(row, col)` looks for a merged range covering the cell. If one exists it returns the **top-left value** plus the range's first and last rows; otherwise the cell's own value. So a merged cell counts as its value on **every row it covers** — the fact the address checks rely on.
- **Packing Spec defines the pack group.** The pack id is `ROW_<first row of the Packing Spec merge>_<spec>`, so all rows under one merged Packing Spec land in the same group, and an unmerged cell is a pack on its own, even when the next row has the same text. Several pack groups may share the same Packing Spec text.
- **Install** is a yes/no flag: `Y`, `Yes`, `True` (any capitals) = installer. Anything else, including blank, = store delivery.
- **Address info** is every available address column joined in order (Address Line 1, Line 2, Suburb, State, Postcode, Country). Some files put the whole address in one Address column; that works the same way.

A pack group looks like this:

```python
{
  'pack_spec_name': 'OB1370170170', 'store_name': 'Provision Clayton', 'install': True,
  'address_1': 'Steven Priestley - Wilson Storage, 68 Ricketts Road, ...', 'suburb': '', 'state': '', 'postcode': '',
  'excel_rows': [3], 'items': [{'job_no': 'J477161-54', 'qty': '2', 'thumbnails': [...], ...}],
  'label_no': 1, 'label_total': 1,          # added by assign_label_numbers()
}
```

**Similar addresses** (`similar_addresses()`): two address texts are probably one place written differently when the postcodes match, the street numbers don't contradict (12 vs 14 is another place; a shop number on one side only is fine; Shop 1040 vs Shop 2210 at one address are different shops), and the street and suburb words are close (`12 Smith St` ~ `12 Smith Street` ~ `Shop 1, 12 Smith St` ~ `12 Smyth St`; not `12 Jones St`, not another suburb). A store whose packs give different addresses is a **warning**, not an error ("written differently" when they look like one place, "differs" when they don't). In the consignment table, **one receiver's consignments at different addresses sit next to each other** (`_receiver_groups()` in app.py: the same store or installer name, ignoring capitals, spacing and punctuation, or one small typo apart; across all selected tabs), under a heading that says whether the addresses look like one place written differently or some really differ. Pick the address to **Keep** and **Merge into one**: it goes on all of them as edits (like ✏️) and the preview updates, so they become one consignment. A store's packs at addresses that really differ are a warning too (no longer an error), so they can be sorted out there.

### 9.4 Comparing text: the two keys

Spreadsheets spell the same thing many ways. Instead of comparing raw text, the code compares a **key**: a function that maps every spelling of the same thing to the same string. Two values are "the same" exactly when their keys are equal. This makes grouping a simple dictionary lookup (one pass, no pairwise comparison).

**`_norm(text)` — the basic key**, used for store names, packing specs and receivers:

```python
" ".join(re.sub(r"[^0-9a-z]+", " ", text.lower()).split())
```

1. lower-case everything,
2. replace every run of characters that isn't a letter or digit with one space,
3. split and re-join, which collapses spaces and trims the ends.

```
"12 Bourke Rd."           → "12 bourke rd"
"12  BOURKE RD"           → "12 bourke rd"      (same key → same address)
"Spectacle Hub Curlewis " → "spectacle hub curlewis"
```

**`_address_key(text)` — addresses only.** `_norm` plus street-type equivalence: Street/St, Road/Rd, Avenue/Ave, Boulevard/Blvd, Highway/Hwy, Drive/Dr, Parade/Pde, Place/Pl, Court/Ct, Crescent/Cres, Lane/Ln, Terrace/Tce, Arcade/Arc (`STREET_TYPES`). The mapping is applied to a word **only if the word before it is alphabetic** — i.e. it follows a *name*:

```
"49 Church Street" → words [49, church, street]   "street" follows "church" → st   → "49 church st"
"12 St Kilda Rd"   → words [12, st, kilda, rd]    "st" follows "12" (a number) → unchanged
```

That one-word look-behind is what keeps "Saint" in St Kilda while still treating Street and St as equal.

**`_row_ranges(rows)` — compact row lists** for messages (run-length compression): `[5, 6, 7, 10, 11, 14] → "5-7, 10-11, 14"`.

### 9.5 Checks: errors and warnings

Checks collect **all** problems before reporting, then raise one `PackCheckError` (a `ValueError` subclass, same idea as `SubgroupValidationError` in [Section 7.1](#71-the-validation-guard-this-is-the-bug-fix-from-earlier-in-this-project)) carrying a list of issues:

```python
{'rows': '37-40', 'col': 'N · Install', 'text': 'Install mixed in pack OB1170170170', 'detail': 'Y on 37, 39-40'}
```

`col_ref` turns a column number into `letter · Header` (`N · Install`); the address columns become a range (`F–I · Address`). The preview shows the list as a compact table (rows · column · problem, with a grey detail line). Errors lock Generate; warnings don't.

**Errors — block generation until fixed**

| # | Situation |
|---|---|
| — | No Packing Spec, Job Number or Install column |
| — | A row has a Job Number but no Packing Spec (rows after the first blank Packing Spec would otherwise be skipped silently) |
| — | A row has no Job Number |
| 2 | Rows in one pack have different addresses |
| 3 | Some rows in a pack have an address and others are blank (with no merged cell covering them) |
| 5 | Rows in one pack have different store names, or some are blank |
| 6 | Install is Y on some rows of a pack and not others — one box can't go to both the store and an installer |
| 7 | An installer pack has different addresses inside it (Install doesn't excuse differences within a pack) |
| 10 | The same store has store packs (Install not Y) with different addresses |

**Warnings — shown on the preview, generation allowed**

| # | Situation |
|---|---|
| 8 | A merged address cell runs past the edge of a pack, so neighbouring packs share it |
| 9 | A pack has no address at all (for example Sample rows) |
| 12 | The same store has several installer packs with different addresses (could be different installers) |
| 15 | Neighbouring packs have the same Packing Spec and store but aren't merged (possibly a missed merge) |
| — | A row has no thumbnail image |

**Always fine:** (1) the address merged across some rows of a pack and typed identically in the rest; (11) a store's installer pack having a different address from its store packs. Addresses are compared with `_address_key`, so capitals, spacing, punctuation and Street/St differences never trigger an error.

**How `_check_pack_consistency` decides.** For each pack it collects every row's store key, address key and install flag, then counts distinct non-blank values (more than one → "differs") and checks for a mix of blank and filled (→ "blank"). A merged address range that starts above or ends below the pack's own rows → "shared with next pack". Across packs of the same store it compares each pack's first row. Neighbouring packs with the same spec and store → "missed merge" warning.

**Showing where two addresses differ (`_variants`).** Long addresses are cut to 40 characters, which would hide the difference if it's at the end. So for each variant the code finds the **common word prefix** with the first variant and shows only what follows, keeping one word of context:

```
62-64: Shop 2 Woolworths Curlewis Town Centre, 90 Centennial Blvd, Curlewis, VIC, 3064
65:    Shop 2 Woolworths Curlewis Town Centre, 90 Centennial Blvd, Curlewis, VIC, 3222
→ "62-64: …VIC, 3064 · 65: …VIC, 3222"
```

### 9.6 Finding each row's image

Excel stores pictures two ways, and both are mapped to a `(row, column)` cell.

**In-cell pictures ("Place in Cell")** live in Excel's newer *rich value* parts of the `.xlsx` zip. `extract_rich_value_images()` follows a chain of references, so these positions are exact:

```
sheet XML: <c r="A88" vm="5">           cell A88 points to value-metadata entry 5
  → xl/metadata.xml                    → rich value block number
  → xl/richData/richValueRel.xml       → relationship id (e.g. rId7)
  → xl/richData/_rels/...rels          → xl/media/image12.png
```

**Floating pictures** are anchored by their corners, measured in EMU. `_floating_image_cell()` converts the anchor to an absolute position and takes the **centre**:

```
row heights:   points × 12,700                    (default 15 pt)
column widths: (width_in_chars × 7 + 5) pixels × 9,525
               (Excel's rule: 7 px per character of the default font, plus 5 px padding)

top    = sum of heights of rows above the anchor row + offset inside that row
bottom = same for the "to" corner (or top + picture height)
centre = (top + bottom) / 2
row    = walk the rows, adding heights, until the running total passes the centre
```

The column is found the same way. Using the centre means a picture whose top edge pokes into the row above still belongs to its own row.

**Matching a row to its pictures.** Images are taken from the row's **own thumbnail cell only** — never a neighbouring row or column. Every image in that cell is used, not just the first: each cell holds a list of images, in reading order (top to bottom, then left to right). A **merged thumbnail cell** counts as one cell: every row it covers gets all the images placed anywhere inside it, so a thumbnail merged over 5 rows gives all 5 boxes the same image(s), and one merged 3 + 2 gives rows 1–3 the first merge's images and rows 4–5 the second's. (Merged Job Number cells work the same way through `get_cell_info`: a job number merged over 5 rows makes 5 boxes with that job number, each with its own row's images.) A row with no image gets a "No image" warning and prints without one.

### 9.7 Drawing the labels: page and cell math

All drawing uses PyMuPDF in **points** (1 mm = 2.83465 pt). A4 landscape is 841.9 × 595.3 pt.

**One box size on every page.**

| Constant | Value |
|---|---|
| margin | 4 mm = 11.34 pt |
| gap between boxes | 2.5 mm = 7.09 pt |
| courier label region | 107 × 150 mm = 303.3 × 425.2 pt, top-left of page 1 |
| footer chip height | 11 pt × 1.2 + 8 = 21.2 pt |

Width comes from page 1, where 4 columns sit beside the courier region; height from the later pages, which lose a footer strip:

```
CELL_W = (page_w − 2·margin − courier_w − gap − 3·gap) / 4
       = (841.9 − 22.7 − 303.3 − 7.1 − 21.3) / 4 = 121.9 pt ≈ 43 mm
CELL_H = (page_h − 2·margin − footer − gap − 2·gap) / 3
       = (595.3 − 22.7 − 21.2 − 7.1 − 14.2) / 3 = 176.7 pt ≈ 62 mm
NEXT_COLS = floor((page_w − 2·margin + gap) / (CELL_W + gap)) = floor(826.3 / 129.0) = 6
```

(The `+ gap` on top accounts for n boxes having only n − 1 gaps.) Page 1 holds 4 × 3 = 12 boxes, later pages 6 × 3 = 18:

```
pages = 1                              if items ≤ 12
pages = 1 + ceil((items − 12) / 18)    otherwise        e.g. 32 items → 1 + ceil(20/18) = 3
```

**What's on each page.**

- **First page of a label:** the courier label region (dashed placeholder), and under it the **Packing Spec** in a black rounded panel with large white text (starts at 22 pt and shrinks 1 pt at a time until it fits on two lines, minimum 11 pt), the store name and address, and `LABEL X OF Y` / `PAGE X OF Y` chips at the bottom-left. The address is the one its consignment is sent to, exactly as finalised in the consignment preview (address-book fills and ✏️ edits included): Generate passes `{pack key: {address, receiver, contact, sent}}` to `generate_packing_labels(..., addresses)` for **every** pack.
  - **Installer packs** (Install = Y) get a red line under the store name: **Installer: <receiver> · Attn <contact>** (e.g. `Installer: Wilson Storage · Attn Steven Priestley`). Any other pack whose receiver isn't the store, or that has an Attn, gets **Deliver to: …**.
  - **A pack left out of the courier CSV** (no complete address) still gets its label. The courier label region then reads **NOT IN COURIER CSV — Address incomplete: no courier label for this box**, and the address shown is what the file gives.
- **Later pages:** no Packing Spec; the store name and the two chips sit in the bottom-left, and the boxes start at the top margin.
- **Label X of Y** (`assign_label_numbers`): a `Counter` of `_norm(store_name)` gives each store's total; a second counter increments as packs are visited in file order. Using `_norm` means `Spectacle Hub Curlewis` and `Spectacle Hub Curlewis ` are one store. When several tabs are generated together, a store is counted across all of them. The PDF and the CSV both call this one function, so their numbers always agree. **Page X of Y** counts pages within one label.
- **Rounded corners:** PyMuPDF takes radius as a fraction of the shorter side, so `radius = r ÷ min(width, height)`, capped at 0.5.

**Inside an item box.** `_layout_cell()` walks the chosen block order top to bottom. Each block's style belongs to its header, wherever it's placed:

| Block | Style | Height (s = text scale, t = thumbnail scale) |
|---|---|---|
| Thumbnail | 35% of the box height. One image: centred, 70% of the box width. Several: side by side across the box, up to 3 per row, then more rows. Every image keeps its own proportions (`_fit_rect`). | 0.35 × box height × t |
| Description | regular text, wrapped | lines × 9s × 1.2 |
| Dimensions | bold text, wrapped | lines × 9s × 1.2 |
| Material, Notes | regular text, value only (no header name) | lines × 9s × 1.2 |
| Install | bold red `INSTALLER` on installer packs | lines × 9s × 1.2 |
| Job Number | bold white text on a black rounded bar | lines × 14s × 1.2 + 6s |
| Quantity | large bold number in a heavy box, no unit | lines × 20s × 1.2 + 6s |
| Barcode | Code 128 of the item's barcode text, as wide as the box allows, with that text printed small underneath | 22s + 2 + 6.5s × 1.2 |
| between blocks | | 3s |

The `× 1.2` is line spacing (120% of the font size). Blocks with no value are skipped, which is how a missing column frees its space.

**Wrapping** (`_wrap`) is greedy: add words to the line while its measured width (`fitz.get_text_length`, real font metrics) fits; otherwise start a new line. A single word wider than the box is cut character by character. Text wraps onto extra lines rather than being cut off.

**Making everything fit** (`_draw_cell`). The cell is measured first and drawn only once a size fits. It tries `s` from 1.00 down in steps of 0.01:

```
text_scale  = max(s, 0.65)                 # text shrinks first, never below 65%
thumb_scale = 1        if s ≥ 0.65
              s / 0.65 otherwise           # only then does the image shrink
```

Every block's height grows with `s`, so the total height only gets smaller as `s` falls — the first `s` that fits is the largest that fits. Nothing is ever drawn outside the box. PyMuPDF silently draws *nothing* when text doesn't fit a text box, which is why everything is measured first (the same lesson as `_fit_textbox()` in the Label Shuffler).

**Barcodes.** Every item box can carry a Code 128 barcode (the **Barcode** block in the layout editor, placed after Job Number by default). `assign_barcodes()` follows the Distribution Mapper's kinds, with rows in place of columns:

- A job number on **more than one row** gets one kind per row, **K1, K2, K3…** in reading order — first tab to last, then top to bottom — so K1 is its first row in the first tab. The barcode encodes `J477161-26 K3`.
- A job number on **only one row** has no kind; the barcode encodes just `J477161-54`.
- Kinds are counted across every selected tab together, so the same job number on two tabs continues K1, K2 → K3.

The Job Number bar always prints the cell's exact text (line breaks kept, long lines wrapped). A job number with a line break or other character Code 128 can't hold gets **no barcode** and a preview warning ("No barcode: Job Number has a line break…"); the rest of the box prints as normal.

**Fitting a barcode in a 43 mm box.** Code 128 needs 10 blank modules either side, and the bars are made as wide as the box allows (`module = inner width ÷ (symbol modules + 20)`, capped at the Label Shuffler's 0.96 pt). To keep them as wide as possible:

- the **K1** format (instead of "Kind 1") shortens the text by 3 characters;
- `code128_modules(text, compact=True)` switches to **code set C** (two digits per symbol) for runs of 4+ digits at either end or 6+ in the middle, e.g. the `477161` in `J477161-26`. It decodes to exactly the same text.

`J477161-54 K1` comes out at **0.61 pt (0.21 mm)** per bar. Every encodable barcode on the 44-page Lux test 18 labels (105 of 105) decoded correctly at both 300 and 600 dpi with a standard decoder, with no wrong reads.

`generate_packing_labels()` returns which PDF pages each pack used, plus the page size and courier region in mm, for the label map ([10.9](#109-the-label-map-json)).

---

## 10. `courier_export.py` — Addresses, Consignments & the Courier CSV

**Purpose:** turn each pack group's address — however messy — into the courier's standard fields, group packs going to the same address into consignments, and write the courier CSV and the label map.

### 10.1 From rubbish addresses to a standard format

`resolve_destination()` turns whatever is in the address columns into the courier's fixed fields:

`receiver · contact · line1 · line2 · suburb · state · postcode · country · authority_to_leave`

There's no machine learning here. The method is: **peel off the parts you can recognise with certainty, from the outside in, and whatever is left is the street.** Each step removes something, so later steps work on a smaller, cleaner string.

**The patterns it has to handle** (from the real Lux file):

| Kind | Example |
|---|---|
| Store, split across columns | `49 Church Street` · `Brighton` · `VIC` · `3186` |
| Installer, all in one cell | `Steven Priestley - Wilson Storage, 68 Ricketts Road, Mount Waverley, Vic, 3149 - ATL` |
| Attention in brackets | `Digi Master Signs - 109A Almond Avenue, Mildura, VIC, 3500 (Attn Allen ATL)` |
| Misspelt "Attn" | `Sign Online - 117 Firebrace St, HORSHAM, VIC, 3400 - Atnn Adam ATL` |
| Two-part street | `Shop T11-14 The Strand Melbourne, 250 Elizabeth St ` |
| Numeric postcode cell | `3149.0` (Excel stored it as a number) |
| Everything in one column, any separator | `Andrew, 27 sirius road, lan cove NSW, 2067` · `Andrew \| 27 Sirius Rd \| Lane Cove NSW 2067` · `Andrew 27 Sirius Road Lane Cove New South Wales 2067` |

The installer shape is the general one:

```
[Name] - [Company,] Street, Suburb, State, Postcode [- Attn X] [ATL]
```

**The steps, with a worked example.** Input: `Sign Online - 117 Firebrace St, HORSHAM, VIC, 3400 - Atnn Adam ATL`

**Step 0 — gather the raw text.** Address Line 1 and Line 2 are joined; Suburb, State and Postcode columns are kept aside as a fallback.

**Step 1 — collapse spaces.** Double spaces and tabs become one space. **Line breaks are kept**: inside a cell (Alt+Enter) they separate parts of the address like commas do (step 6).

**Step 2 — Authority To Leave.** `\bATL\b` (case-insensitive) is searched for anywhere. `\b` is a *word boundary*, so it matches `ATL` as a whole word but not inside `ATLANTIC`. If found: `authority_to_leave = True`, and the word is removed.

```
before: Sign Online - 117 Firebrace St, HORSHAM, VIC, 3400 - Atnn Adam ATL
after:  Sign Online - 117 Firebrace St, HORSHAM, VIC, 3400 - Atnn Adam
```

**Step 3 — Attention name, at the end only.**

```
\(?\b(?:attn|atnn|attention|att)\b[:.]?\s*([^()]*?)\s*\)?\s*$
```

Read it as: *optional "(" · one of the attention words · optional ":" or "." · the name (anything except brackets) · optional ")" · end of text*. The `$` anchors it to the **end**, so a street like `Attwood St` in the middle is never mistaken for "Att". `atnn` is there because it really appears in the data. The name becomes `contact` = `Adam`, and the matched part is cut off.

**Step 4 — trim trailing junk.** `[\s\-–,()]+$` removes leftover spaces, dashes, commas and brackets from the end.

```
after: Sign Online - 117 Firebrace St, HORSHAM, VIC, 3400
```

**Step 5 — the name in front ("head").**

```
^([^\d]+?)\s+[-–]\s+(.*)$
```

*From the start: some text **with no digits**, then a dash with spaces around it, then the rest.* Two details make this safe:

- `[^\d]` — the head may not contain digits. Street addresses almost always do, so `Shop T11-14 The Strand` is never split (it has digits, and its dash has no spaces).
- `\s+[-–]\s+` — the dash must have spaces on both sides. Hyphens inside names and unit numbers (`24-26`) don't count.

```
head = "Sign Online"      text = "117 Firebrace St, HORSHAM, VIC, 3400"
```

**Step 6 — split into parts** (`_segments()`): `["117 Firebrace St", "HORSHAM", "VIC", "3400"]`. Commas, semicolons, `|`, line breaks, tabs, ` - ` and ` / ` all separate. A dash between two numbers (`Shop 1 - 3`, `4-10`) and a slash without spaces (`3/27`) don't. Phone numbers (`0412 345 678`, `+61 …`) are dropped first, and a part starting `Attn` becomes the attention name wherever it is.

**Step 7 — read the tail from right to left.** Australian addresses end in a predictable order — *suburb, state, postcode* — so the code pops tokens off the end while they match:

| Last token is… | Becomes |
|---|---|
| exactly 4 digits | postcode |
| then an Australian state (`VIC NSW QLD SA WA TAS NT ACT`) or a New Zealand region code — the official ISO 3166-2:NZ codes (`AUK BOP CAN CIT GIS HKB MBH MWT NSN NTL OTA STL TAS TKI WGN WKO WTC`) or Toll's own (`AUC CHR WEL MOU`) | state |
| then anything, if more than one token remains | suburb |

`parse_address_text()` reads the tail. The postcode can be a part of its own or end the last part (`Lane Cove NSW 2067`). The state can be a code or a full name (`New South Wales` → `NSW`, `Victoria` → `VIC`…). The suburb is what's left of that part. When the street and suburb share one part with no separator (`27 Sirius Road Lane Cove`), the suburb starts after the **last street word** (`Road`, `St`, `Hwy`… `STREET_WORDS`) that follows a name word. So `Lane Cove` and `St Kilda Road St Kilda` come out right: `Lane` after `Road` and `St` after a number aren't street ends. What the address cell spells out wins over the Suburb / State / Postcode columns, as before. A second pattern handles the case without commas between them: `Mount Waverley VIC 3149` as one token → `(suburb)? (STATE) (dddd)`. If the tail doesn't look like this (a normal store address), the Suburb/State/Postcode **columns** from Step 0 are used instead.

```
postcode = 3400   state = VIC   suburb = HORSHAM   tokens left = ["117 Firebrace St"]
```

**Step 8 — names at the front.** When the file has **no store / receiver column**, every leading part that names someone (`_is_name()`: no digits, no address word such as `Shop`, `Westfield`, `Centre`, `PO Box`, `Level`, not ending in a street word) is a name: `Andrew, 27 Sirius Rd…` → `Andrew`; `Andrew Smith, Acme Signs Pty Ltd, 27 Sirius Rd…` → person + company. A name run straight into the street (`Andrew 27 Sirius Road`) is the words before the street number. With a store column only the old rule applies: if there was a head, more than one token remains, and the first token has **no digits**, that token is a company:

```
Steven Priestley - Wilson Storage, 68 Ricketts Road, ...
head = Steven Priestley    tokens = ["Wilson Storage", "68 Ricketts Road"]
                           → company = "Wilson Storage", street = "68 Ricketts Road"
```

**Step 9 — is the head a person or a company?** `_is_company()` answers with three cheap tests; any one makes it a company: (1) more than 3 words (people's names are usually 2–3), (2) a word from `COMPANY_WORDS` (`sign`, `signage`, `storage`, `pty`, `ltd`, `print`, `display`, `&`, …), (3) any digit. `Sign Online` contains `sign` → company. `Steven Priestley` passes none → person.

**Step 10 — the receiver rule.**

| Name found | Receiver Name | Receiver Contact Name |
|---|---|---|
| Company only | company | attention name, if any |
| Person only | person | attention name, if any |
| Person **and** company | company | person (plus attention name) |
| Neither (store's own address) | store name | attention name, if any |

```
receiver = "Sign Online"   contact = "Adam"
```

**Step 11 — split long streets into two lines.** Couriers limit line length, so street text over `LINE_LIMIT` (30) characters is cut at its **last** separator (`, `, `; ` or `. `), because the building/shop part comes first and the street last:

```
"Shop T11-14 The Strand Melbourne, 250 Elizabeth St"   (50 chars)
→ line1 = "Shop T11-14 The Strand Melbourne"     line2 = "250 Elizabeth St"
```

**Step 12 — tidy each field.** `_tidy()` trims spaces and stray `, ; . -` from both ends; state and **suburb are upper-cased** (`Lane Cove` → `LANE COVE`). Suburbs are kept in capitals everywhere: preview edits, the courier CSV, the label, the sender, bulk-update files and the address book (`_clean()`; entries saved before that are upper-cased once when the book opens). Matching ignores capitals, so this changes no keys or learned spellings. The postcode goes through `normalize_postcode()`: Excel's `3149.0` becomes `3149`, and a 3-digit Australian or New Zealand postcode gets its lost leading zero back (`803` → `0803`), because the courier portal rejects 3-digit postcodes. The same padding is applied in the address book, to edits, and when importing verified addresses; entries saved before it are fixed once when the book opens.

**Country** comes from the Country column (or a trailing `New Zealand` / `Australia` in a one-cell address) when there is one; otherwise it's `NZ` for a New Zealand region code and `AU` for everything else. `TAS` is both Tasmania and the Tasman region, so on its own it means Australia — give such addresses a Country of New Zealand. Changing the state in the ✏️ edit form re-applies this rule.

Final result:

```python
{'receiver': 'Sign Online', 'contact': 'Adam', 'line1': '117 Firebrace St', 'line2': '',
 'suburb': 'HORSHAM', 'state': 'VIC', 'postcode': '3400', 'country': 'AU', 'authority_to_leave': True}
```

**A second example (person + company):** `Steven Priestley - Wilson Storage, 68 Ricketts Road, Mount Waverley, Vic, 3149 - ATL`

| Step | Result |
|---|---|
| 2 | ATL → `True` |
| 3 | no attention name |
| 5 | head `Steven Priestley` |
| 7 | postcode `3149`, state `Vic` → `VIC`, suburb `Mount Waverley` |
| 8 | first token `Wilson Storage` has no digits → company |
| 9 | `Steven Priestley` → person |
| 10 | receiver `Wilson Storage`, contact `Steven Priestley` |

**Files with only one combined address column** (`fill_store_names()`, run right after the file is read, `_read_tab()` in app.py):

- A pack with **no store name** takes the receiver named in its address cell as its store name, so labels are headed with it and numbered per receiver (`LABEL 1 OF 2` for Andrew's two boxes). `store_from_address` remembers this, so the name isn't also left in the street.
- An address with **no name** in it (`27 Sirius Road, Lane Cove NSW 2067`) is named after its street for now. In the preview, the address book supplies the receiver (and Attn) saved at that exact address (`_apply_address_book()` → `at_address()`), and the label and item reference then carry that name. If the book doesn't know it, the preview warns "No receiver name … add it with ✏️".
- A **store / receiver column holding whole addresses** (a `Ship To` column, with no address column) is read as the address.
- With no store column, the "same spec split into separate packs" check compares the addresses of neighbouring packs instead.

**Why rules and not guessing.** Every rule is either **anchored** (to the start, the end, or a word boundary) or **gated by a fact that is almost always true** (streets have digits; names don't; postcodes are 4 digits). That keeps false matches rare, and when a rule doesn't fire, the value falls back to the column it came from rather than being invented. Anything still odd shows up in the consignment table, where the ✏️ button lets a person correct it ([10.6](#106-the-consignment-table-service-codes-and-edits)).

### 10.2 Grouping packs into consignments (`build_consignments`)

```
for each pack:
    dest      = resolve_destination(pack)                 # 10.1
    source_id = _address_key(one_line(dest))              # id of the address as read from Excel
    final     = dest with the user's edits applied        # 10.6
    key       = _address_key(one_line(final))             # what decides the consignment
    consignments[key].cartons.append(pack)                # first pack at this key creates it
```

Because the key is the address alone, packs for **different stores going to the same installer** fall into one consignment automatically, and a retailer with a shop pack and an installer pack gets two. Consignments are numbered in order of first appearance. `one_line()` joins `line1, line2, suburb, state, postcode` with `", "`.

**Packs with a receiver name but no address** (a file with only a Store/Receiver column, or rows with no postcode) aren't dropped. `pack_destination()` gives each one a consignment of its own receiver, with the id `name:<receiver>|<whatever address there is>`, so the address book and the ✏️ form can fill it in:

1. **Learned before:** Generate saved that id as a spelling (alias), so the next file fills the address in like any other learned spelling.
2. **First time:** `_apply_address_book()` looks the receiver up (`find_by_receiver()`). When every saved entry for that receiver is at one address, it's filled in (blue **Address book** tag).
3. **Several saved addresses, or none:** the row is listed first in red, marked **No address**, with a note above the table. Its ✏️ form lists the receiver's saved addresses to click ("Saved for this receiver"), or says there are none and the address should be typed in.

A consignment still without an address when you Generate is left out of the courier CSV, the label map and the address book (`sendable()`). Its label shows just the store name. Packs with neither a name nor an address are left out with a warning, as before.

### 10.3 Carton facts from the Packing Spec

Size, weight and item type come from [packing_specs.py](../packing_specs.py) and the **📐 Packing Specs** page (`/packing-specs`, [packing_specs.html](../templates/packing_specs.html)). The values are saved in `data/packing_specs.json`, written atomically (temp file + replace).

**Codes with the size in them: OB, CS, Pallet, FP** (`parse_formula()`). The numbers are millimetres, ÷ 10 for cm:

```
OB 1370 170 170  → 137 × 17 × 17 cm
CS 170 170 1200  → 17 × 17 × 120 cm
Pallet 1200 800 500 → 120 × 80 × 50 cm
FP 120 170       → 12 × 17 × 5 cm   (FP usually gives two sizes; the third comes from the page, 5 cm by default)
```

The numbers can be separated (`OB 170 x 170 x 1200`) or run together (`OB1701701200`). When they run together, `split_dimensions()` tries every way to cut the digits into sizes of 2–4 digits with no leading zero:

1. Sizes are whole centimetres, so every millimetre value ends in 0: `1701701200` can only be `170 170 1200` (not `1701 701 200`).
2. If that doesn't settle it, the old rule applies: the later sizes are 3 digits and the length takes the rest (`1375170170` → `1375 170 170` = 137.5 × 17 × 17).
3. An FP code is read as two sizes unless only a three-way split gives whole centimetres (`FP12017050` → 12 × 17 × 5 from the code).

Each code type's **weight** and **item type** are set on the page (OB, CS and FP 2 kg Carton; Pallet 200 kg Pallet), plus FP's **3rd size**.

**Other specs** (A4 box, Half A4 box, A3 box, Half A3 box, SRA3, P7 / P5 / P1 jiffy bag…) are matched by name, ignoring capitals and punctuation, and use the length, width, height, weight and item type saved on the page. A spec found in a file that the page doesn't list is added to it with no values, marked **Needs size** (`note_unknown()`). Until it's filled in, its dimensions stay blank in the CSV and the preview warns with a link to the page. A spec with no weight uses 2 kg.

**Weights in the preview.** Every carton in the consignment table has a **Weight** box, starting at its spec's weight from the page (a pallet at 200 kg). A changed value is outlined, kept in the hidden `carton_weights` field (`{pack key: kg}`) and sent with Update Previews and Generate (`carton_weight()`). An empty box goes back to the spec's weight; a value that isn't a number between 0 and 5,000 is outlined red and ignored. Under each carton the table also shows the size and item type that will go into the CSV.

**Check a code** at the bottom of the page shows how any Packing Spec will be read.

| Field | Rule |
|---|---|
| Total Cubic (m³) | L × W × H ÷ 1,000,000, rounded to 3 places: `137 × 17 × 17 = 39,593 cm³ → 0.04` |
| Total Weight | the weight typed in the preview, else the spec's weight from the page |
| Item Type | the spec's item type from the page |
| No Items | always 1 (one row per pack) |

### 10.4 The OpenFreight CSV row

Written only when **OpenFreight** is chosen above the project name (see 10.12 for the choice; Open360's layout is in 10.12). One row per pack, 44 columns matching the courier's import format (UTF-8 with a BOM, like the sample the courier provided):

| Column | Value |
|---|---|
| Dispatch Date | today, as `4-Oct` |
| Reference (column 2) | the consignment reference — **only on the first row of each consignment**; later rows leave it blank and repeat the exact same address text so the courier joins them |
| Receiver Name … Country, Contact | from 10.1 (after any edits) |
| Authority To Leave Flag | `Y` when the address contained `ATL` |
| Who Pays, Charge Account | from the page (Who Pays defaults to `S`) |
| Service Code | the consignment's own pick, or the main Service Code |
| No Items | the pack's boxes: 2 for `2 x OB170170170`, else 1 |
| Weight, Cubic, Item Type, L/W/H | from 10.3; Weight and Cubic count every box |
| Sender Name … Sender Email | blank: the portal uses the account's sender |
| Reference (last column) | the label's **Item Reference**, the same as Open360's: `01T1QWXK Store` (serial + the job's code + store, 10.12), so courier labels booked from either file stitch by code |

The **consignment reference** is pre-filled from the distro's Job Numbers, strictly **J + 6 digits** at the start (`J477161-54` → `J477161`; the most common one when there are several), and is editable. The files are named `<reference> - <project> - OpenFreight.csv` and `<reference> - <project>.labelmap.json`.

### 10.5 Address warnings

- a receiver appearing at more than one address when installer packs are involved,
- the same street (`_address_key` of line1 + line2) with a different suburb/postcode,
- incomplete addresses, unknown carton sizes, packs left out for having no postcode,
- job numbers using more than one series.

### 10.6 The consignment table: service codes and edits

After a clean preview, "3. Courier Consignments" lists each consignment's receiver, address, service code and cartons, and asks for **Consignment Reference**, **Who Pays** and the main **Service Code** (required) and **Charge Account** (optional).

- **Service codes per consignment.** Each row has its own dropdown; "Same as main (CODE)" uses the main Service Code. Filter the table by **State** or by text (receiver, suburb, postcode, store), then **Apply to shown** copies the main code into every visible row.
- **Editing an address.** The ✏️ next to a consignment number opens its receiver, contact, address lines, suburb, state, postcode and Authority to Leave for editing, with the original Excel text underneath. **Save** marks the row "Edited"; **Reset to Excel** puts the original values back. Enter saves, Escape closes.
- **No sender section.** The portal uses the account's sender and the Open360 CSV leaves the sender columns empty, so the preview has no sender box (it, `senders.py` and `/api/senders` were removed).
- **Update Previews reads what's in Excel.** The upload may be open in Excel (the Open button). Changes not saved yet are saved first (`save_if_open_elsewhere()` in core_math.py, through Excel itself), and the preview says which saved version it read ("Read … as saved at 23:00:23"). If Excel is in the middle of editing a cell it can't save: the note says to press Enter or Esc in Excel and update again. Generate does the same before closing the file in Excel, so edits shown in the preview are never thrown away. Editing your own copy elsewhere doesn't reach the upload: the note says so.
- **The browser's own form check is off** (`novalidate` on the form). It used to stop Update Previews too, because Project Name (needed only for Generate) is empty until the end: the browser scrolled down to that box with a bubble that was easy to miss, and nothing was sent, so the file was never re-read. Now Update Previews always goes through (header rows must be numbers), and Generate is checked by `script.js` section 14: every required box filled in, else a red message above the button names them ("Fill in Name this Project / Job before generating."), the box is outlined and the cursor put in it, and nothing is sent. The server checks the same boxes, Project Name included.
- **Enter in the form.** Enter in a box would submit with the form's first button, Update Previews (the page reloads and scrolls, nothing is generated). `script.js` section 14 makes Enter in **Project Name** press Generate, Enter in a **Header Row** press Update Previews, and Enter anywhere else do nothing.
- Edits change only the courier CSV and label map, never the Excel file. They're kept while you refresh the preview, and lost if you scan the file again.

**How edits travel.** The browser keeps them in a hidden JSON field, `consignment_edits`, sent with Preview and Generate:

```json
{"68 ricketts rd mount waverley vic 3149": {"service_code": "IPECY"},
 "shop 1 604 glenhuntly rd elsternwick vic 3185": {"receiver": "...", "line1": "...", "postcode": "3088"}}
```

- The keys are **source ids** (the address as read from Excel), not consignment numbers. Numbers change when consignments merge; the source address doesn't, so edits survive any number of preview refreshes.
- Edits are applied **before** grouping (`_apply_edit`). If an edited address produces the same key as another consignment, the two merge after "Update Previews" — the same thing the courier would do.

### 10.7 Service code order

`SERVICE_CODES` is the list of codes and descriptions. Each Generate adds 1 per consignment to `data/service_code_usage.json` (`record_service_usage`). The dropdowns sort by the tuple

```
(−usage_count, position_in_SERVICE_CODES)
```

Sorting by a tuple compares the first value, and only on a tie the second — so the most used come first, and unused codes keep the catalogue order. The most used code is pre-selected as the main Service Code; with no history yet the dropdown starts on "Choose a service…". Generate refuses any code not in the list.

### 10.8 Order of work in Generate

All courier checks (required fields, known service codes) run **before** the project folder is created or the Excel file moved. If anything still fails afterwards, the Excel file is moved back and the half-made folder removed, so Generate can simply be retried. (Same "validate first, write last" idea as the Sub-Group Engine in [Section 7.2](#72-the-three-stages).)

### 10.9 The label map JSON

`<reference> - <project>.labelmap.json`, saved in the project folder for **Stitch Labels** (10.13) and the address book's bulk update (10.11), never shown in the app:

```json
{
  "consignment_reference": "J477161",
  "courier_label_region_mm": {"x": 4.0, "y": 4.0, "width": 107.0, "height": 150.0},
  "consignments": [{"number": 2, "destination": {...}, "service_code": "STEROAD", "cartons": ["Label 1 - Provision Clayton", "..."]}],
  "not_in_courier_csv": [{"item_reference": "Label 1 - Eyedentity (Elsternwick)", "pdf_pages": [4], "reason": "no complete delivery address", "...": "..."}],
  "stores": {
    "Provision Clayton": [{
      "item_reference": "Label 1 - Provision Clayton", "label": 1, "of": 1,
      "open360_item_reference": "01T1QWXK Provision Clayton", "open360_code": "01T1QWXK",
      "pdf_pages": [2], "courier_label_page": 2, "boxes": 1, "box_pages": [2], "consignment": 2,
      "headed_to": {"receiver": "Wilson Storage", "address": "68 Ricketts Road, ...", "installer": true},
      "csv_row": 3, "excel_rows": [5, 6], "job_numbers": ["J477161-03", "J477161-04"],
      "barcodes": ["J477161-03", "J477161-04 K1"]
    }]
  }
}
```

Each pack's **`open360_code`** (serial + the job's code, `01T1QWXK`) is what Stitch Labels reads back off the courier label (it's in the Item Ref printed on it); it leads to the pack, its boxes' pages (`box_pages`, one per box of a `2 x` pack) and the courier region's position on them (x, y, width, height in mm from the top-left). The internal **item reference** (`Label 1 - Provision Clayton`) names the pack inside the app. `files.courier_csvs` lists the courier CSVs written (Open360 and / or OpenFreight).

### 10.10 The address book

[address_book.py](../address_book.py) keeps every clean delivery address in a local SQLite file, `data/address_book.db`. SQLite needs no server, comes with Python, and only loads the rows a query asks for, so memory stays flat however big the book gets. Everything goes through the `AddressBook` class, so moving to a cloud database later means replacing this one file.

**Tables**

| Table | Holds |
|---|---|
| `addresses` | one row per receiver + Attn + address (unique on address key + receiver key + contact key), with `use_count`, `last_used_at`, `created_at`, `updated_at`, `verified_at`. One address can have several receivers and Attn names, so each combination is its own entry |
| `addresses_fts` | an FTS5 full-text index over receiver, contact, address lines, suburb, state and postcode, kept in sync by triggers |
| `aliases` | raw Excel spelling (the consignment's source id) → the address it was saved as. Deleting an address deletes its aliases |

**How it learns**

1. **Generate** calls `record_used()` with every consignment's final address (after any ✏️ edits) and the source ids of all its packs, in one transaction. A known receiver + Attn + address gets its details refreshed and `use_count` + 1. A receiver or Attn not yet saved at that address is added as its **own entry** next to the others (nothing is overwritten); a blank Attn just counts as a use of the receiver's existing entry. A new entry at an address the courier portal already verified is marked verified too. Every raw spelling is pointed at the entry.
2. **Preview** calls `_apply_address_book()`: for each Excel address not yet asked about, `lookup_aliases()` finds what the book learned for that spelling. If it differs from what was read from Excel, it's applied as an edit marked `'source': 'book'`, and the row shows a blue **Address book** tag.
3. Each spelling is offered **once per scan** (listed in `__offered__`), so **Reset to Excel** sticks. Saving your own edit replaces the book's.

So correcting an address once — in the ✏️ form before Generate, or on the Address Book page — fixes it for every future file that spells it the same way.

**Checking every consignment against the book** (`_check_against_book()` in app.py, on each preview). Each consignment gets a status:

| Status | When | Shown |
|---|---|---|
| book | filled in from a learned spelling (step 2) | blue **Address book** tag |
| matched | the book holds this exact address, under **any** receiver or Attn (`at_address()`) | nothing; if the receiver or Attn is new to that address, a note says it will be added on Generate |
| checked | not in the book, but you've edited it or pressed **Save** in its ✏️ form (`"checked": true` in the edits, so it survives Update Previews) | **Edited** tag if changed |
| unverified | not in the book | listed **first**, amber, **Check address** tag, with a count above the table |

**Every** row's ✏️ form lists the **closest saved addresses**, most likely first (`address_match.suggestions()`, top 5). For a row filled from the book, that entry comes first, so after **Reset to Excel** the saved address is still one click away. The list **re-ranks as you type in any field**: `attachClosest()` in script.js sends what's typed so far to `/api/addresses/suggest` 200 ms after the last keystroke, cancelling a stale request. Half-typed text counts: `Eyed` finds `Eyedentity (Elsternwick)` (`_typed_name_score()`), `604 Glenh` finds the street, and a postcode alone lists what's saved there. Before a name is typed, the ranking uses the address alone. Candidates are the same postcode (`near()`), the same receiver name, and full-text matches for the receiver in case the postcode itself is wrong. The ranking leans on the address (60 %) over the name (40 %), because one address often has several receivers; another postcode counts for less. Each card says why ("Same address, another receiver", "Same receiver, different address", "Similar address"…), and clicking it fills the form. If nothing is close, the form shows a disclaimer instead: check the address against the job before generating, since it will be saved to the book as it is. Saving the form clears the flag. Nothing blocks Generate.

The upgrade to receiver + Attn entries ran once on books made before it (`_add_contact_key()`): the table is rebuilt with the same ids, so learned spellings and the search index stay valid, and a copy is kept first as `data/address_book.before-attn-upgrade.db`.

**Searching**

- **Search-as-you-type:** finished words must match whole words; the last word, still being typed, matches as a prefix (`49 church st` → `"49" AND "church" AND "st"*`), so `49` doesn't also find `499`. A trailing space makes the last word whole too.
- **Ranking:** FTS5's `bm25` relevance with column weights — receiver 4, address line 1 3, suburb and postcode 2, the rest lower — then most used first.
- **One character** matches a large share of the book, and ranking all of it cost ~90 ms at 250k, so a one-character search returns the first matches unranked (~2 ms); the next keystroke gets full ranking.
- Results come back one page at a time (`LIMIT 51` to know whether there's more), never with a full count, so typing stays fast at any size.

**Concurrency.** Flask serves requests on several threads, so each thread gets its own connection (`threading.local`). The database runs in WAL mode, which lets searches read while a Generate is writing.

**Measured** (synthetic addresses, this PC):

| | 50,000 | 250,000 |
|---|---|---|
| File size | 35 MB | 156 MB |
| Search (in Python), p50 | 0.2–1 ms | 0.5–11 ms |
| Search through HTTP (request → JSON), p50 | — | 5–18 ms |
| One Generate: look up 350 spellings / save 40 addresses | 5 ms / 6 ms | 8 ms / 6 ms |
| Extra memory for the book | ~20 MB | ~40 MB |

On the page, the first 50 rows appear ~20 ms after load, "Load more" and saving take ~15 ms, and the page never holds more than the pages loaded (≈1,400 elements, < 1 MB of JavaScript memory at 100 rows). Eight searches at once while a Generate saved ran with no errors.

**The pages**

- **Address Book page** ([address_book.html](../templates/address_book.html), `script.js` section 12): search box (waits 120 ms after the last keystroke and cancels stale requests), 50 rows at a time with **Load more**, ✏️ edit and 🗑️ delete per row, **+ Add address**. Edits happen in a row under the entry (Enter saves, Esc cancels); errors appear in that row. The column headings stay pinned just under the nav bar while you scroll: script.js measures the nav's height into `--nav-height`, and the table's wrapper deliberately has no `overflow` (a scrolling or clipping box would pin the headings to the box instead of the page). Below 1100 px wide, Country, Used and Last used are hidden so the table still fits.
- **Consignment ✏️ form:** the closest saved addresses (above), re-ranked live as any field is typed in; a click fills the whole form. A row that isn't in the book also lists its closest saved addresses (above).

### 10.11 Updating the book from the courier portal

Addresses can only be verified once the consignment CSV is uploaded to the courier portal (Toll), which then offers a CSV of the verified addresses **with the same headers as the CSV this app generates**. The **⬆ Update in bulk** button on the Address Book page opens `/address-book/update` ([address_verify.html](../templates/address_verify.html)). It isn't tied to one project: each row is checked against every project's label map and the address book itself.

**1. Read the file** ([address_import.py](../address_import.py)). CSV **or Excel** (`.xlsx`, or `.xls` converted through Excel), in any layout: other sources won't match the Toll headers, so only the columns the address book needs are read and everything else is ignored.

- `read_sheets()` turns the upload into rows of text (CSV: handles Excel's BOM, older Windows encodings and the delimiter; Excel: every sheet, numbers like `3186.0` read as `3186`).
- `map_columns()` looks for the header row in the first 15 rows (so title rows above it are fine) and scores every header against each field by **keywords** (`FIELD_KEYWORDS`): an exact name scores 3, containing a keyword 2, plus 1 if it says it's the receiver's (receiver, delivery, ship to, consignee). So `Receiver Name`, `Company`, `Customer Name` and `Ship To Name` all work for the receiver; `Street`/`Address` for line 1; `Town/Suburb`/`City` for suburb; `Post Code`/`Zip` for postcode.
- Headers about someone else (`Sender`, `From`, `Pickup`, `Return`) or about phone, email, references, dates, weights… never count, so `Sender Name` or `Contact Phone` are never mistaken for the receiver or contact.
- `pick_sheet()` uses the sheet whose headers fit best. At least a receiver or address column plus a postcode column are needed.
- Columns the file doesn't have are simply not compared, so the entry keeps its existing contact, ATL, country, etc.

**2. Match each row to the entry it corrects** (`build_review()`). The **receiver name** comes first: the same address can serve many jobs, and the receiver is what identifies an entry. Toll's Open360 export, for one, has no item reference at all (its "Shipment Ref" is just the J number).

| Situation | Result |
|---|---|
| one saved entry with this receiver (several Attn entries at one address count as the same place: `find()` prefers the one with the row's Attn) | matched (**receiver name**) |
| several entries with this receiver, one at exactly this address | matched (**receiver name + address**) |
| several, and the row's item reference / J number leads to one of them via a project label map | matched (**receiver name + job**) |
| several, and one of their addresses is obviously this one | likely match (**receiver name + similar address**) |
| several, and nothing tells them apart | not matched, left **unticked** ("several saved addresses for this receiver") — never guessed |
| no receiver column, but an item reference leads to one entry | matched (**item reference**) |
| no entry with exactly this receiver, but one with the same postcode obviously is it | likely match (**likely match: name NN%, address NN%**) |
| none of the above | **new**, ticked by default (untick to skip) |

The portal's text is saved exactly as it comes — including lines Toll has shortened (it cuts long address lines, e.g. `…Westfield Shoppingtown 4-10 Jamie`).

**Likely matches** ([address_match.py](../address_match.py)). The portal often spells the receiver a little differently (`YD DFO HORNBY` for the saved `YD Hornby DFO`, `Rebel Mt Gravatt` for `Rebel Mt Gravatt (Garden City)`). A likely match is ticked like any update, with the note *"Close match, not exact — check it's the same place before applying."* It is only proposed when the match is obvious:

- **Short-list** (`AddressBook.near()`): only entries with the same postcode (indexed). In a crowded postcode (more than 100 entries, e.g. a CBD), the name and street words of all its entries are read once per import (`cache`) and the 30 sharing the most words with the row are scored.
- **Name score** (`name_score()`, 0–1): shared words (ignoring *Pty Ltd, The, Store, Shop…*), one name containing the other (a single shared word like *Rebel* is weak evidence), and character similarity for typos (only worked out when the cheap upper bound could pass).
- **Address score** (`address_score()`, 0–1): the **numbers** must agree (shop/unit/street numbers like `S240`, `4-10`). If each side has a number the other lacks (`Shop 1040` vs `Shop 2210` at the same centre), the score is 0, because that's a different shop. The **street words** must overlap, with Street = St etc., and a word cut short by the portal still counts (`Jamie` ≈ `Jamieson`). The **suburb** must be the same.
- **Same address, several receivers / Attn names:** entries at the same address aren't rivals (a verified address reaches all of them). Among them the best name wins, then the Attn closest to the row's (`_attn_score()`: `Sam` ≈ `Samantha`).
- **Accepted** (`best_match()`) only if name ≥ 45 %, address ≥ 55 %, the average ≥ 70 %, **and** it beats the best candidate at another address by 10 points. Otherwise the row is new, never a guess. So another brand in the same centre, or the same brand at another shop number, is not matched.
- **Speed**: under 1.5 ms a row at 50,000 entries, about 3–4 ms in a postcode with 2,000 entries, so a 1,000-row file takes 1–4 s.

**3. Review.** Several CSV rows for one entry (one per carton) become one proposal. Every row you can act on is a **card** with the portal's address (changed parts highlighted) and a choice of what to do, the suggested one already selected. Tabs (**All · To update · Not in the book · Conflicting**) filter the cards. The buttons **Suggested choice**, **Merge into best match**, **Add all as new** and **Skip all** set every card shown at once. The bar by **Apply N changes** says what will be saved (`3 updates · 5 new · 2 merges · 1 skipped`), and each card says what will happen to it.

| Card | Choices (suggested first) |
|---|---|
| to update | **Update the saved address** (shows what's saved now, the parts that change highlighted, plus any older copies folded in and other receivers corrected) · Merge into another saved address (only for close, not exact, matches) · Add as a new address instead · Skip |
| not in the book | **Add as a new address** (Skip instead when the receiver has several saved addresses) · **Merge into this saved address**, for up to 4 of the closest saved addresses (`suggestions()`, each labelled e.g. "Same address, another receiver", with the parts that would change highlighted) · Skip |
| conflicting | Use row N's address, for each disagreeing row · **Skip** (fix the file) |
| already up to date | no choice: marked ✓ Verified on apply; listed in a separate table below the buttons |

Merging a row into a saved address works like an update of that entry: it takes the portal's address details, keeps its own receiver and Attn, and is marked verified. The form sends one `action_<row>` per card (`update`, `new`, `merge:<entry id>`, `use:<row>`, `skip`) to `apply_review()`.

**What a verified address changes.** Only the address details: lines, suburb, state, postcode, country (`PLACE_FIELDS`) and ATL. The saved **receiver name and Attn are never changed**. They identify the entry, and one address is often saved for several receivers or Attn names. So a portal row reading `TEST OPTICAL PTY LTD`, Attn `Samantha` updates the entry `Test Optical`, Attn `Sam` without renaming it.

**One copy per place.** For each matched entry, `build_review()` also looks for:

| | What | Shown on the review | On apply |
|---|---|---|---|
| `merges` | older copies of the **same place** for the **same receiver + Attn** (`same_identity()`, address score ≥ 0.85), e.g. St Kilda saved once with state `D` and once with `VIC` | "Also replaces an older saved copy…", with the old copy struck through and its differences highlighted. The row is ticked even if the address itself was already right | `merge_into()`: its learned spellings and use count move to the verified entry, and the copy is deleted |
| `siblings` | **other receivers / Attn names** saved at the old address (`at_address()`) | "The same address is also saved for …" | each gets the corrected address details through `apply_verified()`, keeping its own receiver and Attn |

So after a portal file confirms an address, the book holds one copy of it per receiver + Attn, all with the verified details.

**4. Apply** (`apply_review()`): every **already up to date** entry is marked verified (`mark_verified()`), no tick needed. Each ticked update goes through `AddressBook.apply_verified()` with the address details only: the entry takes the portal's address and gets `verified_at` set, shown as **✓ Verified** on the Address Book page. Then its `merges` are folded in and its `siblings` corrected, as above. Its learned Excel spellings stay linked, so the next file with the same spelling gets the verified address. If the verified address is one the book already holds under another entry, the two are **merged**: the existing entry keeps the address, takes over the use count and spellings, and the duplicate is removed. One bad row doesn't stop the others; problems are listed on the result screen. When nothing needs changing, the button reads **Mark as verified**.

---

**Which courier CSVs** (chosen just above the project name for every job: **Open360**, **OpenFreight** or both; nothing is ticked until the user chooses, and Generate refuses without a choice, in the page and on the server, `script.js` section 19). Generate writes `… - Open360.csv` and / or `… - OpenFreight.csv` (`write_courier_csv()`, the older 44-column layout with Who Pays and Charge Account). **Both carry the same Item References** (`01T1QWXK Store`; OpenFreight's last *Reference* column), worked out once (`open360_item_references()`), so courier labels booked from either portal stitch by code with the same safety. In the OpenFreight CSV the first *Reference* column (the Consignment Reference) is filled only on the first row of each address, blank on the rows repeating it (in file order, across tabs), and a pack of several boxes is *No Items* 2, with the weight and cubic of both boxes. Generate stops if neither is ticked ("Fill in Courier CSV …"). The label map lists the files written (`files.courier_csvs`), and the report's page 1 says which were made.

**A different service per consignment.** Each consignment (delivery address) can have its own service, or "Same as main". Every carton row of it carries that code in either CSV, the label map stores it, each is counted in the usage that orders the service lists, and stitching doesn't depend on it. The report's page 1 lists every service in use with its count (`Service: IPECX x2, BORDERP x1`). Merging one receiver's consignments (9.2's "different addresses" groups) gives the merged one **the service of the address kept**; when their services differed, the group heading says "different services: …" and Merge asks first. Still to confirm with the portal: whether Open360 splits a mix of services under one Shipment Reference as wanted, and whether other carriers' labels print the Item Ref (if not, they're matched by address on the matching page).

### 10.12 TIG Open360 bulk-upload CSV

TIG's Open360 portal has its own 46-column bulk-upload template (`OPEN360_HEADER` in courier_export.py, copied exactly, including the spaces in front of some names). When **Open360** is chosen above the project name, **Generate writes `<reference> - <project> - Open360.csv`** (and the OpenFreight CSV of 10.4 too when that's chosen). It holds the finalised consignments (`open360_items()` → `write_open360_csv()`). Its **Shipment Reference** is the preview's **Consignment Reference** (one field; there is no separate Shipment Reference field). The label map records the file (`files.open360_csv`), the reference (`open360_shipment_reference`) and each label's `open360_item_reference`. For a project generated before this, `project_open360_items(project_dir)` rebuilds the rows from its courier CSV as last saved plus its label map.

The format is locked to the file the portal accepted:

- **One row per packing label, in the distribution file's order.** The portal joins a receiver's rows into one consignment itself: a test upload gave Wilson Storage's 6 cartons as items 1–6 of one consignment.
- **Shipment Reference** = the reference, on every row.
- **Item Reference** = `01T1QWXK Provision Clayton` (`open360_item_references()`):
  - the label's **serial**: its count in its tab, zero-padded within the tab (`01…42`, `001` with 100+ packs), then `T` and the **tab** (`T1` is the first selected tab, `T2` the next one to the right, …). Generate works the serials out for every pack in the distribution file's order, packs left out of the CSV included, and prints the same serial on the packing label (`01T1. Provision Clayton`), so a courier label and its packing label can be checked against each other by eye. A pack left out of the CSV leaves a gap in the CSV's serials.
  - the **job's code**: one random code of capital letters for the whole Generate (`unique_codes(1)`), the same on every row
  - a space and the store
  **Serial + job code (`01T1QWXK`) is the label's unique code**: the label map stores it as `open360_code`, and Stitch Labels reads it back (`find_code()` / `_tab_code()`; the capitals after the serial are tried from the longest down, so a store name the portal runs straight on doesn't get in). The job's code is what keeps labels safe: courier labels booked from another job's CSV, or from an earlier Generate of the same job (whose serials may have shifted), carry another code, so they're never placed on this job's packing labels; they're flagged, with the other project named (`codes_elsewhere()` knows both the full codes and the job codes). Projects generated with a code per label (`01T1QWXK` with QWXK per label) or the older `01-QWXK-L-1-Store` layout are still read. All selected tabs go into the **one** Open360 CSV. Sorting by Item Reference in the portal puts `01T1, 01T2, 02T1 …` together (the serial comes first).
- **Internal Reference** is kept but empty.
- **Printing order in the portal:** its sort options sort *shipments*, and it prints a shipment's labels together, last carton first, placing the shipment by its last carton's Item Reference. So labels follow the file's order exactly only when a receiver's cartons are next to each other in the file. Stitch Labels doesn't depend on the order.
- **Item Quantity** and **Total Items** = the number of boxes on that row: 2 for a Packing Spec `2 X OB170170170` (`split_count()` in packing_specs.py; size and weight are those of one OB170170170), else 1. **Total Weight / Total Volume** = the whole consignment's, counting every box.
- **Item columns:** Item Weight (kg), L / W / H (cm), Item Type, Item Description = Packing Spec.
- **Service** = our service code. **Despatch Date** `d/m/yyyy` (as Excel saves it). **Authority To Leave** `Y` / `N`; **Is Receiver Residential** and **Is Dangerous Goods** `N`.
- **Sender columns kept but empty**, Is Sender Residential included: the portal fills in the account's sender. So are Internal Reference's neighbours Cost Centre, Special Instructions, receiver email / phone and the DG fields.
- **File format:** UTF-8 without a BOM, CRLF line endings.

### 10.13 Stitch Labels

**🧵 Stitch Labels** on a courier project in the dashboard opens `/stitch/<project>` ([stitch.html](../templates/stitch.html), [label_stitcher.py](../label_stitcher.py)). Upload the courier portal's label PDFs: any number, picked or dropped all at once or a few at a time (the upload box adds to what's already chosen, and **✕ Clear** empties it). Or tick ones already in the project. Uploads are kept in the project folder, and the matching page's **➕ Add courier label PDFs** stitches PDFs downloaded later in with the ones already used (`/stitch/<project>/add`, in the background). Every courier label lands on its packing label, in whatever order the portal printed them.

**Matching** (`find_code()`):
- The label's text is read. The value after **Item Ref:** or **Item Reference:** (any capitals, with or without the colon; a label may carry both) is taken, and the unique code in it (`01-QWXK-L-…` → `QWXK`) is looked up in the label map. The longer word is matched whole, so "Item Reference" is never read as "Item Ref" + "erence". A code in the same `-QWXK-L-` shape elsewhere on the label also counts.
- **A project with codes matches by code only**, so a label from another job can't land on the wrong packing label: it's listed as not placed.
- **A project generated before codes existed** falls back to the Item Ref text (`10 Label 2 - Q Eyewear`, the row number dropped) against the label map's item references.
- **Reported, not placed:** pages with no text (a scan), no Item Ref, or an unknown code, plus the same label twice (the first is used). Packing labels left without a courier label are listed.
- **A code from another project** (`codes_elsewhere()` reads the other projects' label maps, only when needed) is named: "code of project Lux test 33 (booked from its Open360 CSV)". Every Generate makes new codes, so courier labels booked from one project's CSV don't match another project's packing labels; the report's page 1 sums it up ("42 labels of project Lux test 33: book this job from this project's own Open360 CSV, or match them by hand").

**Packs of several boxes.** A Packing Spec with a count in front (`2 x OB170170170`, `2x OB…`, `2 xOB…`, `2 OB…`, `3 × CS…`; `2OB170170170` only when the rest is a formula code, so `3D Sign` stays one box: `split_count()` in packing_specs.py) is that many boxes. Generate draws **one full set of packing label pages per box**, each with a diagonal see-through **"BOX 1 OF 2"** watermark across the item grid (region C) and a **BOX 1 OF 2** chip beside LABEL / PAGE, so the packers make that many boxes. The Open360 CSV keeps one row for the pack with Item Quantity 2 (the portal prints a courier label per box, all with the same Item Ref code). The label map gives the pack `boxes` and `box_pages` (each box's first page); stitching treats each box as its own place for a courier label (`packing_slots()`, ids like `Label 1 - Store · box 2 of 2`): courier labels with the pack's code fill its boxes in turn, and only one more than the boxes is a duplicate.

**Two steps: match, then stitch** (`plan_stitch()`, then `render_stitch()`). Pressing **🧵 Stitch Labels** reads the text of every courier label page of every PDF at once and matches it (codes, then the matches saved by hand) — nothing is drawn, so it takes a fraction of a second (0.11 s for the 44 real portal labels; 0.06 s for 200 pages). Then:
- **Everything matches** (every courier label has its packing label, and every packing label that should get one has one): the Complete Labels PDF is drawn straight away and the page shows just **⬇ Download** and **📂 Open file** (opens it in the PDF viewer, `/open_local`).
- **Something doesn't:** straight to the matching page; nothing is drawn. There, **💾 Save changes** keeps the pairs (`/stitch/<project>/match`, instant) and **✅ Finalise** draws the PDF once (`/stitch/<project>/finalise`, in the background behind an overlay), then offers Download and Open file. Finalise asks first whether courier labels still matching nothing go 4-up at the end or are left out.
- **One Complete Labels file per project:** a new one replaces the earlier ones (`_remove_old_labels()`); one that's open in a PDF viewer and can't be deleted is named on the page. Adding courier label PDFs on the matching page removes the earlier Complete Labels (they're out of date) and matches again; if everything then matches, it's stitched straight away.

**Placing** (`render_label()`, `place_label()`):
- The courier label page is rendered as a **lossless 600 dpi grayscale PNG**: a 0.25 mm bar is 6 pixels wide, so Code 128 and QR codes stay scannable. The blank margin is cropped off, so a label in the corner of an A4 sheet isn't shrunk.
- It's **turned so its text reads upright**, judged from the text's own direction plus the page's rotation flag. So a 4 × 6 label printed sideways on a landscape page comes back the right way up. A genuinely landscape label is then turned a quarter to fit the portrait courier region, instead of shrinking to its width.
- It's **scaled to fit** the courier region from the label map (107 × 150 mm), aspect kept, centred and top-aligned, covering the dashed placeholder. A 4 × 6 portal label comes out at about 105–113 % once its margin is cropped.
- **Measured** on the real portal PDF: all 42 labels placed in about 20 s, and every Code 128 decoded at 300 dpi identical to the portal's own label. The QR code on those portal labels is cut off at the page edge in the portal's own PDF, so it can't be read there either.

**The Complete Labels PDF, in order:**
1. **The job report** (page 1, `draw_report_pages()` in packing_label_generator.py), kept **compact**: the job number in a black panel with the project name beside it (no caption) and when it was stitched; one row of chips (courier labels placed *of all packing labels*, **without a courier label** in red when any, not-matched and duplicates when any, the pages of the labels and of the 4-up, and labels per courier PDF); then each finding as a coloured heading with its items run together on one paragraph. The findings: packing labels without a courier label (**including packs left out of the courier CSV**, marked "not in the courier CSV", so they're never missed), courier labels that matched nothing (added 4-up with the cut-and-stack order, or left out), duplicates, and what Generate found. Generate's "Not in the courier CSV" section isn't repeated, and an older project's installer list is shortened to "Installer x6" (`_compact()`). It replaces Generate's own report page.
2. **Every packing label, in the packing labels' own order**, with its courier label where one was found. Matched and unmatched labels aren't grouped, so the pages follow the Packing Labels Only PDF one for one, after the report.
3. **Courier labels that matched no packing label** (no Item Ref / Item Reference, or a code from another job), if chosen. They're printed 4-up on pages the size of the packing labels, with the same 4 mm margin round each quarter and dashed cut guides, in **cut-and-stack order** (`four_up_order()`). With P pages, the top-left quarters hold labels 1…P, top-right P+1…2P, then bottom-left and bottom-right. Cut the printed stack into quarters and put the piles together top-left, top-right, bottom-left, bottom-right, and the labels are back in file order.

**Courier labels that still match nothing: added or left out?** Finalise asks ("Add them 4-up at the end" / "Leave them out"). The report says which was chosen.

**Matching by hand** (`/stitch/<project>/match`, [stitch_match.html](../templates/stitch_match.html), `script.js` section 16). Opened straight from Stitch when something doesn't match (or **🔗 Continue matching** on the Stitch page). It lists every courier label the codes didn't place, including ones matched by hand before (`match_session()`: they come back as matched pairs, `matched_to`, and can be changed):
- **Pairs** (top): each courier label sits **side by side** with its packing label (courier thumbnail on the left, click to zoom; packing label on the right), grouped by receiver and address ("N courier labels ↔ M packing labels · K to check"). The images are rendered on demand from the uploaded PDF (`/stitch/<project>/label/<n>.png`), so nothing extra is saved.
- **Suggested pairs start already paired, to check** (`suggest_pairs()`). Each courier label is scored against every waiting packing label (`pair_score()`): the **address printed on it** (`address_score()`: postcode, suburb, street numbers and receiver name) and, when it has an Item Ref, **the store in the Item Ref** (`ref_parts()` reads `42T1LCQK Blue Star Moorabbin`, `01-QWXK-L-1-Store` and `10 Label 2 - Store`; `_store_score()`, a store name the portal ran into the code included). Several stores' packs can go to one address (an installer such as Andrew Dalgleish taking Blue Star's and Eyedentity's packs), and the Item Ref says which store each courier label was booked for. Its serial (`42T1`) or label number (`L-2`, `Label 2`) adds a little, to tell one store's labels apart. Packing labels at one address for one store form a group; the clearly best group wins (by 0.12), and its packing label named by the serial / label number is taken first, else the next in label order. On the real labels of Lux test 36 stitched onto Lux test 37 (booked from the other job's CSV, so no code matches) this took the suggestions from 24 right and 18 wrong to all 42 right. A pair to check has dashed amber boxes, a "suggested %" tag, a **✓** and a **✕**.
- **Courier cards** show the address printed on the label in **bold** and its **Item Ref in bold red** (the file and page are in the thumbnail's tooltip), with a short reason underneath.
- **✓ confirms a pair**, where it is: the gap closes and the two boxes join into one green box (a CSS transition, the row changed in place, not re-drawn). The page then scrolls by exactly the distance from the clicked ✓ to the next ✓ to check, so it lands **under the pointer**: click, click, click in one spot. Focus goes to the next ✓ too, so Enter, Enter, … works from the keyboard. (The site sets `scroll-behavior: smooth`, so scrolls meant to be instant say `behavior: 'instant'`.) Hovering over a matched pair shows **↶ Undo**, which puts it back to check (or, for a pair made by hand, sends both labels back to Still to match). Motion is skipped for people who've asked their system for reduced motion.
- **✓ Match all 100% (N)** (green, in the toolbar and beside the Pairs heading) confirms every suggestion that scored 100% at once, leaving the amber ones that need a look. At (0) it's greyed out, with a tooltip saying why.
- **Still to match** (below): every courier label with no pair sits beside its **closest packing label still free** ("closest 60%"), with a **✓** to match them and **Other…** to pick a different packing label (closest addresses first, then all the others). `suggest_pairs()` gives each courier label up to 8 `candidates` ([item reference, score], every packing label of a place in label order); two courier labels never show the same packing label. After a ✓ there, the next row's ✓ is kept under the pointer. Packing labels that are nobody's closest are listed under **Other packing labels waiting** (store, "Label X of Y", Packing Spec, receiver and Attn, full address, Installer / Not-in-courier-CSV tags). All packing-label details come from the label map JSON, not from the packing labels PDF.
- **Pairing by hand:** drag a courier label (from either area) onto any packing label, or click one and then the other; the new pair is matched straight away. A packing label holds one courier label, so dropping onto a paired one swaps it, and the old courier label returns to Still to match. **↺ Suggested pairs** puts back the suggestions whose labels are free; **Clear all pairs** splits everything.
- **Search** narrows both areas. When the unpaired courier labels and packing labels shown are the same number, **Match these N in order** pairs them all at once (for a receiver the suggestions couldn't place).
- **Save changes** sends every pair, matched or still to check, as `{courier page key: slot}`; for the labels on this page they replace what was saved before in `stitch_matches.json` (so a pair undone and saved is dropped). Nothing is stitched. **Finalise** sends them too, then stitches once in a background thread. A full-page overlay covers the page until it's done (a second click does nothing). The page asks `/stitch/<project>/status?wait=1`, which **waits** (up to 25 s, on a `threading.Event` set when the stitch finishes) before answering, so it's one request per stitch rather than one a second; those requests and the label pictures are also kept out of the terminal log (a filter on the `werkzeug` logger). Then it shows the result with a Download link. Adding courier label PDFs uses the same overlay.
- **Matches are remembered.** A courier page is known by `page_key()`, a hash of its text (or of a small render, for a scanned page), so the same label in a re-downloaded PDF is matched again on the next stitch. What's left after each stitch is kept in `stitch_session.json`.

**File names.**
- Generate writes **`Packing Labels Only <n> - <job number> - <project> - <yymmdd_HHMM>.pdf`**.
- Stitching writes **`Complete Labels <n> - <job number> - <project> - <yymmdd_HHMM>.pdf`** (`numbered_name()`) and removes the earlier ones, so the folder holds one. The job number is the Consignment Reference.
- The dashboard and the Stitch page list files **newest first**.

**The packing labels' own report** (page 1 of Packing Labels Only, `_generation_report()` in app.py) shows:
- the job number, project and when it was generated
- counts as chips: packing labels, courier labels to come, **not in courier CSV** (red, when any), consignments, installer packs, service
- what needs a look while preparing labels, each a heading with its items run together:
  - packs left out of the courier CSV
  - specs with no size
  - addresses new to the address book
  - courier checks
  - spreadsheet checks
  - the installer packs, as "Installer (Attn) x2" each

The label map stores it (`project.report`, `project.report_pages`) and counts its pages in every page number.

**Project names are at most 50 characters** (`PROJECT_NAME_MAX`, `_project_name()` in app.py; `maxlength="50"` on the Distribution Mapper, Label Shuffler and Packing Labels boxes, with a "n / 50 characters" count, `script.js` section 20). The name is in the folder name and in every file name, and Excel can't open a file whose full path is longer than about 218 characters: at 100 characters a job's CSV path was 288 and wouldn't open; at 50 the longest path in a project is 216.

### 10.14 Pages that survive a refresh

Every upload page (Distribution Mapper, Label Shuffler, Packing Labels, Stitch Labels, Address Book update) keeps its place on a refresh:
- **A page shown after a form post is kept** (`_show()` in app.py): its template and data go into `data/page_views/<id>.json` (kept 2 days) and the browser is sent to a plain GET of it, `?view=<id>` (a 303 redirect). Refreshing reopens the same page (`_kept_view()`) without sending the form again: nothing is uploaded, previewed, generated or stitched twice, and there's no "resend form?" prompt. Each post gets its own id, so pages never overwrite each other. Only full-page form posts do this: the pages' background calls (Save changes, Finalise, adding PDFs, the address and sender lookups, status checks) answer directly as before, and a form sent from a kept page posts normally (the `?view=` is ignored for posts).
- **An upload isn't deleted by a refresh.** Leaving a page with an upload not yet used still asks for it to be deleted (discard-on-leave), but `_discard_later()` waits `DISCARD_GRACE` (2 minutes): a refresh comes straight back, and a page that shows the upload again cancels the deletion (`_keep()`).
- **Files picked but not sent yet** (any upload box, `.dropzone-input`) are kept in this browser (IndexedDB, `script.js` section 8b) until the form is sent or the box is cleared, a day at most, and put back in the box after a refresh ("kept from before the refresh").

### 10.15 The log: what happened, step by step (`app_log.py`)

Every step of Packing Labels, the address book, Packing Specs and Stitch Labels writes a line to the log, so a run that went wrong can be traced.

- **Terminal:** INFO and up, one line per step, with counts and timings, e.g.
  ```
  22:13:25 INFO    wh.packing|   tab VIC (header row 2): 43 packs, 35 stores, last row 109, 2 warning(s)
  22:13:25 INFO    wh.book   | Address book: 31 Excel addresses, 31 new to ask about, 22 filled from the book
  22:13:26 INFO    wh.packing|   Open360 CSV J477161 - ZZ_Match - Open360.csv done in 0.04s
  22:13:51 INFO    wh.stitch | Stitch done in 23.80s -> Complete Labels 1 - ...: 30 placed, 12 courier label(s) unmatched, ...
  ```
  Set `WAREHOUSE_LOG_LEVEL=DEBUG` before starting the app to see everything in the terminal too.
- **`logs/warehouse.log`:** everything, DEBUG included: each courier label page (`p31: Label 2 - … (matched code QWXK) -> packing label page 32 (111%)` or why it wasn't placed), each address filled from the book, each Open360 Item Reference, each pair saved on the matching page. 2 MB a file, the last 3 kept; not tracked by git.
- **Loggers:** `wh.packing` (upload, preview: which saved version of the file was read and whether it was open in Excel, per tab consignments and address-book status, Generate: left-out packs, PDF pages, CSV rows, label map), `wh.book` (lookups, fills, bulk update read/apply with problems, add/edit/delete), `wh.specs`, `wh.stitch` (each PDF and page, the summary, unmatched labels with the reason, packing labels without a courier label, duplicates, matching-page saves and added PDFs).
- **Failures** are logged with the full traceback (`log.exception`), including Generate (whose half-made project is undone), a stitch in the background thread, and a courier label image that couldn't be drawn. `app_log.step(log, "what")` times a step and logs its failure.

## 11. `templates/*.html` — The Web Pages

**Jinja2 primer, in three rules**, since every template leans on this:
- `{{ some_python_variable }}` prints a value.
- `{% if condition %} ... {% endif %}` / `{% for x in list %} ... {% endfor %}` are control structures — indistinguishable in spirit from Python's own `if`/`for`, just wrapped in `{% %}` instead of a colon+indent.
- `{{ url_for('route_function_name', arg=value) }}` asks Flask to build the correct URL for a given route function — so if a route's URL pattern ever changes in `app.py`, every template using `url_for` updates automatically; nothing is hard-coded.

### 11.1 `index.html` — the Dashboard
Renders the `projects` list built in `app.py`'s `/` route as a list of collapsible "accordion" cards (the actual expand/collapse behavior lives in JS, [Section 12.2.B](#12-staticscriptjs--the-frontend-brain) — this template just marks up the structure with `.accordion-header` / `.accordion-content` classes and lets CSS start them hidden `style="display: none;"`).

The most interesting piece is the **Create Sub-Group modal**, one per project, using the native HTML `<dialog>` element:
```html
<a onclick="event.preventDefault(); document.getElementById('modal-{{ project.name }}').showModal()">✂️ Create Sub-Group</a>
<dialog id="modal-{{ project.name }}">
    <form action="{{ url_for('setup_subgroup', project_name=project.name) }}" method="GET">
        <select name="target_json">
            {% for j_file in project.json_files %}
                <option value="{{ j_file }}">{{ j_file | replace('.json', '.xlsx') }}</option>
            {% endfor %}
        </select>
        ...
```
`<dialog>.showModal()` is a **built-in browser API** — no JavaScript library needed to pop up a proper modal dialog with its own backdrop. The dropdown lets the user choose *which* metadata JSON to treat as the sub-grouping baseline (displayed with a `.xlsx` extension via the `replace` filter purely for readability, even though the actual value sent to the server is the real `.json` filename) — this becomes the `target_json` GET parameter that `app.py`'s `/subgroup` route reads.

Projects that contain a courier CSV show a **🚚 Courier CSV** file group and a disabled **🧵 Stitch Labels** button (`.btn-stitch`, tooltip "Coming soon") **in place of** the Create Sub-Group button and its modal.

### 11.2 `matrix.html`
One template, four possible states, controlled entirely by which Jinja variables `app.py` decided to pass in:

| Passed in from `app.py` | What's shown |
|---|---|
| nothing (`tabs` is `None`) | The instructions box + upload dropzone |
| `tabs` (list of sheet names) | One card per tab with 3 coordinate inputs |
| `previews` (blueprint list) | The blueprint cards + pack checkboxes + "Generate" button |
| `generation_complete=True` | Download links |

The upload dropzone (added when we made drag-and-drop work) is worth re-reading here since it's a pattern reused three times across the app:
```html
<div class="dropzone">
    <div class="dropzone-content"> ...icon, text, a blank <p class="dropzone-filename"> ... </div>
    <input type="file" name="file" class="dropzone-input" accept=".xlsx, .xls" required>
</div>
```
The `<input>` is positioned (via CSS, see [Section 13](#13-staticstylescss--the-look--feel)) as an **invisible layer covering the entire box**. That's the whole trick — there's no custom drag-and-drop JavaScript logic needed at all, because the browser already knows how to drag-and-drop a file directly onto a file input; making that input physically as big as the box just makes the *whole box* a valid drop target. `script.js` only adds the cosmetic touches (border highlight while dragging, showing the chosen filename) — see [Section 12.2.H](#12-staticscriptjs--the-frontend-brain).

**Two "Open Excel File" links, both reusing the exact same open-locally mechanism as the dashboard.** Once a file is scanned (`filename` is set), a small `<a class="btn-file open" href="{{ url_for('open_upload_file', filename=filename) }}">` sits right next to the "✅ File Detected" heading — always visible, so the user can glance at the raw spreadsheet before typing any coordinates. A second, identical link appears specifically inside the duplicate-store hard-lock warning (below), so the fix-and-recheck path doesn't require scrolling back up to the header. Both hit `/open_upload/<filename>` rather than `/open_local/<folder>/<filename>` because at this point in the wizard the file is still loose in `projects/` — no project folder exists yet (see [Section 4.4](#44-project--file-management-delete-delete_file-open_local-open_upload)). Being plain `<a class="btn-file open">` links, `script.js`'s generic fetch-interceptor (Section 12.2.D) picks them up automatically, with zero page-specific JavaScript needed.

The **hard-block logic** near the bottom is worth understanding since it silently disables the biggest button on the page:
```jinja
{% set lock = namespace(has_dupes=false) %}
{% for bp in previews %}
    {% if bp.duplicate_warning %}{% set lock.has_dupes = true %}{% endif %}
{% endfor %}
{% if lock.has_dupes %}
    <button disabled>Cannot Generate</button>
{% else %}
    <button type="submit">Generate Packing Sheet</button>
{% endif %}
```
Plain Jinja variables set inside a `{% for %}` loop don't survive past the loop (a Jinja quirk, similar to variable scoping oddities in some templating languages) — `namespace()` is Jinja's workaround: it creates a small mutable object whose attributes *do* persist outside the loop, which is why `lock.has_dupes` can be read correctly afterward.

The "Cannot Generate" lock is deliberately just a soft speed bump, not a dead end: the file is still sitting right there under `filename`, so the fix is "open it, fix the duplicate, save, click **Update Previews**" — that button resubmits `/preview` with the same `filename`, which re-reads the file fresh off disk (`generate_tab_map()` opens a brand-new `openpyxl.load_workbook`, no caching anywhere) and recomputes `duplicate_warning` from scratch. No dedicated "recheck" endpoint was needed here the way `/pdf` needed one — resubmitting the *existing* form already happens to do exactly the right thing.

### 11.3 `pdf.html`
Three step blocks, gated the same way (`{% if step == 1/2/3 %}`). The `duplicate-modal` `<dialog>` in Step 1 is auto-opened without any click, purely by `script.js`'s Section 9 (`dupeModal.showModal()` fires automatically on page load if the element exists) — because the server only renders that `<dialog>` into the page at all when `duplicate_errors` is non-empty, its mere *presence* in the HTML is the signal to pop it open immediately.

**The modal's footer holds three small buttons, all sharing the `.btn-file` sizing convention** (see [Section 13](#13-staticstylescss--the-look--feel)) so they read as one coherent group instead of three mismatched controls:
```html
<button class="btn-file ghost" onclick="...close()">Close</button>
<a class="btn-file open" href="{{ url_for('open_local_file', ...) }}">📂 Open Excel File</a>
<form method="POST" style="display: contents;">
    <input type="hidden" name="existing_project" value="{{ duplicate_project_name }}">
    <input type="hidden" name="resume_filename" value="{{ duplicate_excel_filename }}">
    <button type="submit" class="btn-file primary">🔄 Recheck File</button>
</form>
```
Ordered left-to-right from least to most important, matching the usual modal-footer convention (dismiss on the left, primary action on the far right where the eye lands last): **Close** (neutral, `.btn-file.ghost` — dismisses without doing anything), **Open Excel File** (`.btn-file.open`, the same class/route/fetch-interception as everywhere else in the app — see [Section 11.2](#112-matrixhtml)), **Recheck File** (`.btn-file.primary`, solid green — the actual call to action). The `<form style="display: contents;">` wrapper is a small CSS trick: it makes the form's own box "disappear" from the flex layout while its child `<button>` still participates directly in the parent `.modal-actions-end` flex row, so the button lines up and gets the same `gap` spacing as its siblings, as if the form tag weren't there at all. `duplicate_project_name` / `duplicate_excel_filename` (only present when the failed upload/recheck happened against a real project folder) are what let "Recheck File" resubmit Step 1 pointing at the *exact* file that was just flagged, via the `resume_filename` mechanism described in [Section 4.6](#46-the-pdf-label-shuffler-route-pdf).

Step 2's per-pack layout is the "2-column grid, grouped per tab" structure:
```html
{% for tab, packs in tabs_data.items() %}
<div class="tab-card">
    <div class="pack-grid">
        {% for pack in packs %}
        <div class="project-name-group"> ...one compact dropzone per pack... </div>
        {% endfor %}
    </div>
</div>
{% endfor %}
```
The **outer loop** (`tab_card`) keeps packs visually grouped under their own tab; the **`pack-grid`** CSS class (a `display: grid; grid-template-columns: repeat(2, 1fr)`) only affects layout *inside* that group, so packs from different tabs are never mixed into the same row.

### 11.4 `sub-group.html`
The shortest template by far — because almost the entire form is **built dynamically by JavaScript**, not by Jinja. The template only renders:
1. A checkbox per tab (from `metadata.tabs.keys()`).
2. An empty `<div id="tab-blocks-container">` — JS fills this in the moment a tab checkbox is checked.
3. A hidden `<script id="meta-data" type="application/json">{{ metadata_json | safe }}</script>` — this is how the *entire* metadata dictionary (which packs exist, which were previously selected, etc.) gets from Python into JavaScript's hands: it's dumped as a JSON string and just sits inertly in the page's HTML until `script.js` reads and `JSON.parse()`s it.

The `error` banner at the top (added for the item-number validation fix) is the one piece of genuinely server-rendered feedback on this page — everything else here is either static structure or JS-generated.

### 11.5 `packing_labels.html`

One template, gated by which variables `app.py` passes — the same idea as `matrix.html`:

| Passed in from `app.py` | What's shown |
|---|---|
| nothing | Requirements box + upload dropzone |
| `tabs` | One card per tab with a Header Row Number input |
| `previews` | Per-tab result: metrics and warnings, or the issue table (rendered by the `issue_table` macro at the top of the file) |
| `previews` with no errors | The two-column layout editor (1. Mapped Headers / 2. Customize Cell Layout) |
| `courier` | 3. Courier Consignments: reference, Who Pays, Charge Account, Service Code dropdown + "Apply to shown", State and text filters, and the consignment table with ✏️ edit rows and a service-code dropdown per row |
| `generation_complete=True` | Download buttons for the label PDF and the courier CSV |
| `page_error` | A red "Could Not Process File" card |

Things worth knowing:

- **The layout editor** draws each block exactly as it will print (black job-number bar, boxed quantity…). Every block carries two looks: a short `.badge-text` name shown in the left "available" list, and a `.cell-only-styling` preview shown once dragged into the cell; CSS switches between them depending on which list the block is in. The final order travels in the hidden `attribute_order` field.
- **No styles inside the template.** Every style lives in `styles.css` under `.pl-page` ([Section 13](#13-staticstylescss--the-look--feel)).
- **No Jinja inside `<script>` tags.** Values JavaScript needs are written into HTML instead: `data-` attributes (`data-id`, `data-state`, `data-original` on each consignment row) and hidden inputs (`consignment_edits`). The page-leave cleanup reads `<div id="discard-on-leave" hidden data-url data-field data-value>`.
- Every edit-row input has a stable `id` (`edit-<n>-<field>`) and a `<label for>`, and the per-row dropdowns have `aria-label`s, so the table works with a keyboard and screen readers.

---

## 12. `static/script.js` — The Frontend Brain

### 12.1 The pattern used almost everywhere: **event delegation**

Instead of finding a specific button and attaching a listener to it directly, nearly every handler in this file is attached once to `document.body`, and then checks *what was actually clicked* using `e.target.matches(...)`:
```js
document.body.addEventListener('click', function(e) {
    const header = e.target.closest('.accordion-header');
    if (header) { /* ...do the accordion toggle... */ }
});
```
```mermaid
flowchart LR
    A["User clicks ANYTHING on the page"] --> B["The one 'click' listener on document.body fires"]
    B --> C{"Does e.target (or an ancestor) match\n a selector we care about?"}
    C -- yes --> D[Run that specific handler's logic]
    C -- no --> E[Do nothing]
```
**Why this matters:** many elements on these pages (pack checkboxes, "Remove" buttons, tab cards) don't exist yet when the page first loads — they're created later by JavaScript itself (e.g. when a tab checkbox is checked in `sub-group.html`). A listener attached directly to an element that doesn't exist yet would simply never fire. A listener on `document.body` is always there from page-load, and `e.target.closest(...)` / `.matches(...)` correctly catches clicks on elements added at any point afterward — no need to re-attach listeners every time new HTML is injected.

### 12.2 Section-by-section walkthrough

**0. Discard Ungenerated Uploads on Leave.** Runs on any page containing `#discard-on-leave`. It listens for `pagehide` (the page is being left). If the page is left **without** one of its own forms being submitted, and the browser isn't keeping it in the back/forward cache (`e.persisted`), it sends the `data-field`/`data-value` pair to the `data-url` with `navigator.sendBeacon()` — a small request the browser delivers even while the page is closing. This is how scanned-but-never-generated uploads get deleted ([Section 4.4](#44-project--file-management-delete-delete_file-open_local-open_upload)).

**1. Global Utilities.** Intercepts clicks on any `<a href="#something">` link and smooth-scrolls to that element instead of the browser's default instant jump.

**2. Dashboard Logic (`index.html`).**
- **A. Scroll Spy** — toggles the `.active` class between the "Home" and "Dashboard" nav links based on scroll position, so the nav bar reflects which section you're currently looking at.
- **B. Accordion Toggle** — the project-card expand/collapse. Reads the clicked header's very-next-sibling element (`header.nextElementSibling`) and flips its `display` between `none` and `block`, also flipping the ▼/▲ icon.
- **C. Preserve State After Deletions** — a UX nicety: normally, deleting a file causes a full page reload, which would snap the scroll position back to the top and collapse whichever project card you had open. Just before a delete form submits, the current scroll position and the currently-open project's ID are written into `sessionStorage` (a small key-value store the browser keeps per-tab, surviving a page reload but not a tab close). On the *next* page load, this saved state is read back and the page immediately re-scrolls and re-opens that same card, and then the saved values are deleted — so the illusion is "the delete happened in place."
- **D. "Open Locally" Fetch Requests** — clicking "📂 Open" doesn't navigate the browser at all; it fires a background `fetch()` to `/open_local/...` (which just tells *your Windows machine* to open the file) and shows an `alert()` only if that request failed. The selector is `a.btn-file.open` generically (not tied to the dashboard specifically), which is exactly why the newer "Open Excel File" links added to `matrix.html` and the duplicate modal in `pdf.html` (pointing at `/open_upload/...` or `/open_local/...` respectively) get this same no-navigation behavior automatically, with no extra JavaScript written for either page.

**3. Sub-Group Engine (Dynamic UI).** The most complex section — it manually builds HTML strings and injects them with `insertAdjacentHTML` in response to checkbox changes, mirroring what a template engine would normally do, but at runtime in the browser instead of on the server:
```mermaid
flowchart TD
    A["User checks a TAB checkbox"] --> B["Look up this tab's packs in the meta object\n(parsed from the hidden JSON script tag)"]
    B --> C["Build an HTML block: Item Number Row input\n+ one checkbox per PREVIOUSLY-SELECTED pack"]
    C --> D[Inject it into #tab-blocks-container]
    D --> E["User checks a PACK checkbox inside that block"]
    E --> F["Inject a Start/End Item # row for that pack\n+ an 'Add another sub-group' button"]
    F --> G["User clicks 'Add Another Sub-Group'"]
    G --> H[Append one more Start/End row]
```
Only packs where `packData.is_selected` is `true` are offered here — i.e., only packs that already have a signature code column from the original Matrix Engine run can be sub-grouped; a pack nobody selected the first time around has nothing to slice further.

**4. Upload Form (Step 1 of `pdf.html`).** A `validateForm()` function re-run on every relevant keystroke/change, encoding these rules:
- Typing in the "new project" text box **disables** the dropdown (and vice-versa) — you can't do both at once.
- If an existing project is picked and it already has a Signature Links file on disk (`projectData[dropdown.value]` truthy — this data came from the `projects_json` script tag, same trick as the sub-group metadata), uploading a new file becomes optional and the submit button enables immediately.
- Otherwise (new project, or existing project with no file yet), a file **must** be attached before the button enables.

**5. PDF Label Shuffler (Step 2).** Two small independent-but-linked behaviors: each per-pack file input shows/hides its own "add divider sheets" checkbox the moment a PDF is attached (and un-checks it if the file is removed), while a single master checkbox at the bottom can check/uncheck every *currently visible* per-pack checkbox at once (only ones whose container is actually showing, i.e. only packs that have a PDF attached).

**6. Matrix Anchor Points Memory.** Pure convenience: every tab's Start Cell / Job ID Cell / Store Column values are saved to `localStorage` under `anchor_points_by_tab` each time the `/preview` form is submitted. `localStorage` persists across browser restarts, unlike `sessionStorage`. The saved values replace the hard-coded defaults (`B8`, `E1`, `A`) the *next* time you upload a file, so you stop retyping the same coordinates every job. Each tab is filled from its own memory: first **by tab name** (`byName`, the same file uploaded again, up to 200 tab names kept), then **by position** (`byIndex`, the last upload's tabs in order, for a new campaign file whose tab names changed), then the last upload's first tab. Browsers that only have the older single-value keys (`anchor_start` / `anchor_job` / `anchor_store`) fall back to those. Only fields still showing a built-in default are filled, so values the server re-rendered after "Update Previews" are never overwritten.

**7. Auto-Scroll to Previews.** After generating blueprint previews, a `setTimeout(..., 150)` waits briefly for the page to finish rendering, then smooth-scrolls down to the `#preview-section` anchor — otherwise the user would submit the coordinate form and land back at the top of a long page, with no visual cue that anything happened below the fold.

**8b. Drag & Drop File Inputs.** As explained in [Section 11.2](#112-matrixhtml), the actual drop behavior is native/free. This code only adds:
```js
input.addEventListener('dragenter', () => zone.classList.add('dropzone-active'));
input.addEventListener('dragleave', () => zone.classList.remove('dropzone-active'));
```
...for the blue-highlight-while-dragging effect, and a `change` listener that writes the selected file's name into the little `<p class="dropzone-filename">` placeholder — run once immediately (`showFilename()` called right after being defined) in case a browser ever pre-fills a file input on page restore.

**8. Loading Overlays.** A catch-all: any form submission on the page shows the full-screen spinner overlay, **except** deletions and the `/preview` submission (both of those are meant to feel instant, not like a long background job — showing a spinner would be misleading since they usually resolve in well under a second).

**9. Duplicate Store Name Alert.** Notably declared **outside** the main `DOMContentLoaded` listener that wraps sections 1–8 — a small inconsistency in the file's history, but harmless since `DOMContentLoaded` listeners can be registered as many times as you like and all of them still fire. Its only job is calling `.showModal()` on the duplicate-store dialog the instant the page loads, if that dialog exists in the HTML at all.

**10. Label Maker Drag and Drop Configurator** (`packing_labels.html`). Native HTML drag events on every `.sortable-item`. While dragging over a list, `getDragAfterElement()` finds the block whose vertical middle is just below the mouse and inserts the dragged block before it. When a drag ends, the order of `data-id`s in the right-hand list is written into the hidden `attribute_order` field (e.g. `thumbnail,desc,job_no,qty`); empty slots are filtered out on the server.

**11. Courier Consignments** (`packing_labels.html`). Keeps every change in one object, `edits`, keyed by consignment id, and writes it as JSON into the hidden `consignment_edits` field after every change:
- **✏️** toggles the row's edit form; **Save** compares the form with the row's `data-original` values and stores only real changes (keeping any service code already chosen), updates the visible text, and shows the "Edited" tag. **Reset to Excel** refills the form with the original values. Enter saves and Escape closes, so pressing Enter never submits the whole page.
- **Service codes:** a row dropdown's `change` stores or clears its code; changing the main dropdown relabels every row's "Same as main (CODE)" option; **Apply to shown** copies the main code into every row not hidden by the filters.
- **Filters:** the State dropdown and text box hide non-matching rows (and their open edit forms), update the "Showing N of M" count, and relabel the button "Apply to N shown".

---

## 13. `static/styles.css` — The Look & Feel

Rather than walking every single CSS rule (CSS has no "logic" to trace — it's declarative styling, one selector at a time), here's what each named block of the stylesheet is *for*, grouped by the UI component it styles:

| CSS section | Powers |
|---|---|
| **Global Styles** | Base font, page background, the white rounded `.container` card every page sits inside. |
| **Navigation** | The dark sticky top bar (`.sticky-nav`) present on every page, and its active-link underline. |
| **Index/Dashboard** | The two big home-page "portal" cards (Create Packing Sheets / Label Shuffler). |
| **Projects List** | The dashboard's project cards, the auto-wrapping button grid (`.actions-grid` — `repeat(auto-fit, minmax(180px, 1fr))` means "as many equal-width columns as fit, each at least 180px, wrapping automatically on narrow screens"), and the color-coded action buttons (blue=download, red=delete, green=PDF). |
| **Matrix Engine Styles** | `.dropzone` / `.dropzone-compact` / `.dropzone-active` (the drag-and-drop boxes), `.pack-grid` (the 2-column PDF-pack layout), `.tab-card`, `.input-field`, `.blueprint-card` / `.blueprint-error` (preview cards, green vs. red left border), `.metrics-bar`, `.project-name-group`. |
| **Loading Overlay** | The full-screen dark spinner shown during long operations; `@keyframes spin` is what makes the spinner ring actually rotate. |
| **Tooltips** | The `.tooltip-btn::before`/`::after` pair — pure-CSS hover tooltips (a little triangle + a label bubble) with no JavaScript at all, using the `content: attr(data-title)` trick to pull the tooltip text straight from an HTML attribute. |
| **Instructions & Info Boxes** | The blue "Required Spreadsheet Formatting" box on `matrix.html`, and the red `.warning-box` used for duplicate-store warnings. |
| **Accordion & File List** | The dashboard's expandable project sections and the individual file rows (with their Open/Download/Delete button trio, `.btn-file` + `.open`/`.download`/`.delete`). |
| **Sub-group Tab Selection** | The responsive checkbox grid (`.tab-checkbox-grid`) in `sub-group.html`. |
| **Duplicate Store Error Modal** | Styles the `<dialog>` popup shown in `pdf.html` Step 1, and `matrix.html`'s hard-lock warning. `.modal-actions-end` (`display: flex; justify-content: flex-end; gap: 10px;`) is the shared footer-row layout for modal buttons. |
| **Packing Labels page** (end of the file) | Everything for `packing_labels.html`: the layout editor, fixed badges, the issue tables, the consignment table, filters, edit rows, service-code dropdowns. All rules are written as `.pl-page .pl-…`. |

**A pattern worth noticing:** almost every "card" component in this app (`.tab-card`, `.blueprint-card`, `.project-card`) shares the same visual recipe — white/light background, `border-radius: 6-8px`, and a thick colored **left border** (`border-left: 5px solid <color>`) used as a quick color-coded status signal (blue = neutral/info, green = success, red = error) — rather than one shared CSS class, this recipe is just repeated per-component, so if you ever want to change "the card look" everywhere, you'd currently need to update several rules instead of one.

**The `.btn-file` family** is the small-button system reused across the dashboard, `matrix.html`, and `pdf.html`'s duplicate modal — one base class (`padding: 6px 12px; font-size: 13px; border-radius: 4px;`) plus a color modifier: `.open` (light blue — opening a file), `.download` (neutral gray), `.delete` (light red), `.primary` (solid green — the modal's main call-to-action, e.g. "Recheck File"), `.ghost` (white/gray outline — a modal's dismiss action, e.g. "Close"). Before `.primary`/`.ghost` existed, "Recheck File" borrowed `.btn-generate` (a *different*, much bigger full-width CTA class meant for standalone buttons like "Scan File") and "Close" had no matching class at all, so it fell back to the page's generic `button { background: #3498db; }` rule — that mismatch (an oversized green button next to a plain-blue one, on two separate stacked rows) is what made the modal look inconsistent before these two modifiers were added specifically to let every modal button share one small, consistent size and differ only by role color.

**Why `.pl-page .pl-…` (two classes)?** The Packing Labels styles used to be inline `style="…"` attributes, which beat every stylesheet rule. When they moved into `styles.css`, some older rules — e.g. `.blueprint-card h4` (a class plus an element) — were *more specific* than a single new class and would have won, changing the look. Prefixing every rule with the page wrapper's `.pl-page` class gives it enough weight. The move was checked by recording every element's computed style in headless Chrome before and after: zero differences. One more rule, `.pl-page [hidden] { display: none !important; }`, makes the HTML `hidden` attribute beat any `display` rule (otherwise a tag styled `display: inline-block` stays visible even when hidden).

**`.btn-stitch`** styles the dashboard's purple Stitch Labels placeholder; `:disabled` greys it out.

---

## 14. End-to-End Journeys

### Journey A: Creating a Packing Sheet from scratch

```mermaid
sequenceDiagram
    participant U as User
    participant App as app.py
    participant ME as matrix_engine.py
    participant CM as core_math.py
    participant XL as Excel (via xlwings)
    participant FS as projects/ folder

    U->>App: POST /matrix (upload .xlsx)
    App->>ME: scan_excel_tabs()
    ME-->>App: ["Store A", "Store B"]
    App-->>U: tab cards + coordinate inputs
    U->>App: POST /preview (coordinates)
    App->>ME: generate_tab_map() per tab
    ME-->>App: blueprint (packs, warnings)
    App-->>U: blueprint preview + pack checkboxes
    U->>App: POST /generate (packs + project name)
    App->>FS: create project folder, move upload
    App->>ME: generate_all_outputs()
    ME->>CM: generate_pack_signatures() per pack
    ME->>XL: insert columns, write letters, format cells
    ME->>FS: save File1 (.xlsx), File2 (.xlsx), File3 (.xlsx), metadata (.json)
    App-->>U: download links
```

### Journey B: Sub-grouping an existing project

```mermaid
sequenceDiagram
    participant U as User
    participant App as app.py
    participant Sub as subgroup_engine.py
    participant CM as core_math.py
    participant XL as Excel (via xlwings)
    participant FS as projects/ folder

    U->>App: GET /subgroup/PROJECT_NAME (from dashboard modal)
    App->>FS: read chosen metadata .json
    App-->>U: sub-group.html (JS builds form from metadata)
    U->>App: POST /subgroup/PROJECT_NAME (item ranges)
    App->>Sub: execute_subgroups()
    Sub->>Sub: Stage 1: map item row → columns (validate!)
    alt validation fails
        Sub-->>App: raise SubgroupValidationError
        App-->>U: re-show form + error banner (nothing saved)
    else validation passes
        Sub->>CM: generate_pack_signatures() per sub-range
        Sub->>XL: Stage 2 insert columns, Stage 3 build matrix
        Sub->>FS: save Stage-N .xlsx + Stage-N .json
        App-->>U: redirect to dashboard
    end
```

### Journey C: Shuffling PDF labels

```mermaid
sequenceDiagram
    participant U as User
    participant App as app.py
    participant PE as pdf_engine.py
    participant FS as projects/ folder

    U->>App: POST /pdf (step=1: pick/create project + Signature Links Excel)
    App->>App: pandas reads every tab, checks for duplicate store names
    alt duplicates found
        App-->>U: Step 1 again + duplicate-names modal (nothing deleted)
        U->>App: fix + save the file, click "Recheck File" (resume_filename)
        App->>App: re-read the SAME file fresh off disk
    else clean
        App-->>U: Step 2 (one dropzone per Tab x Pack)
        U->>App: POST /pdf (step=2: attach PDFs + divider checkboxes)
        App->>App: build store→code mapping per Tab/Pack
        App->>PE: process_and_shuffle_pdf() per attached PDF
        PE->>PE: detect layout, match store names, build audit report
        PE->>FS: save "..._Shuffled.pdf"
        App-->>U: Step 3, download links
    end
```

### Journey D: Packing labels and the courier CSV

```mermaid
sequenceDiagram
    participant U as User
    participant App as app.py
    participant PL as packing_label_generator.py
    participant CE as courier_export.py
    participant FS as projects/ folder

    U->>App: POST /packing-labels (upload .xlsx)
    App-->>U: tab cards + Header Row inputs
    U->>App: POST preview
    App->>PL: parse_packing_data() per tab
    alt PackCheckError
        App-->>U: issue table, Generate locked
    else clean
        App->>CE: build_consignments(edits)
        App-->>U: layout editor + consignment table
        U->>U: drag blocks, ✏️ fix addresses, pick service codes
        U->>App: POST generate
        App->>CE: validate + build consignments
        App->>FS: create project folder, move Excel
        App->>PL: generate_packing_labels() → PDF + page map
        App->>CE: write_open360_csv(), write_label_map()
        App->>FS: PDF, CSV, labelmap.json; usage counts in data/
        App-->>U: PDF + CSV download buttons
    end
```

---

## 15. Quick Reference: "I Want to Change X"

| I want to... | Look here |
|---|---|
| Change what counts as "zero" in a quantity cell | `core_math.py` → `sanitize_cell()` |
| Change how signature letters are assigned (e.g. numbers instead of letters) | `core_math.py` → `generate_pack_signatures()` |
| Change the default Start Cell / Job ID / Store Column | `templates/matrix.html` (the `value="B8"` etc. defaults) and `static/script.js` Section 6 (localStorage memory) |
| Change how packs are detected (merged cells vs single column) | `matrix_engine.py` → `generate_tab_map()` |
| Change what makes File 1 / File 3 skip generation | `matrix_engine.py` → the `any_packs_selected` flag in `generate_all_outputs()` |
| Change the sub-group item-number validation messages | `subgroup_engine.py` → `_map_item_columns()` and the two `raise SubgroupValidationError(...)` call sites |
| Change the app's port, or what happens to whatever is already using it at startup | `app.py` → `PORT` and `free_port()` |
| Change how long until old projects auto-delete | `app.py` → `clean_old_projects()` (`seven_days_in_seconds`) |
| Change the duplicate-store-name check for the PDF shuffler | `app.py` → `/pdf` route, Step 1 (`dupes = store_col[store_col.duplicated()]...`) — note it no longer deletes anything on a duplicate; see [Section 4.6](#46-the-pdf-label-shuffler-route-pdf) |
| Change the "Open Excel File" / "Recheck File" duplicate-fix flow | `app.py` → `/pdf` Step 1's `resume_filename` handling, and `templates/pdf.html`'s duplicate modal footer |
| Change how a PDF page is matched to a store (the text region it reads) | `pdf_engine.py` → the `fitz.Rect(...)` clip rectangles in `process_standard_layout` / `process_split_layout` |
| Change whether/how divider sheets look, or the green/red match logic | `pdf_engine.py` → `build_divider_sheet()` |
| Change the "Unmatched Pages" divider sheet | `pdf_engine.py` → `build_unmatched_divider_sheet()` |
| Change what an installer divider's barcode encodes (currently the job number, plus " KN" (e.g. " K1") for multi-kind jobs) | `matrix_engine.py` → `build_divider_barcode_rows()` (the last value of each row is the `Barcode` column) |
| Change how kinds are numbered | `matrix_engine.py` → `assign_job_kinds()` |
| Change barcode size limits, caption text, or the barcode divider layout | `pdf_engine.py` → the `BARCODE_*` constants, `barcode_caption()` and `_build_barcode_divider_sheet()` |
| Change the audit report's colors/text, or the Missing Stores column layout | `pdf_engine.py` → `build_audit_report()` and `group_missing_stores_by_code()` |
| Change the near-duplicate store-name collision detection (only flips a divider red now, no audit-report text) | `pdf_engine.py` → `find_name_collisions()` and `analyze_matches()` |
| Change the same-store-matched-more-than-once detection (standard layout only) | `pdf_engine.py` → `process_standard_layout()`'s `store_instance_count` / `is_continuation` logic |
| Change when the app force-closes a file the user has open in Excel | `core_math.py` → `close_if_open_elsewhere()` (called at the start of every Generate in `app.py`, from `subgroup_engine.py`, and by the upload cleanup); how windows are found → `_excel_windows()` |
| Change any button/card color or spacing | `static/styles.css` (grouped by component — see [Section 13](#13-staticstylescss--the-look--feel) table) |
| Change what happens when a form is submitted (spinner, validation) | `static/script.js` (find the relevant numbered section — see [Section 12](#12-staticscriptjs--the-frontend-brain)) |
| Add a brand-new page/route | Add a `@app.route(...)` function in `app.py`, a matching file in `templates/`, and link to it from `templates/index.html`'s nav bar |
| Change the street-type equivalences (Street = St …) | `packing_label_generator.py` → `STREET_TYPES` |
| Change which words mark a name as a company | `courier_export.py` → `COMPANY_WORDS` |
| Change when a street splits into Address Line 1 / Line 2 | `courier_export.py` → `LINE_LIMIT` |
| Change how a messy address is read | `courier_export.py` → `_split_name_address()` and `resolve_destination()` |
| Change sizes, weights or Item Type of packing specs | the **📐 Packing Specs** page; defaults for a fresh install: `packing_specs.py` → `DEFAULT_FORMULA`, `DEFAULT_NAMED` |
| Change how OB / CS / Pallet / FP codes are read | `packing_specs.py` → `parse_formula()`, `split_dimensions()` |
| Add or rename a courier service code | `courier_export.py` → `SERVICE_CODES` |
| Change the courier CSV columns | `courier_export.py` → `OPEN360_HEADER`, `open360_items()` and `write_open360_csv()` (format locked to the file the portal accepted) |
| Change the packing-label checks (errors vs warnings) | `packing_label_generator.py` → `_check_pack_consistency()` and the end of `parse_packing_data()` |
| Change the packing-label header keywords | `packing_label_generator.py` → the header loop in `parse_packing_data()` |
| Change packing-label margins, gaps or the courier region size | `packing_label_generator.py` → constants at the top of `generate_packing_labels()` |
| Change what a packing-label barcode encodes, or how kinds are numbered | `packing_label_generator.py` → `barcode_text()` and `assign_barcodes()` |
| Change the size of packing-label barcodes | `packing_label_generator.py` → the `'barcode'` branch of `_layout_cell()` |
| Change how blocks look inside a packing-label box | `packing_label_generator.py` → `TEXT_STYLES` and `_layout_cell()` (and the matching preview in `packing_labels.html`) |
| Change how old abandoned uploads must be before cleanup | `app.py` → `clean_old_projects()` (`24 * 60 * 60`) |
| Change what counts as an abandoned Label Shuffler folder | `app.py` → `is_abandoned_project()` |
| Change how the courier portal's verified CSV is read or matched | `address_import.py` → `HEADER_NAMES`, `map_columns()`, `build_review()`; how close a "likely match" must be: `address_match.py` → `MIN_NAME`, `MIN_ADDRESS`, `MIN_TOTAL`, `MIN_LEAD` |
| Change how the address book searches or ranks | `address_book.py` → `_match_query()`, `RANK`, `search()` |
| Change what the address book saves after Generate, or when it fills addresses in | `address_book.py` → `record_used()`; `app.py` → `_apply_address_book()` |
| Move the address book somewhere else (e.g. a cloud database) | replace `address_book.py`, keeping the `AddressBook` methods; the file path is set in `app.py` (`ADDRESS_BOOK`) |
