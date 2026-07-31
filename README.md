# 📦 Warehouse Packaging Automation Suite

A local Flask-based web application built to automate the generation of warehouse packing sheets and signature distribution matrices from raw Excel data.

This tool eliminates manual data entry by parsing raw distributions, mapping store allocations, and generating print-ready matrix files in seconds.

---

## ✨ Core Features

* **Matrix Engine:** Upload raw `.xlsx` files, select specific tabs, and visually map Excel coordinates before generating files.
* **Automated Outputs:** Generates up to three distinct files simultaneously:
  1. **Modified Master Sheet:** Injects signature code columns directly alongside selected pack ranges.
  2. **Aggregated Packing Sheet:** Unifies counts and signature codes into a clean summary view.
  3. **Signature Links Matrix:** Formats distribution mapping via Pandas and XlsxWriter.
* **Smart Dashboard:** Tracks recent jobs, auto-extracts Job IDs, handles multi-tab setups, and provides instant download links.
* **Auto-Cleanup & Protection:**
  * **Bulldozer File Deletion:** Custom `force_delete_handler` overrides Windows/Excel read-only locks to ensure seamless updates and manual deletions.
  * **7-Day Sweeper:** Automatically cleans old job folders from the local drive on startup.
  * **Ghost Column Detector:** Automatically scans Excel's `job_id_row` to truncate phantom formatted columns (e.g., empty `CS1` cells) and preserve layout integrity.
  * **Universal Sanitizer:** Cleans project and job names for Windows compatibility while safely retaining spaces.

---

## 🛠️ Tech Stack

* **Backend:** Python 3, Flask, Werkzeug
* **Data Processing:** Pandas, OpenPyXL, xlwings, XlsxWriter
* **Frontend:** HTML5, CSS3, Jinja2 Templates

---

## ⚠️ Important System Requirements & Best Practices

1. **Microsoft Excel:** Local installation of Microsoft Excel is strictly required because `xlwings` uses Excel's native engine to render complex grid modifications.
2. **Local Directory:** Run this application strictly from a local directory (e.g., `C:\Warehouse_app`). **Do not run inside cloud-synced folders** (OneDrive, SharePoint, Dropbox), as cloud engines lock newly created Excel files and crash cleanup routines.
3. **Data Hygiene:** The application is built to automatically detect true sheet boundaries. However, keeping input files trimmed of unused rows/columns is recommended for maximum processing speed.

---



## 🚀 Installation & Setup

### Using VSCode

### 1. Clone the Repository
```bash
git clone https://github.com/nesubat/Warehouse_app.git
cd Warehouse_app 
```
### 2. Create and activate a virtual environment
```bash
python -m venv venv 
venv\Scripts\activate
```
### 3. Install dependencies
```bash
pip install flask pandas openpyxl xlwings werkzeug xlsxwriter fitz pyinstaller 
```
### 4.Run application
```bash
python app.py 
```
The app will be available in your browser at http://127.0.0.1:5000

### create .exec file
```bash
pyinstaller --onefile --name "WarehouseApp" app.py
```


