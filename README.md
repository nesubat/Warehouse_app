# 📦 Packaging Automation Suite

The **Warehouse Automation Suite** is a robust, local Flask-based web application designed to automate complex warehouse packaging and distribution workflows. It streamlines the process of calculating item allocations, generating iterative packing sheets, and sorting PDF shipping labels based on dynamic signature codes.

---

## ✨ Key Features & Modules

### 1. 📊 Matrix Engine (Create Packing Sheets)
*   **Automated Allocation:** Upload raw Excel distribution lists and dynamically scan for active tabs.
*   **Custom Coordinate Mapping:** Visually map start cells, job IDs, and store columns via the frontend.
*   **Signature Generation:** Mathematically calculates unique packing signatures and outputs a clean, side-by-side Matrix layout alongside the original data.

### 2. ✂️ Sub-Group Engine (Iterative Processing)
*   **Targeted Sub-divisions:** Break down parent packing groups into smaller, specific item ranges (e.g., items 1-5, 6-10).
*   **Stackable Stages:** Fully iterative workflow. Generate `Stage 1`, and then use `Stage 1` as the baseline to seamlessly generate `Stage 2` without overwriting historical data.
*   **Smart Layouts:** Automatically handles Right-to-Left column insertions in source files and Side-by-Side matrix layouts in master packing sheets using `xlwings`.

### 3. 🖨️ PDF Label Shuffler
*   **Signature Matching:** Upload raw PDF store labels and automatically sort them to perfectly match the Excel Signature Code groups.
*   **Divider Injection:** Optional toggle to insert visual divider pages between different packing groups to assist floor workers.

### 4. 🗂️ Project Dashboard & File Management
*   **Persistent Sessions:** Jobs are organized into dedicated, timestamped project folders (e.g., `ProjectName_Job-123_260803_1430`).
*   **Native Integration:** Open generated Excel or PDF files directly in their native Windows applications from the browser.
*   **Auto-Cleanup:** A built-in "Zombie Sweeper" automatically clears out empty directories and deletes projects older than 7 days to preserve disk space.

---

## 🛠️ Technology Stack

*   **Backend:** Python 3, Flask, Werkzeug
*   **Data Processing:** Pandas, OpenPyxl
*   **Excel Automation:** xlwings (Runs Excel invisibly in the background for advanced formatting)
*   **Frontend:** HTML5, CSS3, Vanilla JavaScript, Jinja2 Templating

---

## 🚀 Workflow Overview

1.  **Start a Job:** Navigate to the Home page and select **Create Packing Sheets**.
2.  **Upload & Map:** Upload your raw distribution Excel file, select the tabs to process, and define the starting coordinates.
3.  **Generate:** The Matrix Engine creates your baseline `Project Folder`, Master Packing Sheet, and JSON metadata file.
4.  **Iterate (Optional):** From the Dashboard, click **Create Sub-Group** on your project. Select your baseline metadata (e.g., Original or Stage 1) and define your item ranges to generate the next Stage.
5.  **Shuffle Labels:** Navigate to the **Labels Shuffler**, point it to your project's Excel data, upload the raw PDF labels, and download the perfectly sorted batches.

---



> ** ⚠️ Important System Requirements & Best Practices

> 1. **Microsoft Excel:** Local installation of Microsoft Excel is strictly required because `xlwings` uses Excel's native engine to render complex grid modifications.
> 2. **Local Directory:** Run this application strictly from a local directory (e.g., `C:\Warehouse_app`). **Do not run inside cloud-synced folders** (OneDrive, SharePoint, Dropbox), as cloud engines lock newly created Excel files and crash cleanup routines.
> 3. **Data Hygiene:** The application is built to automatically detect true sheet boundaries. However, keeping input files trimmed of unused rows/columns is recommended for maximum processing speed.

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


