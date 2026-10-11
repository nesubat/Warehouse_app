# Warehouse_app — notes for Claude Code

Local Flask app (Windows, Microsoft Excel via xlwings) that turns retail distribution spreadsheets into packing
sheets, sorted courier labels, packing labels and courier CSVs. Floor staff print what it makes, so a quiet wrong
answer costs real money: **check, report with the Excel row and column, and refuse to generate** rather than guess.

Full plain-language guide to every file: `docs/CODE_GUIDE.md`. Feature list: `README.md`.

## Run and test
- `python app.py` → http://127.0.0.1:5001 (`set WAREHOUSE_DEBUG=1` for Flask's reloader/debugger; never in production)
- `pip install -r requirements.txt -r requirements-dev.txt`, then `python -m pytest -q` (runs without Excel)
- `python tests/fixtures.py OUT_DIR` writes a sample horizontal Final file + Signature Links file to try by hand
- The hook in `.claude/settings.json` lints each edited `.py` and runs the tests; fix what it reports before moving on

## Where things are
- `matrix_engine.py` Packing Sheets (xlwings). Writes Final file (a "Code for Pack <name>" column at the start of
  each selected pack), Packing Sheet, and Signature Links (a tab of store codes per tab, `Divider Barcodes`,
  `Label Plugin`). `subgroup_engine.py` adds "Code for pack <name>: 1-5" columns.
- `pdf_engine.py` Label Shuffler (codes A, B … Z, AA order, divider sheets, audit report).
- `packing_label_generator.py` vertical files → packing labels (`parse_packing_data`, `generate_packing_labels`).
- `horizontal_distribution.py` + `horizontal_labels.py` horizontal files → the same pack groups → labels in
  packing-sheet order (audit page per pack, divider per code). Item Reference `03T1P2 Store`. Each label starts
  with a box per packing-sheet code (`code_cells`: header in the job number's place, code in the quantity's).
- `courier_import.py`, `courier_export.py` consignments and the Open360 / OpenFreight CSVs.
- `label_stitcher.py` puts courier labels on packing labels by the code in their Item Ref.
- `app.py` routes. Packing Labels, horizontal labels and Courier Import share `_distribution_page(mode)`.

## Rules that keep labels right
- Never address cells by fixed row/column: find rows and columns by header text and pattern (layouts drift).
- Every unit in a distribution grid must land on exactly one label; reconcile totals and say where they differ.
- Code order is `(len(code), code)` everywhere (Packing Sheets, Shuffler, labels). Don't invent another.
- Item References must stay unique within a job and match what Stitch Labels parses (`label_stitcher.find_code`).
- Changing barcode text, references or label order changes what's printed: add or update a test in `tests/`.
- Signature Links tabs that aren't store codes are listed in `core_math.NOT_STORE_TABS`; add new ones there.

## Security (applied; keep it that way)
- Local app on 127.0.0.1. Actions are guarded by `same_origin_only` in `app.py`; don't add GET routes that change
  things without adding them to `ACTION_GETS`. No `debug=True`, no `host="0.0.0.0"` without a login in front.
- JSON inside `<script>`: `|tojson` or `script_json()`, never `json.dumps(...)|safe`. In JS, `textContent` for data
  from files; `innerHTML` only with `esc()`.
- Uploaded files: `secure_filename`, check the extension, keep them inside `UPLOAD_FOLDER`. XML from xlsx files is
  parsed with defusedxml. SQL values always as `?` parameters.

## Style
- UI copy is plain, short and tells the user what to do next. Reuse the existing classes in `static/styles.css`
  (`.pl-csv-option`, `.blueprint-card`, `.metric-badge` …) so new screens look like the rest of the app.
- Comments explain *why*, in plain language, as in the existing code.
