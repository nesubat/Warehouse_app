---
name: distribution-files
description: How this app's retail distribution spreadsheets are laid out (vertical packing lists, horizontal Packing Sheets layout, Final files with code columns, Signature Links with Label Plugin and Divider Barcodes tabs) and how packing labels, codes and Item References are derived from them. Use when reading, changing or debugging anything that parses a distribution file, makes packing labels, shuffles labels, writes courier CSVs or stitches courier labels.
---

# Distribution files in Warehouse_app

## Two layouts
**Vertical** (Packing Labels, `parse_packing_data`): one row per item. The Packing Spec cell is merged down every
row of one box; that merge defines the box. Required columns: Packing Spec, Job Number, Install.

**Horizontal** (Packing Sheets, Courier Import, horizontal labels): one row per store, one column per item.
- Job number row: a job number over each item column (`J476699-03`).
- Pack row: one merged header over each pack's item columns (or one header over a single column).
- Store rows below; quantities in the grid; a TOTAL row may follow.
- Packing Spec per store per pack: a column headed with the pack's number (`1`, `P1`), or the stock columns
  Packing Sheets was given; `x` or blank means no pack.
- Item details (Description, Size, Artwork …) in labelled rows above the header.

## What Packing Sheets adds
- **Final file**: a column at the start of each selected pack, job-row header `Code for Pack <name>`, each store's
  code (A, B …) in it. Sub-Group adds `Code for pack <name>: 1-5` columns after it. Treat these as headers whose
  cells hold codes, not quantities.
- **Signature Links**: a tab per distribution tab (store names + `Code for <name>` per pack), `Divider Barcodes`
  (installer tabs: job numbers per code) and `Label Plugin` (rows, store column, packs, stock columns per tab).

Codes come from `core_math.generate_pack_signatures` over rows `pack_group_row+1 … last_row`: stores with the
same quantities in a pack share a code; letters follow `sort_key`, and every order uses `(len(code), code)`.

## Horizontal packing labels
- Codes: the file's code column, else Signature Links, else worked out the same way. Disagreement = error.
- Order: tab → pack (left to right) → code → sub-group codes → store row. An audit page opens each pack; a
  divider (`pdf_engine.build_divider_sheet`) opens each code when dividers are on.
- Item Reference: `<store's number in the tab>T<tab>P<pack> <store>`, e.g. `03T1P2 Chadstone`. The serial is the
  code Stitch Labels matches (`label_stitcher._CODE_TP`).
- Code boxes: each label starts with one item box per code (`group['code_cells']`): the code column's header
  (`Code for Pack 2`, `Code for pack 2: 1-2`) in the job number's place, the store's code in the quantity box,
  no barcode. Drawn only: never counted as items, units, job numbers or barcode kinds.
- Audit page per pack: labels, boxes, codes, units, where the codes came from, what needs a look (no table).
- Barcode kinds: per item column (`source_key`), so `J…-03 K2` is the same piece on every store's label.

## Checks before anything prints
Quantities: blank, `-` and `0` are empty; a number with a note in brackets (`4 (1 of each kind)`) is 4 units, the note
under the quantity box; anything else (`TBC`, `-2`, `1.5`) goes in the quantity box as written; every item has a job number; a pack marked x has no quantities; every unit is on a label;
TOTAL row agrees; codes agree with quantities and with Signature Links. Report each with its Excel row and column.

## Try it
`python tests/fixtures.py out/` → upload the Final file with the Signature Links file at /packing-labels/horizontal.
