# Packing Labels

The Packing Labels page (`/packing-labels`) turns a distribution Excel file into a printable A4 PDF. Each pack group gets one label: a courier-label area, the pack's details, and one box per item.

Code: [packing_label_generator.py](../packing_label_generator.py) (reading, checks and PDF drawing) and the `create_packing_labels` route in [app.py](../app.py).

## Workflow

1. **Scan File**: upload an `.xlsx` or `.xls` file. The tool lists its tabs.
2. **Auto-Map & Update Previews**: choose tabs and each tab's header row. The tool reads the data and runs every check below. Errors lock generation. Warnings are shown but don't block it.
3. **Customise Cell Layout**: drag blocks to set their order inside each item box. Blocks for columns the file doesn't have show as empty slots and take no space on the label.
4. **Create Project & Generate PDF**: saves the Excel file and the PDF into a new project folder.

## Columns

Headers are matched by keywords, ignoring capitals.

| Data | Header contains | Required |
|---|---|---|
| Packing Spec | `packing spec` | **Yes**, on every row |
| Job Number | `job` | **Yes**, on every row |
| Store Name | `store` or `retailer` | No |
| Quantity | `qty` or `quantity` | No (defaults to 1) |
| Description | `desc` | No |
| Thumbnail | `thumb`, `image`, `picture` or `art` | No |
| Dimensions | `dimension`, or `width` / `height` (`w` / `h`) | No |
| Address info | `address` / `address line 1` / `street address`, `address line 2`, `suburb`, `state`, `postcode`, `country` | No |
| Material | `material` | No |
| Notes | `note` | No |
| Install | `install` | No |

## Pack groups

- **Packing Spec defines the pack group.** Rows covered by one merged Packing Spec cell form one pack. A Packing Spec cell that isn't merged is a pack on its own, even when the next row has the same text.
- Several pack groups may share the same Packing Spec text.
- The data ends at the first row with an empty Packing Spec.
- **Install** is a yes/no flag. `Y`, `Yes` and `True` (any capitals) mean the pack goes to an **installer**, not the store. Anything else, including blank, means it goes to the store. Installer packs print `INSTALLER` in red when the Install block is in the layout.
- **Address info** is every available address column joined together, in the order Address Line 1, Address Line 2, Suburb, State, Postcode, Country. Some files put the whole address in one Address column, and that works the same way.
- Addresses and store names are compared **ignoring capitals, spacing and punctuation**, so `12 Bourke Rd.` matches `12  BOURKE RD`.
- Addresses also treat these street types as equal when they follow a name: Street/St, Road/Rd, Avenue/Ave, Boulevard/Blvd, Highway/Hwy, Drive/Dr, Parade/Pde, Place/Pl, Court/Ct, Crescent/Cres, Lane/Ln, Terrace/Tce, Arcade/Arc. So `49 Church Street` matches `49 Church St`, but in `12 St Kilda Rd` the `St` (after a number) stays as Saint.
- A merged cell counts as its value on **every row it covers**.

## Checks

Every message names the Excel column letter, the header and the row numbers.

### Errors: block generation until fixed

| # | Situation |
|---|---|
| — | No Packing Spec column, or no Job Number column |
| — | A row has a Job Number but no Packing Spec (rows after the first blank Packing Spec would otherwise be skipped) |
| — | A row has no Job Number |
| 2 | Rows in one pack have different addresses |
| 3 | Some rows in a pack have an address and others are blank (with no merged cell covering them) |
| 5 | Rows in one pack have different store names, or some are blank |
| 6 | Install is Y on some rows of a pack and not others |
| 7 | An installer pack has different addresses inside it (same rule as #2: Install doesn't excuse differences within a pack) |
| 10 | The same store has store packs (Install not Y) with different addresses |

### Warnings: shown on the preview, generation allowed

| # | Situation |
|---|---|
| 8 | A merged address cell runs past the edge of a pack, so neighbouring packs share it |
| 9 | A pack has no address at all (for example Sample rows) |
| 12 | The same store has several installer packs with different addresses (could be different installers) |
| 15 | Neighbouring packs have the same Packing Spec and store but aren't merged (possibly a missed merge) |
| — | A row has no thumbnail image |
| — | More than one image sits in a row's thumbnail area |

### Always fine

| # | Situation |
|---|---|
| 1 | The address is merged across some rows of a pack and typed identically in the rest |
| 11 | A store's installer pack has a different address from its store packs |

## Thumbnails

- Images stored **inside a cell** (Excel's "Place in Cell") are matched to their exact row.
- **Floating images** are matched to the cell that holds the **centre** of the image, so an image overlapping a neighbouring row still goes to its own row.
- Only the row's own thumbnail cell and the cells one column either side are searched. A row never borrows an image from the row above or below.

## Label layout (A4 landscape)

- **Margins:** 4 mm on every side.
- **Courier label region:** 107 × 150 mm, top-left of a label's first page.
- **Item boxes:** the same size on every page (about 43 × 62 mm), with rounded corners and 2.5 mm gaps. The first page has 4 × 3 boxes beside the courier region. Later pages have 6 × 3.
- **First page of a label**, under the courier region:
  - the Packing Spec in a black panel with large white text
  - the store name and address
  - `LABEL X OF Y` and `PAGE X OF Y` in the bottom-left
- **Later pages of a label:** no Packing Spec. The store name and the two counters sit in the bottom-left.
- **Label X of Y** counts a store's pack groups in file order. When several tabs are generated together, a store's packs are counted across all of them.
- **Page X of Y** counts pages within one label.

### Inside an item box

The order of blocks follows the layout chosen on the page. Each block's style belongs to its header, wherever it's placed:

| Block | Style |
|---|---|
| Thumbnail | 35% of the box height, centred |
| Description | Regular text, wrapped |
| Dimensions | Bold text, wrapped |
| Material, Notes | Regular text, value only (no header name) |
| Install | Bold red `INSTALLER` on installer packs |
| Job Number | Bold white text on a black rounded bar |
| Quantity | Large bold number in a heavy box, no unit |

Text wraps onto extra lines rather than being cut off. If the blocks don't fit in the box, the text shrinks evenly down to 65% of its normal size. Only after that does the thumbnail shrink. Nothing is drawn outside the box.
