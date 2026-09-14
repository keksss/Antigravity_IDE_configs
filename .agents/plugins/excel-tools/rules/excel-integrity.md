# Excel Spreadsheet Integrity, Formatting, and Safety Rules

When reading, modifying, or creating Microsoft Excel (`.xlsx`, `.xlsm`, `.xltx`, `.csv`, `.tsv`) files using the `excel-tools` plugin, adhere to the following safety, formula validation, and styling preservation guidelines.

---

## 1. Safety and Automatic Backup

1. **Always Backup Before In-Place Editing:**
   - Never perform destructive in-place writes to an existing spreadsheet without creating a `.bak` backup first.
   - When updating workbooks, preserve original sheets, unedited columns, and existing formulas untouched.

2. **Verify File Existence and Handle Locks:**
   - Before modifying an existing file, verify that it is not locked by an active process.
   - In Microsoft Windows, an open file in Excel generates a temporary lock file prefixed with `~$` (e.g., `~$Workbook.xlsx`). If access is locked, advise the user or handle gracefully.

---

## 2. Formula Integrity and the Zero Error Policy

1. **Zero Formula Errors:**
   - Never deliver a workbook that yields formula errors: `#VALUE!`, `#REF!`, `#DIV/0!`, `#NAME?`, `#NULL!`, `#NUM!`, `#N/A`.
   - Always run formula verification via `uv run scripts/recalc.py <file.xlsx>`. Do not ship while `recalc.py` reports `errors_found`.
   - If an error is suspected to be inherited from an original file, verify by inspecting the original with `data_only=True`.

2. **Always Use Formulas, Never Hardcode Computed Totals:**
   - Write dynamic formulas (e.g., `sheet['B10'] = '=SUM(B2:B9)'`), never write the Python-computed total as a static literal number.
   - The spreadsheet must automatically update when inputs or assumptions change.

3. **Pre-Recalculation Requirement:**
   - `openpyxl` writes formula strings into XML without calculating or caching evaluated results.
   - To external consumers (e.g., `pandas`, previewers, downstream automated parsers), formula cells read as `None` until recalculated.
   - Run `recalc.py` to trigger formula calculation via Microsoft Excel (COM) or LibreOffice Calc (headless).

---

## 3. Choosing Compatible Formulas

1. **Excel 2007 Standard Functions (Universal Compatibility):**
   - Standard functions require no prefix and evaluate reliably across all spreadsheet engines: `SUM`, `SUMIFS`, `AVERAGE`, `INDEX`, `MATCH`, `IF`, `IFERROR`, `COUNTIF`, `COUNTIFS`, `SUMPRODUCT`, `VLOOKUP`.

2. **Post-2007 Functions and `_xlfn.` Prefix in openpyxl:**
   - `openpyxl` writes raw XML strings. Excel stores certain functions introduced after 2007 with internal `_xlfn.` prefixes (which Excel's UI hides from users).
   - When writing these functions in `openpyxl`, you must include the prefix:
     - `_xlfn.TEXTJOIN(...)`
     - `_xlfn.CONCAT(...)`
     - `_xlfn.IFS(...)`
     - `_xlfn.SWITCH(...)`
     - `_xlfn.MAXIFS(...)`
     - `_xlfn.MINIFS(...)`
   - Writing these without `_xlfn.` causes `#NAME?` errors in Excel.

3. **Spilling Dynamic Array Functions:**
   - Functions like `XLOOKUP`, `XMATCH`, `SORT`, `FILTER`, `UNIQUE`, `SEQUENCE` require dynamic array spill metadata.
   - If evaluated in LibreOffice without Excel 365 spill support, only the top-left cell receives a value.
   - If writing files for general compatibility without guaranteed Excel 365, prefer `INDEX`/`MATCH` over `XLOOKUP`, and perform sorting/filtering in Python prior to populating static tables.

4. **External Workbook References:**
   - Formulas referring to external files (e.g., `='[ExternalFile.xlsx]Sheet1'!$A$1`) lose their cached values if rewritten and recalculated without the external file present on disk.
   - Preserve existing cached values or warn the user before overwriting.

---

## 4. `openpyxl` Gotchas & Safe Practices

1. **Two Loads for Reading Models:**
   - `load_workbook(file, data_only=True)` loads cached evaluation values but strips all formula definitions.
   - `load_workbook(file, data_only=False)` loads formula strings but has `None` for uncalculated values.
   - One pass cannot provide both; load twice if both formulas and values are needed.

2. **`data_only=True` is Destructive on Save:**
   - Never call `.save()` on a workbook loaded with `data_only=True`. Doing so permanently overwrites all formulas with static values.

3. **Merged Cells:**
   - Always write data and formatting to the **top-left anchor cell** of a merged range.
   - All other cells in the range are `MergedCell` proxies whose `.value` is read-only.

4. **Preserving Macros in `.xlsm`:**
   - Always pass `keep_vba=True` to `load_workbook()` when modifying `.xlsm` files to avoid discarding macro modules and VBA projects.

5. **Quoting Sheet Names with Spaces:**
   - When referencing a worksheet with spaces or special characters in a formula, always wrap the sheet name in single quotes:
     `='Assumptions & Inputs'!$B$5` (unquoted evaluates to `#VALUE!`).

6. **External Workbook Links Corruption on Save:**
   - `openpyxl` does NOT safely serialize OpenXML `/xl/externalLinks/externalLink1.xml`. Saving a workbook containing external references (`=[Other.xlsx]...`) via `openpyxl.save()` corrupts cached values and triggers Microsoft Excel's recovery dialog (`[RecoveredExternalLink1]`).
   - For workbooks containing external links, perform edits and recalculations using native Microsoft Excel COM or avoid round-tripping through `openpyxl.save()`.

---

## 5. Financial Modeling & Visual Design Standards

Unless specified otherwise by the user, follow professional spreadsheet standards:

### Typography and Layout
- **Font:** Use a clean, consistent font family throughout the entire workbook (e.g., `Arial`, `Calibri`, or `Segoe UI`).
- **Hierarchy:** Distinct sizing for Title (14–16pt bold), Section Headers (11–12pt bold), Data Rows (10–11pt regular), Totals/Summary (10–11pt bold with top border and double bottom border).
- **Column Widths:** Auto-fit column widths with adequate padding to prevent `###` overflow errors.
- **Gridlines:** Ensure gridlines remain visible (`ws.views.sheetView[0].showGridLines = True`).
- **Freeze Panes:** Freeze header rows and identifier columns on wide or tall tables for usability.

### Color Palette Conventions
- **Blue text (`#0000FF` / `0,0,255`):** Hardcoded inputs, user-editable parameters, and scenario levers.
- **Black text (`#000000` / `0,0,0`):** Formulas and calculated logic.
- **Green text (`#008000` / `0,128,0`):** Internal references to other sheets within the same workbook.
- **Red text (`#FF0000` / `255,0,0`):** External references to other workbook files.
- **Soft Yellow fill (`#FFFFCC` or `#FFF2CC`):** Key assumption cells and callout inputs intended for user entry.
- **Muted Header fills:** Subtle navy (`#1B365D`), dark slate (`#2F5597`), or charcoal with white bold text.

### Number Formatting
- **Currency:** `$#,##0` or `$#,##0.00` (declare units in header: e.g., `Revenue ($K)`).
- **Zero Values:** Render zeros cleanly as `-` (e.g., `$#,##0;($#,##0);"-"`).
- **Negative Values:** Display in parentheses: `($1,250)` instead of `-$1,250`.
- **Percentages:** Stored as decimal fractions (`0.15` for 15.0%), formatted as `0.0%`.
- **Multiples / Ratios:** `0.0x` (e.g., `6.5x`).
- **Dates / Years:** Format years as plain text (`"2025"`, not `2,025`) or standard date formats (`YYYY-MM-DD`).

### Structural Transparency
- Keep assumptions separated into dedicated, labeled input cells. Formulas must reference the input cell (e.g., `=B5*(1+$B$6)` instead of `=B5*1.05`).
- Ensure formula consistency across projection columns.
