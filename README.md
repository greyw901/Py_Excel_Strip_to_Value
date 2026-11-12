# Excel Strip to Value

A Windows tool that cleans up messy exported Excel reports — the kind most accounting/ERP systems produce for a Balance Sheet, P&L, or similar.

## The problem

Reports exported from accounting software are usually painful to work with directly:

- Cells are merged across columns, so sorting, filtering, or copying breaks.
- Amounts are stored as **text**, not numbers — they look right but can't be summed, charted, or referenced in a formula.
- The sheet is padded with rows and columns that are entirely blank, left over from the export layout.
- Column widths and alignment are inconsistent.

This tool fixes all of that in one pass, on a **copy** of the file — your original export is never touched.

## What it does

Given one or more `.xlsx` / `.xlsm` / `.xls` files, it produces a `<name>_clean` copy (same extension) for each:

- Unmerges every merged cell.
- Deletes rows and columns that are blank across the **entire** sheet (not just blank in one row — it won't break reports that use indentation depth across several columns).
- Converts text-formatted amounts (`1,234.56`, `(500.00)`, `-50`) into real numbers with proper accounting number formatting.
- Right-aligns and auto-fits amount columns; resets other columns to a sensible default width.
- Two modes:
  - **Report mode** — flattens everything to static values. Use this for a straight system export.
  - **Working mode** — preserves live formulas (`=SUM(...)` etc.) in sheets you're still building, and only converts the text-amount cells around them.

It drives Excel itself via COM automation (not a file-format library), so existing fonts, borders, and fills in your report survive untouched.

## Who this is for

Anyone who regularly receives or exports accounting/ERP reports into Excel and needs them in a state where totals actually sum, filters actually work, and the data can be pasted into another workbook without carrying merged-cell formatting problems with it.

## Requirements

- Windows, with Microsoft Excel installed (the script automates real Excel via COM).
- To run the `.py` script directly: Python 3.10+ and `pip install pywin32`.
- To use the standalone `.exe` instead: nothing — no Python required. See [Building the .exe](#building-the-exe) to build your own.

## Usage

1. Put `excel_strip_to_value.py` (or the compiled `.exe`) in the same folder as the Excel file(s) you want to clean. A file that is open in Excel can stay open — save it first, since the tool works on a copy of what is on disk.
2. Run it. It lists the files it found; pick which ones to clean (`[Enter]` for all), then choose Report mode or Working mode.
3. It processes each file and reports what changed per sheet (rows removed, columns removed, numbers converted).
4. The cleaned file(s) land in the same folder as `<original name>_clean.xlsx` (same extension as the original). Nothing is written to the original file.

Full walkthrough with sample output: [`excel_strip_to_value - Instruction.txt`](excel_strip_to_value%20-%20Instruction.txt).

## Building the .exe

```bash
python -m venv venv
venv\Scripts\pip install pyinstaller pywin32
venv\Scripts\python -m PyInstaller --onefile --console --name "Excel Strip to Value" ^
  --version-file=version_info.txt ^
  --hidden-import win32com --hidden-import win32com.client ^
  --hidden-import win32com.server --hidden-import pywintypes ^
  excel_strip_to_value.py
```

Building from a clean virtual environment (rather than a Python install with lots of unrelated packages already on it) keeps the resulting `.exe` small — PyInstaller only needs to bundle `pywin32` and the standard library.

`version_info.txt` embeds the author and version into the exe's Properties > Details tab, so it travels with the exe even when it is forwarded on its own.

## License

MIT — see [LICENSE](LICENSE).

---
Version: v1.1_20261113

Author: Yeoh Huey Sun
Email: greyw901@gmail.com
GitHub: https://github.com/greyw901
