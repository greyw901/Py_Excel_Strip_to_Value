"""
Clean messy Excel financial reports: unmerge cells, remove rows/columns that
are blank across the whole sheet, convert text-formatted numbers to real
numbers with thousands-separator formatting, and resize columns to fit.

Author: Yeoh Huey Sun
Email:  greyw901@gmail.com
GitHub: https://github.com/greyw901
"""

import os
import re
import sys
import time
import shutil
import traceback
import warnings
from datetime import datetime

import win32com.client as win32

warnings.filterwarnings("ignore")

SCRIPT_DIR = os.path.dirname(os.path.abspath(
    sys.executable if getattr(sys, "frozen", False) else __file__
))
os.chdir(SCRIPT_DIR)

INPUT_FOLDER = SCRIPT_DIR
OUTPUT_FOLDER = SCRIPT_DIR
OUTPUT_SUFFIX = "_Clean"
AMOUNT_NUMBER_FORMAT = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105
XL_ALIGN_LEFT = -4131
XL_ALIGN_RIGHT = -4152


def col_letter(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def parse_amount(text: str) -> float | None:
    s = text.strip()
    if not s:
        return None
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1].strip()
    if s.startswith("-"):
        neg = True
        s = s[1:].strip()
    if not re.fullmatch(r"[\d,]*\.?\d+", s):
        return None
    try:
        val = float(s.replace(",", ""))
    except ValueError:
        return None
    return -val if neg else val


def as_grid(value, rows: int, cols: int) -> tuple:
    if rows == 1 and cols == 1:
        return ((value,),)
    if rows == 1:
        return (value,)
    if cols == 1:
        return tuple((v,) for v in value)
    return value


def find_input_files() -> list[str]:
    files = []
    for name in os.listdir(INPUT_FOLDER):
        if not name.lower().endswith((".xlsx", ".xlsm", ".xls")):
            continue
        if name.startswith("~$"):
            continue
        if name.endswith(OUTPUT_SUFFIX + ".xlsx"):
            continue
        files.append(name)
    return sorted(files)


def resolve_collision(path: str) -> str:
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    n = 1
    while True:
        candidate = f"{base}_{n}{ext}"
        if not os.path.exists(candidate):
            return candidate
        n += 1


class UserCancelled(Exception):
    pass


class BackOneStep(Exception):
    pass


def clear_screen() -> None:
    os.system("cls")


def print_banner() -> None:
    print("=" * 60)
    print("Excel Strip To Value")
    print("=" * 60)


def parse_selection(raw: str, n: int) -> list[int] | None:
    """Parses a comma/space-separated list of 1-based numbers into sorted,
    de-duplicated 0-based indices. Returns None if any token isn't a valid
    in-range number."""
    tokens = [t for t in re.split(r"[,\s]+", raw.strip()) if t]
    if not tokens:
        return None
    indices = set()
    for t in tokens:
        if not t.isdigit():
            return None
        idx = int(t)
        if idx < 1 or idx > n:
            return None
        indices.add(idx - 1)
    return sorted(indices)


def prompt_select_input(inputs: list[str]) -> list[str]:
    """Shows the file-selection screen, returns the chosen subset in
    original order. Raises UserCancelled on [0]."""
    while True:
        clear_screen()
        print_banner()
        print()
        print("Select Input:")
        for i, name in enumerate(inputs, 1):
            print(f"   [{i}] \"{name}\"")
        print()
        print("-" * 60)
        print("[Enter] Select all")
        print("[#] Select (comma-separated for multiple, e.g. 1,3)")
        print("[0] Cancel & close")
        print()
        choice = input("Input: ").strip()
        if choice == "0":
            raise UserCancelled()
        if choice == "":
            return list(inputs)
        indices = parse_selection(choice, len(inputs))
        if indices is None:
            print("\nInvalid Input: enter one or more listed numbers, or [Enter] for all.")
            input("[Enter] to continue...")
            continue
        return [inputs[i] for i in indices]


def prompt_mode() -> bool:
    """Shows the mode-selection screen, returns True for Working mode
    (preserve formulas). Raises BackOneStep on [0], UserCancelled on [00]."""
    while True:
        clear_screen()
        print_banner()
        print()
        print("Select mode:")
        print("   [1] Report mode  - flatten everything to static values (system-exported reports)")
        print("   [2] Working mode - keep live formulas (sheets you're still building/using)")
        print("   [0] Back one step")
        print("   [00] Cancel & close")
        print()
        print("-" * 60)
        choice = input("Mode: ").strip()
        if choice == "00":
            raise UserCancelled()
        if choice == "0":
            raise BackOneStep()
        if choice == "1":
            return False
        if choice == "2":
            return True
        print("\nInvalid Input: enter 1, 2, 0, or 00.")
        input("[Enter] to continue...")


def clean_worksheet(ws, preserve_formulas: bool) -> dict:
    ws.Cells.UnMerge()

    used = ws.UsedRange
    used.WrapText = False
    start_row, start_col = used.Row, used.Column
    rows, cols = used.Rows.Count, used.Columns.Count
    vals = as_grid(used.Value, rows, cols)

    formula_mask = None
    if preserve_formulas:
        formulas = as_grid(used.Formula, rows, cols)
        formula_mask = [
            [isinstance(formulas[r][c], str) and formulas[r][c].startswith("=") for c in range(cols)]
            for r in range(rows)
        ]

    col_has_data = [False] * cols
    blank_rows_local = []
    numeric_positions = {}
    already_numeric_cols = set()

    for r in range(rows):
        row_has_data = False
        for c in range(cols):
            if formula_mask and formula_mask[r][c]:
                # A formula cell is always "content" - never parsed, never
                # treated as blank, no matter what it currently displays.
                row_has_data = True
                col_has_data[c] = True
                continue
            v = vals[r][c]
            if v is None:
                continue
            if isinstance(v, str):
                if v.strip() == "":
                    continue
                row_has_data = True
                col_has_data[c] = True
                parsed = parse_amount(v)
                if parsed is not None:
                    numeric_positions[(r, c)] = parsed
            else:
                row_has_data = True
                col_has_data[c] = True
                if isinstance(v, (int, float)):
                    already_numeric_cols.add(c)
        if not row_has_data:
            blank_rows_local.append(r)

    blank_cols_local = [c for c in range(cols) if not col_has_data[c]]
    amount_cols_local = set(already_numeric_cols) | {c for (_, c) in numeric_positions}

    # Number format + numeric conversion: one Range call per amount column.
    # NumberFormat is set before Value, so the Text-formatted source cell
    # accepts the incoming float as a real number instead of coercing it back
    # to text. A column with no formula in it converts in one bulk Value
    # write; a column that mixes a formula with text-amounts falls back to
    # per-cell writes touching only the text cells, so the live formula next
    # to it is never overwritten by a batched write covering the whole range.
    for c in sorted(amount_cols_local):
        col_num = start_col + c
        rng = ws.Range(ws.Cells(start_row, col_num), ws.Cells(start_row + rows - 1, col_num))
        rng.NumberFormat = AMOUNT_NUMBER_FORMAT
        col_numeric_rows = [r for r in range(rows) if (r, c) in numeric_positions]
        if not col_numeric_rows:
            continue
        col_has_formula = formula_mask and any(formula_mask[r][c] for r in range(rows))
        if not col_has_formula:
            col_array = tuple((numeric_positions.get((r, c), vals[r][c]),) for r in range(rows))
            rng.Value = col_array
        else:
            for r in col_numeric_rows:
                ws.Cells(start_row + r, col_num).Value = numeric_positions[(r, c)]

    # Remove rows/columns that are blank across the ENTIRE sheet (not just
    # locally blank in one row), each as a single batched delete so column
    # and row alignment between sections never drifts relative to each other.
    if blank_rows_local:
        addr = ",".join(f"A{start_row + r}" for r in blank_rows_local)
        ws.Range(addr).EntireRow.Delete()
    if blank_cols_local:
        addr = ",".join(f"{col_letter(start_col + c)}1" for c in blank_cols_local)
        ws.Range(addr).EntireColumn.Delete()

    # Alignment: text left, amounts right (accounting convention) - applied
    # on the final range since unmerge/merge formatting often left things
    # centered. Column width: amount columns autofit (so numbers are never
    # clipped); every other column resets to the sheet's own default width
    # instead of stretching to fit long wording. Amount columns are detected
    # fresh (any numeric cell in the final column) so the check is immune to
    # the row/column shifts the deletes above just caused.
    final = ws.UsedRange
    final.HorizontalAlignment = XL_ALIGN_LEFT
    f_start_col, f_rows, f_cols = final.Column, final.Rows.Count, final.Columns.Count
    f_vals = as_grid(final.Value, f_rows, f_cols)
    standard_width = ws.StandardWidth

    for c in range(f_cols):
        col_num = f_start_col + c
        is_amount = any(isinstance(f_vals[r][c], (int, float)) for r in range(f_rows))
        col_range = ws.Columns(col_num)
        if is_amount:
            col_range.HorizontalAlignment = XL_ALIGN_RIGHT
            col_range.AutoFit()
        else:
            col_range.ColumnWidth = standard_width

    return {
        "sheet": ws.Name,
        "rows_removed": len(blank_rows_local),
        "cols_removed": len(blank_cols_local),
        "numbers_converted": len(numeric_positions),
    }


def process_workbook(src_path: str, preserve_formulas: bool) -> dict:
    base, ext = os.path.splitext(os.path.basename(src_path))
    out_path = resolve_collision(os.path.join(OUTPUT_FOLDER, f"{base}{OUTPUT_SUFFIX}{ext}"))
    shutil.copy2(src_path, out_path)

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    sheet_stats = []
    try:
        wb = excel.Workbooks.Open(os.path.abspath(out_path))
        excel.ScreenUpdating = False
        excel.Calculation = XL_CALC_MANUAL
        try:
            for ws in wb.Worksheets:
                sheet_stats.append(clean_worksheet(ws, preserve_formulas))
        finally:
            excel.Calculation = XL_CALC_AUTOMATIC
            excel.ScreenUpdating = True
        wb.Save()
        wb.Close(False)
    finally:
        excel.Quit()

    return {"output": os.path.basename(out_path), "sheets": sheet_stats}


def write_error_log(failures: list[dict]) -> None:
    log_path = resolve_collision(f"error log_{datetime.now():%Y%m%d}.txt")
    with open(log_path, "w", encoding="utf-8") as f:
        for item in failures:
            f.write(f"\"{item['item']}\"\n")
            f.write(f"  Step: {item['step']}\n")
            f.write(f"  Error: {item['error']}\n")
            f.write("  Traceback:\n")
            for line in item["trace"].splitlines():
                f.write(f"    {line}\n")
            f.write("\n")


def main() -> None:
    start_time = time.time()
    clear_screen()
    print_banner()

    all_inputs = find_input_files()
    if not all_inputs:
        print(f"\nInputs: 0")
        print("No Excel files found to clean.")
        print("-" * 60)
        input("[Enter] to close...")
        return

    try:
        while True:
            selected = prompt_select_input(all_inputs)
            try:
                preserve_formulas = prompt_mode()
                break
            except BackOneStep:
                continue
    except UserCancelled:
        print("\nCancelled.")
        return

    clear_screen()
    print_banner()
    print()
    print(f"Inputs: {len(selected)}")
    for i, name in enumerate(selected, 1):
        print(f"   {i}. \"{name}\"")

    outputs = []
    failures = []

    print("\nProcessing")
    for name in selected:
        print(f"Cleaning \"{name}\" ...")
        try:
            result = process_workbook(os.path.join(INPUT_FOLDER, name), preserve_formulas)
            outputs.append(result["output"])
            for s in result["sheets"]:
                print(f"   Sheet \"{s['sheet']}\" : rows removed {s['rows_removed']}, "
                      f"cols removed {s['cols_removed']}, numbers converted {s['numbers_converted']}")
        except Exception as e:
            failures.append({
                "item": name,
                "step": "cleaning workbook",
                "error": f"{type(e).__name__} - {e}",
                "trace": traceback.format_exc(),
            })

    print(f"\nOutputs: {len(outputs)}")
    for i, name in enumerate(outputs, 1):
        print(f"   {i}. \"{name}\"")

    if failures:
        print(f"Failed: {len(failures)}")
        for i, f in enumerate(failures, 1):
            print(f"   {i}. \"{f['item']}\"   :   {f['error']}")

    print(f"\nElapsed: {time.time() - start_time:.1f}s")
    print("-" * 60)
    print("[Enter] to close...")
    if failures:
        print("[0] save error log and close")
    choice = input("")
    if choice.strip() == "0" and failures:
        write_error_log(failures)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        print(traceback.format_exc())
        input("[Enter] to close...")
