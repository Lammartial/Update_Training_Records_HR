#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HR_IMPORT4 Rev 5.1  (combined A+B)
Author: (generated)
Date: 2025-08-26 06:08:02

WHAT THIS SCRIPT DOES (high-level):
1) Import a new Training Evidence scan Excel file (File4) into File1 (Training_Evidence_Scan.xlsx).
   - Appends new rows to File1 (one file only; no _with_import.xlsx intermediate).
   - Writes the same "Import date" timestamp to column Q for each newly-added row.
   - BEFORE writing anything, checks for duplicates using ONLY the 3 fields:
       ["Training_No.", "Training_date", "Employee_ID"]
     If any potential new row already exists in File1 with the exact same values in these three fields,
     the script logs the offending combination(s), aborts, and does NOT save changes to File1.
     The log also states that the duplicate check was performed.

2) Enrich ONLY the newly imported rows (do NOT touch existing rows unless linking to them):
   - Look up "Vector_value from Training Matrix" (Column G) using File2 (Training_Matrix.xlsx)
     based on the employee's "Training Group No." from File3 (Employee master data) and the "Training_No.".
   - Rebuild hyperlinks in Column D for the new rows only (based on Filename in Column C).
   - Write an "Expiration date calculated according Vector_value" formula in Column H for the new rows.
   - Fill "Active Flag" in Column O for the new rows (from row 4 in File2).
   - Compute "Training_Compliance_Status (calculated)" in Column I for the new rows only.
   - Write **Training_Content** in Column K from **File2 row 3** (same column as Training_No on row 6).
   - Link follow-up: for any older NOK rows with the same (Employee_ID, Training_No.),
     if the newly imported row is OK, write its Evidence No. to Column P of those older NOK rows and
     set their Column I to literal 'Recertified'.

Assumptions / Notes:
- Sheet is assumed to be the active sheet in File1 (as in Script A). Script B logic expects a sheet named 'DATA';
  if present, you can rename the active sheet in File1 to 'DATA'. The script will use the active sheet.
- Column meanings (as used in previous scripts):
    A: Primary Key (integer, increases sequentially)
    B: Training Evidence No. (same for all rows created from one source-sheet in File4)
    C: Filename (used to build hyperlink in D)
    D: Hyperlink (formula set here)
    E: Training_No.
    F: Training_date  (datetime)
    G: Vector_value from Training Matrix (computed here for new rows)
    H: Expiration date calculated according Vector_value (formula set here for new rows)
    I: Training_Compliance_Status (calculated)  -> 'OK' or 'NOK' or later 'Recertified'
    K: Training_Content (from File2 row 3)
    M: Training Group No.
    O: Active Flag (copied from File2 per Training_No.)
    P: Follow up Training Evidence No. (linked here for older NOK rows)
    Q: Import date (this run's timestamp; header must be 'Import date' in Q1)
- This script writes to File1 only once, after all checks pass.

Configuration below mirrors your existing scripts; adjust paths and the source cells as needed.
"""

import os
import sys
import logging
import re
from typing import List, Optional, Tuple, Dict
from datetime import datetime, date, timedelta

import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font
from openpyxl.styles import Alignment, numbers
from openpyxl.utils import get_column_letter


# =========================
# GUI (Tkinter) – choose File4 and run import, show log
# =========================
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
except Exception as _e:
    # If Tkinter is unavailable, fallback to CLI behavior
    tk = None

class _TextHandler(logging.Handler):
    """Logging handler that writes log records into a Tkinter Text widget."""
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
        widget = self.text_widget
        def _write():
            widget.insert("end", msg + "\n")
            widget.see("end")
        try:
            widget.after(0, _write)
        except Exception:
            # If widget is gone, ignore
            pass

def launch_gui():
    if tk is None:
        print("Tkinter not available; starting CLI import...")
        return main()

    root = tk.Tk()
    root.title("HR_IMPORT Script 2 Rev 5.4 - Training Record Evidence IMPORT")

    # --- Styles ---
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure("Start.TButton", font=("Segoe UI", 10, "bold"), padding=8)

    # Predefined folder for the file picker
    predefined_dir = USER_HOME + r"\RRC power solutions\RRC VN - Documents\999_SHARE_VN\280_HR\004_Training\100_Training\Training_Evidence_Scan\XLSX_to_import"
    # predefined_dir = r"C:\Training HR scripts\Input"

    # --- Styles / layout ---
    root.geometry("960x640")
    root.columnconfigure(0, weight=1)
    root.rowconfigure(1, weight=1)

    header = ttk.Frame(root, padding=12)
    header.grid(row=0, column=0, sticky="ew")
    header.columnconfigure(1, weight=1)

    ttk.Label(header, text="Step 1: Choose the Excel File with the training evidence", font=("Segoe UI", 11, "bold")).grid(row=0, column=0, sticky="w", padx=(0,8), pady=(0,4))
    ttk.Label(header, text="Import File:").grid(row=1, column=0, sticky="w", padx=(0,8))

    file_var = tk.StringVar(value="")
    entry = ttk.Entry(header, textvariable=file_var)
    entry.grid(row=1, column=1, sticky="ew", padx=4)

    def browse():
        initialdir = predefined_dir if os.path.isdir(predefined_dir) else os.path.expanduser("~")
        path = filedialog.askopenfilename(
            title="Choose Training Evidence source (File4)",
            initialdir=initialdir,
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")],
        )
        if path:
            file_var.set(path)
            start_btn.state(["!disabled"])

    browse_btn = ttk.Button(header, text="Browse…", command=browse)
    browse_btn.grid(row=1, column=2, padx=(8,0))

    # Start import button
    ttk.Label(header, text="Step 2: Start the import", font=("Segoe UI", 11, "bold")).grid(row=2, column=0, sticky="w", padx=(0,8), pady=(10,4))
    start_btn = ttk.Button(header, text="START THE IMPORT", style="Start.TButton")
    start_btn.state(["disabled"])
    start_btn.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(4,0))

    # Separator
    ttk.Separator(root, orient="horizontal").grid(row=1, column=0, sticky="ew")

    # Log area
    main_area = ttk.Frame(root, padding=12)
    main_area.grid(row=2, column=0, sticky="nsew")
    main_area.rowconfigure(0, weight=1)
    main_area.columnconfigure(0, weight=1)

    ttk.Label(main_area, text="Progress", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w", pady=(0,6))
    progress = ttk.Progressbar(main_area, mode="indeterminate")
    progress.grid(row=1, column=0, sticky="ew", pady=(0,10))

    ttk.Label(main_area, text="Log output", font=("Segoe UI", 10, "bold")).grid(row=2, column=0, sticky="w", pady=(0,8))

    text = tk.Text(main_area, wrap="word", height=20)
    text.grid(row=3, column=0, sticky="nsew")
    sb = ttk.Scrollbar(main_area, orient="vertical", command=text.yview)
    sb.grid(row=3, column=1, sticky="ns")
    text.configure(yscrollcommand=sb.set)

    # Attach a GUI logging handler
    gui_handler = _TextHandler(text)
    gui_handler.setLevel(logging.INFO)
    gui_handler.setFormatter(logging.Formatter("%(message)s"))
    logging.getLogger().addHandler(gui_handler)

    # Define the start action
    def do_start():
        global FILE4
        path = file_var.get().strip()
        if not path:
            messagebox.showerror("Missing file", "Please choose File4 (.xlsx) before starting the import.")
            return
        if not os.path.exists(path):
            messagebox.showerror("Not found", f"File does not exist:\n{path}")
            return
        FILE4 = path
        start_btn.state(["disabled"])
        browse_btn.state(["disabled"])
        entry.state(["disabled"])
        progress.start(10)
        text.insert("end", f"Using File4: {FILE4}\n")
        text.see("end")
        try:
            main()
            messagebox.showinfo("Done", "Import finished successfully.")
        except SystemExit as e:
            messagebox.showwarning("Aborted", f"The script exited early with code {e.code}. Check log for details.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{e}")
        finally:
            progress.stop()
            start_btn.state(["!disabled"])
            browse_btn.state(["!disabled"])
            entry.state(["!disabled"])

    start_btn.configure(command=do_start)

    root.mainloop()

# =========================
# CONFIGURATION (edit as needed)
# =========================
USER_HOME = os.path.expanduser("~")

FILE1 = USER_HOME + r"\RRC power solutions\RRC VN - Documents\999_SHARE_VN\280_HR\004_Training\100_Training\Training_Evidence_Scan_Test.xlsx"
# FILE1 = r"C:\Training HR scripts\Output\Training_Evidence_Scan_Test.xlsx"
# FILE4 = r"C:\Training HR scripts\Input\Training_Evidence.xlsx"
FILE4 = USER_HOME + r"\RRC power solutions\RRC VN - Documents\999_SHARE_VN\280_HR\004_Training\100_Training\Training_Evidence_Scan\XLSX_to_import\Training_Evidence.xlsx"
FILE2 = USER_HOME + r"\RRC power solutions\RRC VN - Documents\999_SHARE_VN\280_HR\004_Training\100_Training\Training_Matrix.xlsx"
FILE3 = r"C:\Training HR scripts\RRC_Employee_Master_Data_2025-07-10.xlsx"
# FILE3 = USER_HOME + r"\RRC power solutions\RRC VN - Documents\280_HR\104_ Employee_Management_Table\02. Employee management\RRC_Employee_Master_Data_2025-07-10.xlsx"

# Logging to same folder as File1
LOG_FILE = os.path.join(os.path.dirname(FILE1), "HR_IMPORT4_Rev5_1.log")

# Destination columns (1-based index) used consistently
COL_PRIMARY_KEY = 1     # A
COL_TRAIN_EVID_NO = 2   # B
COL_FILENAME = 3        # C (for hyperlink building in D)
COL_HYPERLINK = 4       # D
COL_TRAIN_NO = 5        # E
COL_TRAIN_START = 6     # F  ("Training_date")
COL_VECTOR_VALUE = 7    # G
COL_EXPIRY = 8          # H
COL_STATUS = 9          # I
COL_TRAIN_CONTENT = 11  # K ('Training_Content')
COL_TRAIN_GROUP_NO = 13 # M ('Training Group No.')
COL_ACTIVE_FLAG = 15    # O
COL_FOLLOWUP = 16       # P
COL_IMPORT_DATE = 17    # Q

# File4 (source) cells/areas
SRC_TRAIN_NO_CELL = "C1"
SRC_TRAIN_START_CELL = "C9"
SRC_TRAIN_CONTENT_CELL = "C3"
SRC_EMPLOYEE_ID_COL = "B"   # start at B21 downward
SRC_EMPLOYEE_ID_START_ROW = 21  # start at B21 downward

# =========================
# LOGGING
# =========================
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
console = logging.StreamHandler()
console.setLevel(logging.INFO)
console.setFormatter(logging.Formatter("%(message)s"))
logging.getLogger().addHandler(console)

# =========================
# UTILITIES
# =========================
def parse_datetime(value) -> Optional[datetime]:
    """Convert File4 C9 value to datetime (Excel numeric, datetime, or string like '19.08.2025  14:00:00')."""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value
    s = str(value).strip()
    for fmt in ("%d.%m.%Y %H:%M:%S", "%d.%m.%Y  %H:%M:%S", "%d.%m.%Y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    try:
        # ISO fallback
        return datetime.fromisoformat(s)
    except Exception:
        return None

def parse_training_no_list(value) -> List[int]:
    """Parse comma-separated training numbers from C1 into a list of ints."""
    if value is None or str(value).strip() == "":
        return []
    parts = [p.strip() for p in str(value).split(",")]
    nums = []
    for p in parts:
        if p:
            try:
                nums.append(int(p))
            except ValueError:
                pass
    return nums

def get_last_numeric_in_column(ws: Worksheet, col_idx: int) -> int:
    """Return max integer in a column (ignores header). If none, return 0."""
    max_val = 0
    for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx, values_only=True):
        v = row[0]
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            try:
                iv = int(v)
                if iv > max_val:
                    max_val = iv
            except Exception:
                pass
    return max_val

def read_employee_ids(ws: Worksheet) -> List[str]:
    """Read Employee_IDs from column B starting B21 down until first empty."""
    ids = []
    row = SRC_EMPLOYEE_ID_START_ROW
    while True:
        cell = ws[f"{SRC_EMPLOYEE_ID_COL}{row}"]
        val = cell.value
        if val is None or str(val).strip() == "":
            break
        ids.append(str(val).strip())
        row += 1
    return ids

def canonical_key_tuple3(training_no, training_dt, emp_id) -> Tuple[str, str, str]:
    """
    3-field canonical key for duplicate checks:
    (Training_No., Training_date, Employee_ID)
    """
    def clean_numlike(x):
        if x is None:
            return ""
        s = str(x).strip()
        try:
            return str(int(float(s)))
        except Exception:
            return s

    if isinstance(training_dt, datetime):
        dt_key = training_dt.replace(microsecond=0).strftime("%Y-%m-%d %H:%M:%S")
    else:
        dt = parse_datetime(training_dt)
        if dt is None:
            dt_key = str(training_dt).strip()
        else:
            dt_key = dt.replace(microsecond=0).strftime("%Y-%m-%d %H:%M:%S")
    return (clean_numlike(training_no), dt_key, clean_numlike(emp_id))

def clean_decimal_str(x) -> str:
    """Coerce numeric-like values to string without decimals (e.g., '123.0' -> '123')."""
    try:
        if x is None or str(x).strip() == "":
            return ""
        return str(int(float(str(x).strip())))
    except Exception:
        return str(x).strip()


def extract_evidence_no_from_sheet_name(ws_name: str):
    """
    Extract Training Evidence No. from a sheet name like 'Training_1234.pdf'.
    Returns the part after 'Training_' and before '.pdf'.
    """
    name = os.path.basename(str(ws_name)).strip()
    m = re.match(r'(?i)^Training_(.+?)\.pdf$', name)
    if not m:
        raise ValueError(f"Cannot derive Training Evidence No. from sheet name: {ws_name!r}")

    evid = m.group(1).strip()
    # Keep digits as int when possible; otherwise preserve as text
    try:
        return int(evid)
    except ValueError:
        return evid

# =========================
# STEP 0: Load File1 and prepare structures
# =========================
def step0_load_file1():
    logging.info("=== START (Rev 5.1) ===")
    logging.info(f"File1 (base): {FILE1}")
    logging.info(f"File4 (source): {FILE4}")
    logging.info(f"File2 (matrix): {FILE2}")
    logging.info(f"File3 (employee master): {FILE3}")
    logging.info(f"Log file: {LOG_FILE}")

    if not os.path.exists(FILE1):
        logging.error(f"File1 not found: {FILE1}")
        sys.exit(1)
    if not os.path.exists(FILE4):
        logging.error(f"File4 not found: {FILE4}")
        sys.exit(1)
    if not os.path.exists(FILE2):
        logging.error(f"File2 not found: {FILE2}")
        sys.exit(1)
    if not os.path.exists(FILE3):
        logging.error(f"File3 not found: {FILE3}")
        sys.exit(1)

    wb1 = load_workbook(FILE1)
    ws1 = wb1.active  # as in Script A
    return wb1, ws1

# =========================
# STEP 1: Build existing composite-key set from File1 (E,F,L)
# =========================
def step1_build_existing_keyset(ws1: Worksheet) -> Tuple[set, int, int]:
    """
    Build a set of tuple-keys for fast duplicate check:
    key = (Training_No[E], Training_date[F], Employee_ID[L])  # 3-field key
    Return: (keyset, last_primary_key, last_training_evidence_no)
    """
    keyset = set()
    last_pk = get_last_numeric_in_column(ws1, COL_PRIMARY_KEY)
    last_evid = get_last_numeric_in_column(ws1, COL_TRAIN_EVID_NO)

    max_row = ws1.max_row
    for r in range(2, max_row + 1):
        tno = ws1.cell(row=r, column=COL_TRAIN_NO).value
        tdt = ws1.cell(row=r, column=COL_TRAIN_START).value
        eid = ws1.cell(row=r, column=12).value  # L: Employee_ID
        key = canonical_key_tuple3(tno, tdt, eid)
        keyset.add(key)

    logging.info(f"Detected last Primary Key in File1: {last_pk}")
    logging.info(f"Detected last Training Evidence No. in File1: {last_evid}")
    return keyset, last_pk, last_evid

# =========================
# STEP 2: Dry-run the import from File4 and perform duplicate check
# =========================

def step2_dry_run_and_dupe_check(ws1: Worksheet, last_pk: int, last_evid: int) -> Tuple[list, int, int]:
    """
    Read File4 across all sheets and build a list of candidate new rows.
    Each candidate row is a dict with keys:
        {'pk','evidence_no','training_no','training_dt','emp_id','import_ts'}
    We check each candidate against the existing keyset (E,F,L); if any duplicate is found
    inside a source sheet, that whole source sheet is skipped and the script continues
    with the next sheet.
    Returns (candidates, next_pk, next_evid) where next_pk / next_evid are the counters after planning.
    """
    from openpyxl import load_workbook as _lw
    wb_src = _lw(FILE4, data_only=True)
    candidates = []
    pk = last_pk
    evid = last_evid

    # Gather existing keys for check and keep extending this set as we accept rows
    # so duplicates inside the File4 batch are also caught.
    existing_keys, _, _ = step1_build_existing_keyset(ws1)

    # Import timestamp (same for all appended rows)
    import_ts = datetime.now()

    skipped_sheets = 0
    imported_sheets = 0
    skipped_sheets_list = []
    imported_sheets_list = []

    for ws_name in wb_src.sheetnames:
        ws_src = wb_src[ws_name]
        logging.info(f"-- Planning import for source sheet: {ws_name} --")

        training_nos = parse_training_no_list(ws_src[SRC_TRAIN_NO_CELL].value)
        training_dt  = parse_datetime(ws_src[SRC_TRAIN_START_CELL].value)
        training_content = ws_src[SRC_TRAIN_CONTENT_CELL].value
        employee_ids = read_employee_ids(ws_src)

        if not training_nos:
            logging.info(f"[{ws_name}] No Training_No. in {SRC_TRAIN_NO_CELL}. Skip.")
            continue
        if training_dt is None:
            logging.info(f"[{ws_name}] Training_date in {SRC_TRAIN_START_CELL} invalid. Skip.")
            continue
        if not employee_ids:
            logging.info(f"[{ws_name}] No Employee_IDs from {SRC_EMPLOYEE_ID_COL}{SRC_EMPLOYEE_ID_START_ROW}. Skip.")
            continue

        # First validate the entire source sheet against existing/batch keys.
        sheet_rows = []
        sheet_has_duplicate = False

        for tno in training_nos:
            for emp in employee_ids:
                key = canonical_key_tuple3(tno, training_dt, emp)
                if key in existing_keys:
                    logging.error("DUPLICATE DETECTED: A row with the same 3 fields already exists in File1.")
                    logging.error(f"Duplicate combination -> Training_No={key[0]}, Training_date={key[1]}, Employee_ID={key[2]}")
                    logging.error(f"Skipping source sheet '{ws_name}' and continuing with the next sheet.")
                    logging.info("LOG NOTE: Duplicate-check across the 3 fields (Training_No. + Training_date + Employee_ID) was executed before import.")
                    sheet_has_duplicate = True
                    skipped_sheets_list.append(str(extract_evidence_no_from_sheet_name(ws_name)) + ".pdf")
                    break
                sheet_rows.append((int(tno), str(emp).strip(), key))
            if sheet_has_duplicate:
                break

        if sheet_has_duplicate:
            skipped_sheets += 1
            continue

        # # Evidence No. increases once per imported sheet
        # evid += 1
        # per_sheet_evid = evid
        # logging.info(f"[{ws_name}] Assigned Training Evidence No.: {per_sheet_evid}")

        # Evidence No. comes from the source sheet name: Training_xxxx.pdf
        per_sheet_evid = extract_evidence_no_from_sheet_name(ws_name)
        logging.info(f"[{ws_name}] Derived Training Evidence No. from sheet name: {per_sheet_evid}")

        for tno, emp, key in sheet_rows:
            pk += 1
            candidates.append({
                "pk": pk,
                "evidence_no": per_sheet_evid,
                "evidence_file": str(per_sheet_evid) + ".pdf",
                "training_no": int(tno),
                "training_dt": training_dt,
                "emp_id": str(emp).strip(),
                "training_content": training_content,
                "import_ts": import_ts
            })
            existing_keys.add(key)

        imported_sheets += 1
        imported_sheets_list.append(str(per_sheet_evid) + ".pdf")

    logging.info(f"Planned new rows: {len(candidates)}")
    logging.info(f"Source sheets imported: {imported_sheets}, skipped due to duplicates: {skipped_sheets}")
    logging.info(f"Full list of source sheets skipped: {skipped_sheets_list}\n")
    logging.info(f"Full list of source sheets imported: {imported_sheets_list}")


    if candidates:
        logging.info("LOG NOTE: Duplicate-check across the 3 fields (Training_No. + Training_date + Employee_ID) was executed; duplicates were skipped sheet-by-sheet.")
    else:
        logging.info("LOG NOTE: Duplicate-check across the 3 fields (Training_No. + Training_date + Employee_ID) was executed; no rows were imported.")
    return candidates, pk, evid

# =========================
# STEP 3: Append planned rows to File1 (write Import date in Q)
# =========================
def step3_append_rows(ws1: Worksheet, candidates: list):
    """Append the planned rows to File1, writing Import date in Column Q."""
    # Ensure headers
    if ws1.cell(row=1, column=COL_IMPORT_DATE).value is None:
        ws1.cell(row=1, column=COL_IMPORT_DATE, value="Import date")
    if ws1.cell(row=1, column=COL_TRAIN_CONTENT).value is None:
        ws1.cell(row=1, column=COL_TRAIN_CONTENT, value="Training_Content")

    first_new_row = ws1.max_row + 1
    for item in candidates:
        next_row = ws1.max_row + 1
        if next_row < 2:
            next_row = 2

        ws1.cell(row=next_row, column=COL_PRIMARY_KEY, value=item["pk"])                # A
        ws1.cell(row=next_row, column=COL_TRAIN_EVID_NO, value=item["evidence_no"])     # B
        ws1.cell(row=next_row, column=COL_FILENAME, value=item["evidence_file"])        # C

        ws1.cell(row=next_row, column=COL_TRAIN_NO, value=item["training_no"])          # E
        cF = ws1.cell(row=next_row, column=COL_TRAIN_START, value=item["training_dt"])  # F
        cF.number_format = "dd.mm.yyyy hh:mm:ss"

        ws1.cell(row=next_row, column=12, value=item["emp_id"])                         # L (Employee_ID)
        ws1.cell(row=next_row, column=COL_TRAIN_CONTENT, value=item["training_content"])  # K
        cQ = ws1.cell(row=next_row, column=COL_IMPORT_DATE, value=item["import_ts"])    # Q (Import date)
        cQ.number_format = "dd.mm.yyyy  hh:mm:ss"  # double space before time, per example

    last_new_row = ws1.max_row
    logging.info(f"Appended {len(candidates)} rows to File1. New rows: {first_new_row}-{last_new_row}.")
    return first_new_row, last_new_row

# =========================
# STEP 4: Prepare lookups (Employee group map from File3, Matrix & Flags from File2)
# =========================
def step4_load_support_data() -> Tuple[Dict[str, str], pd.DataFrame, Dict[str, object], Dict[str, str]]:
    """
    Returns:
      - empid_to_group: {Employee_ID(str) -> Training Group No.(str)}
      - matrix_df: DataFrame for D:FF (row 6 header) with first col = 'Training Group No.' and
                   other cols = training numbers as strings
      - active_flag_map: {training_no(str) -> flag value from row 4 (Excel)}
    """
    # Employee master: columns A (0) and E (4), skip 4 header rows
    df3 = pd.read_excel(FILE3, sheet_name="MASTER_DATA", header=None, skiprows=4, dtype={0: str})
    df3 = df3[[0, 4]]
    df3.columns = ["Employee_ID", "Training_Group_No"]
    df3["Employee_ID"] = df3["Employee_ID"].astype(str).str.strip()
    def _clean(x):
        try:
            if pd.isnull(x) or x == "":
                return ""
            return str(int(float(x)))
        except Exception:
            return str(x)
    df3["Employee_ID"] = df3["Employee_ID"].apply(_clean)
    df3["Training_Group_No"] = df3["Training_Group_No"].apply(_clean)
    empid_to_group = dict(zip(df3["Employee_ID"], df3["Training_Group_No"]))

    # Training Matrix (D:FF, header row 6 -> header=5)
    matrix_df = pd.read_excel(FILE2, sheet_name="MATRIX", header=5, usecols="D:FF")
    matrix_df.columns = matrix_df.columns.map(str)
    matrix_df = matrix_df.dropna(how="all")

    def normalize_training_no(value) -> str:
        if value is None:
            return ""
        s = str(value).strip().replace("\u00a0", "")
        if not s:
            return ""
        s = s.replace(",", ".")
        try:
            return str(int(float(s)))
        except Exception:
            return s
    
    wb2 = load_workbook(FILE2, data_only=True, read_only=True)
    ws2 = wb2["MATRIX"]

    active_flag_map: Dict[str, object] = {}
    # Scan from column E to the last used column in the worksheet.
    for col in range(5, ws2.max_column + 1):
        training_no_key = normalize_training_no(ws2.cell(row=6, column=col).value)
        if not training_no_key:
            continue

        # MATCH(...,0) returns the first exact match, so keep the first one only.
        if training_no_key not in active_flag_map:
            active_flag_map[training_no_key] = ws2.cell(row=4, column=col).value

    wb2.close()

    return empid_to_group, matrix_df, active_flag_map

# =========================
# STEP 5: Enrich ONLY new rows (Vector value G, Hyperlink D, Expiry H, Active O, Status I, Group No. M, Content K)
# =========================
def step5_enrich_new_rows(ws1: Worksheet, first_new_row: int, last_new_row: int,
                          empid_to_group: Dict[str, str], matrix_df: pd.DataFrame,
                          active_flag_map: Dict[str, object]):
    """
    For r in [first_new_row..last_new_row], compute & fill:
      - Column G: Vector_value from Training Matrix (by Employee_ID -> Group_No + Training_No mapping)
      - Column D: Hyperlink formula if there is a file name in Column C
      - Column H: Expiration date formula (based on G and F)
      - Column O: Active Flag (by Training_No. -> flag)
      - Column I: Training_Compliance_Status (OK/NOK) for new rows only
      - Column M: Training Group No.
      - Column K: Training Content (from File2 row 3)
    """
    # Helper to clean Training_No. and Employee_ID
    def _to_clean_str(x):
        try:
            if x is None or str(x).strip() == "":
                return ""
            return str(int(float(str(x).strip())))
        except Exception:
            return str(x).strip()

    # Build matrix row lookup by Training Group No. (first col of matrix_df)
    matrix_group_col = matrix_df.columns[0]
    # Pre-index for speed
    matrix_index = { _to_clean_str(r[matrix_group_col]) : r for _, r in matrix_df.iterrows() }

    # Status calculation helper
    def compute_status(active_flag_val, vec_val, tdate_cell):
        # Parse date from cell value
        v = tdate_cell
        if isinstance(v, datetime):
            tdate = v.date()
        elif isinstance(v, date):
            tdate = v
        else:
            return "OK"  # no valid date, treat as OK (no expiry)
        # Vector -> expiry
        try:
            vec = int(float(str(vec_val)))
        except Exception:
            vec = None
        if vec == 2:
            expiry = tdate + timedelta(days=365)
        elif vec == 3:
            expiry = tdate + timedelta(days=2*365)
        else:
            expiry = date(9999, 12, 31)

        # Active flag: treat 1 as active
        try:
            af = int(float(str(active_flag_val)))
        except Exception:
            af = 0

        return "NOK" if (af == 1 and expiry < date.today()) else "OK"

    def normalize_key(v) -> str:
        """Normalize Excel-like numeric text so VALUE()/MATCH-style comparisons work."""
        if v is None:
            return ""
        s = str(v).strip().replace("\u00a0", "").replace(" ", "")
        if not s:
            return ""
        # handle decimal comma or mixed separators
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                s = s.replace(",", "")
        elif "," in s:
            s = s.replace(",", ".")
        try:
            # Match Excel VALUE() behavior for numbers like 7, 7.0, 07
            n = float(s)
            if n.is_integer():
                return str(int(n))
            return str(n)
        except Exception:
            return s 

    # Vector value lookup helper
    def build_vector_lookup(file2_path: str):
        """
        Build an exact Excel-equivalent lookup for:
        INDEX(E7:MAX_COLUMN_62, MATCH(M2, D7:D62, 0), MATCH(E2, E6:MAX_COLUMN_6, 0))
        """
        wb = load_workbook(file2_path, data_only=True, read_only=True)
        ws = wb["MATRIX"]

        # Column letters:
        # D = 4, E = 5,
        row_lookup = {}
        col_lookup = {}

        # MATCH(VALUE(M2), D7:D62, 0)
        for r in range(7, 63):
            key = normalize_key(ws.cell(r, 4).value)  # D column
            if key and key not in row_lookup:         # first exact match like MATCH(...,0)
                row_lookup[key] = r

        # MATCH(VALUE(E2), E6:MAX_COLUMN, 0)
        for c in range(5, ws.max_column + 1):  # E..JA
            key = normalize_key(ws.cell(6, c).value)
            if key and key not in col_lookup:         # first exact match like MATCH(...,0)
                col_lookup[key] = c

        return wb, ws, row_lookup, col_lookup

    # Ensure headers
    if ws1.cell(row=1, column=COL_VECTOR_VALUE).value is None:
        ws1.cell(row=1, column=COL_VECTOR_VALUE, value="Vector_value from Training Matrix")
    if ws1.cell(row=1, column=COL_EXPIRY).value is None:
        ws1.cell(row=1, column=COL_EXPIRY, value="Expiration date calculated according Vector_value")
    if ws1.cell(row=1, column=COL_STATUS).value is None:
        ws1.cell(row=1, column=COL_STATUS, value="Training_Compliance_Status (calculated)")
    if ws1.cell(row=1, column=COL_ACTIVE_FLAG).value is None:
        ws1.cell(row=1, column=COL_ACTIVE_FLAG, value="Active Flag")
    if ws1.cell(row=1, column=COL_FOLLOWUP).value is None:
        ws1.cell(row=1, column=COL_FOLLOWUP, value="Follow up Training Evidence No.")
    if ws1.cell(row=1, column=COL_TRAIN_GROUP_NO).value is None:
        ws1.cell(row=1, column=COL_TRAIN_GROUP_NO, value="Training Group No.")
    if ws1.cell(row=1, column=COL_TRAIN_CONTENT).value is None:
        ws1.cell(row=1, column=COL_TRAIN_CONTENT, value="Training_Content")

    # Build once before the loop
    matrix_wb, matrix_ws, vector_row_map, vector_col_map = build_vector_lookup(FILE2)

    # Iterate new rows
    updated = 0
    for r in range(first_new_row, last_new_row + 1):
        emp_id = _to_clean_str(ws1.cell(row=r, column=12).value)  # L
        tno    = _to_clean_str(ws1.cell(row=r, column=COL_TRAIN_NO).value)
        tdate  = ws1.cell(row=r, column=COL_TRAIN_START).value

        # Vector Value (G): lookup row by group then column = training_no
        vec_value = ""
        group_no = empid_to_group.get(emp_id, "")
        row_idx = vector_row_map.get(normalize_key(group_no))
        col_idx = vector_col_map.get(tno)

        if row_idx is not None and col_idx is not None:
            vec_value = matrix_ws.cell(row=row_idx, column=col_idx).value

        ws1.cell(row=r, column=COL_VECTOR_VALUE, value=vec_value)

        # Training Group No. -> column M
        ws1.cell(row=r, column=COL_TRAIN_GROUP_NO, value=group_no)

        # Hyperlink (D): if C has filename
        fname = ws1.cell(row=r, column=COL_FILENAME).value
        if fname:
            formula = f'=HYPERLINK("https://rrcpowersolutions.sharepoint.com/sites/RRCVN/Shared Documents/999_SHARE_VN/280_HR/004_Training/100_Training/Training_Evidence_Scan/PDF_to_import/" & C{r}, "Open " & C{r})'
            ws1.cell(row=r, column=COL_HYPERLINK, value=formula)
            ws1.cell(row=r, column=COL_HYPERLINK).font = Font(color="0000FF", underline="single")

        # Expiration (H): formula using F & G
        ws1.cell(row=r, column=COL_EXPIRY).value = (
            f'=IF(OR(G{r}=0,G{r}=1),2958465,IF(G{r}=2,F{r}+365,IF(G{r}=3,F{r}+2*365,"")))'
        )
        ws1.cell(row=r, column=COL_EXPIRY).number_format = numbers.FORMAT_DATE_XLSX14

        # Active Flag (O)
        active_flag = active_flag_map.get(tno, "")
        ws1.cell(row=r, column=COL_ACTIVE_FLAG, value=active_flag)

        # Status (I) for new row only
        status_val = compute_status(active_flag, vec_value, tdate)
        ws1.cell(row=r, column=COL_STATUS, value=status_val)

        updated += 1

    logging.info(f"Enriched {updated} newly imported rows (K,G,D,H,O,I,M).")

# =========================
# STEP 6: Link follow-ups (old NOK rows => point to newest OK row imported this run)
# =========================
def step6_link_followups(ws1: Worksheet, first_new_row: int, last_new_row: int):
    """
    For each (Employee_ID, Training_No.) where the newly imported rows contain an OK,
    find any *older* NOK rows in the sheet and link them:
      - write the OK row's Evidence No. into Column P on the NOK row
      - set Column I on the NOK row to literal 'Recertified'
    """
    # Build dict of latest OK evidence_no per (emp_id, tno) among NEW rows only
    latest_ok_by_key: Dict[Tuple[str, str], Tuple[int, date, int]] = {}
    for r in range(first_new_row, last_new_row + 1):
        emp_id = clean_decimal_str(ws1.cell(row=r, column=12).value)
        tno    = clean_decimal_str(ws1.cell(row=r, column=COL_TRAIN_NO).value)
        status = str(ws1.cell(row=r, column=COL_STATUS).value or "").strip().upper()
        evno   = ws1.cell(row=r, column=COL_TRAIN_EVID_NO).value
        tdatev = ws1.cell(row=r, column=COL_TRAIN_START).value
        # Normalize date
        if isinstance(tdatev, datetime):
            tdate_key = tdatev.date()
        elif isinstance(tdatev, date):
            tdate_key = tdatev
        else:
            tdate_key = date.min

        if status == "OK" and emp_id and tno:
            key = (emp_id, tno)
            prev = latest_ok_by_key.get(key)
            # Keep the latest by date, then by row index
            if prev is None or (tdate_key, r) > (prev[1], prev[2]):
                latest_ok_by_key[key] = (evno, tdate_key, r)

    if not latest_ok_by_key:
        logging.info("No new OK rows to link as follow-ups. Skipping STEP 6.")
        return

    # Iterate all *older* rows (1..first_new_row-1):
    updates = 0
    misses  = 0
    checked = 0
    for r in range(2, first_new_row):
        emp_id = clean_decimal_str(ws1.cell(row=r, column=12).value)
        tno    = clean_decimal_str(ws1.cell(row=r, column=COL_TRAIN_NO).value)
        status = str(ws1.cell(row=r, column=COL_STATUS).value or "").strip().upper()

        if status == "NOK":
            checked += 1
            ok_info = latest_ok_by_key.get((emp_id, tno))
            if ok_info is None:
                misses += 1
                continue
            evno_ok, _, _ = ok_info
            # Only set follow-up if empty to avoid overwriting prior link
            if not ws1.cell(row=r, column=COL_FOLLOWUP).value:
                ws1.cell(row=r, column=COL_FOLLOWUP, value=evno_ok)
            ws1.cell(row=r, column=COL_STATUS, value="Recertified")
            updates += 1

    logging.info(f"STEP 6 follow-up linking: checked {checked} old NOK rows, updated {updates}, no-match {misses}.")

# =========================
# STEP 7: Light formatting for usability (keep Q visible)
# =========================
def step7_format(ws1: Worksheet):
    """Apply light formatting: wrap & center header, auto filter, freeze top row, date formats for F & H, widths."""
    max_row = ws1.max_row
    max_col = ws1.max_column
    last_col_letter = get_column_letter(max_col)

    # Wrap + center header
    for c in range(1, max_col + 1):
        cell = ws1.cell(row=1, column=c)
        cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")

    # Filter + freeze
    ws1.auto_filter.ref = f"A1:{last_col_letter}{max_row}"
    ws1.freeze_panes = "A2"

    # Format date columns F & H (if values exist)
    for r in range(2, max_row + 1):
        if ws1[f"F{r}"].value is not None:
            ws1[f"F{r}"].number_format = numbers.FORMAT_DATE_DATETIME
        if ws1[f"H{r}"].value is not None:
            ws1[f"H{r}"].number_format = numbers.FORMAT_DATE_XLSX14

    # Column widths
    for col_letter in list("ABCDEFGHIJ"):
        ws1.column_dimensions[col_letter].width = 14
    ws1.column_dimensions["D"].width = 18
    ws1.column_dimensions["K"].width = 50

# =========================
# MAIN
# =========================
def main():
    # Load File1
    wb1, ws1 = step0_load_file1()

    # Build existing keys and counters
    existing_keys, last_pk, last_evid = step1_build_existing_keyset(ws1)

    # Plan and duplicate-check
    candidates, new_last_pk, new_last_evid = step2_dry_run_and_dupe_check(ws1, last_pk, last_evid)
    if not candidates:
        logging.info("No candidates to import. Nothing to do.")
        return

    # Append rows + remember new range
    first_new_row, last_new_row = step3_append_rows(ws1, candidates)

    # Load mappings (File2 + File3)
    empid_to_group, matrix_df, active_flag_map = step4_load_support_data()

    # Enrich new rows only
    step5_enrich_new_rows(ws1, first_new_row, last_new_row, empid_to_group, matrix_df, active_flag_map)

    # Link follow-ups from new OK rows to older NOK rows
    step6_link_followups(ws1, first_new_row, last_new_row)

    # Format (light)
    step7_format(ws1)

    # SAVE File1 (single write at the end)
    wb1.save(FILE1)
    logging.info(f"SAVED updates to File1: {FILE1}")
    logging.info("=== COMPLETED (Rev 5.1) ===")

if __name__ == "__main__":
    launch_gui()
