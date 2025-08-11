import csv
from pathlib import Path
from openpyxl import Workbook, load_workbook
import tkinter as tk
from tkinter import filedialog, messagebox

# ==== CONFIG (you can tweak if you like) ====
SHEET_NAME = "Inputfile"
FORCE_DELIMITER = None
ENCODING = "utf-8"
OVERWRITE_SHEET = True
NUMERIC_COLUMNS = ["E", "F"]  # force numeric from row 2+
# ===========================================

def col_letter_to_index0(letter: str) -> int:
    """A->0, B->1, ..."""
    letter = letter.upper()
    idx = 0
    for c in letter:
        idx = idx * 26 + (ord(c) - ord("A") + 1)
    return idx - 1

def col_letter_to_index1(letter: str) -> int:
    """A->1, B->2, ... (openpyxl 1-based)"""
    return col_letter_to_index0(letter) + 1

NUMERIC_COL_IDX0 = {col_letter_to_index0(c) for c in NUMERIC_COLUMNS}

def detect_delimiter(path: Path, encoding="utf-8") -> str:
    with path.open("r", encoding=encoding, newline="") as f:
        sample = f.read(4096); f.seek(0)
        try:
            return csv.Sniffer().sniff(sample).delimiter
        except Exception:
            for d in [",", ";", "\t", "|"]:
                if d in sample: return d
            return ","

def force_number(v: str):
    if v is None: return None
    v = v.strip()
    if v == "": return None
    v = v.replace(",", ".")
    try:
        if "." not in v:
            return int(v)
        return float(v)
    except ValueError:
        return None

def smart_convert(v: str):
    if v is None: return None
    v = v.strip()
    if v == "": return ""
    # int?
    if v.isdigit() or (v.startswith("-") and v[1:].isdigit()):
        try: return int(v)
        except: pass
    # float?
    try: return float(v)
    except: return v

def csv_to_existing_xlsx(src: Path, dst: Path, sheet_name: str,
                         delimiter=None, encoding="utf-8", overwrite_sheet=True):
    if not src.exists():
        raise FileNotFoundError(src)

    # Load or create workbook
    wb = load_workbook(dst) if dst.exists() else Workbook()

    # Ensure target sheet
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if overwrite_sheet:
            ws.delete_rows(1, ws.max_row or 1)
    else:
        # If the default "Sheet" exists and is empty, you can reuse it, else create
        if len(wb.sheetnames) == 1 and wb.active.max_row == 1 and wb.active.max_column == 1 and wb.active["A1"].value is None:
            wb.active.title = sheet_name
            ws = wb.active
        else:
            ws = wb.create_sheet(title=sheet_name)

    if delimiter is None:
        delimiter = detect_delimiter(src, encoding=encoding)

    with src.open("r", encoding=encoding, newline="") as f:
        reader = csv.reader(f, delimiter=delimiter)
        for row_idx, row in enumerate(reader, start=1):
            new_row = []
            for i, val in enumerate(row):
                if row_idx == 1:
                    # Header row: keep exactly as text
                    new_row.append("" if val is None else str(val))
                else:
                    # Data rows
                    if i in NUMERIC_COL_IDX0:
                        num = force_number(val)
                        new_row.append(num if num is not None else ("" if val is None else str(val)))
                    else:
                        new_row.append(smart_convert(val))
            ws.append(new_row)

    wb.save(dst)
    return dst  # return path

def post_process_bbv(dst: Path, input_sheet_name="Inputfile", bbv_sheet_name="BBV Vorlage"):
    wb = load_workbook(dst)

    # Get input sheet
    if input_sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{input_sheet_name}' not found in {dst}")

    ws_in = wb[input_sheet_name]

    # Get or create BBV Vorlage sheet
    if bbv_sheet_name in wb.sheetnames:
        ws_bbv = wb[bbv_sheet_name]
        # Optional: clear rows from row 3 downward to avoid old leftovers
        if ws_bbv.max_row >= 3:
            ws_bbv.delete_rows(3, ws_bbv.max_row - 2)
    else:
        ws_bbv = wb.create_sheet(title=bbv_sheet_name)

    # Determine data rows in Inputfile: from row 2 to last row
    last_row_in = ws_in.max_row

    # Column letter → 1-based indices for openpyxl
    IN_J = col_letter_to_index1("J")
    IN_B = col_letter_to_index1("B")
    IN_I = col_letter_to_index1("I")
    IN_C = col_letter_to_index1("C")
    IN_E = col_letter_to_index1("E")
    IN_K = col_letter_to_index1("K")

    OUT_B = col_letter_to_index1("B")
    OUT_C = col_letter_to_index1("C")
    OUT_G = col_letter_to_index1("G")
    OUT_K = col_letter_to_index1("K")
    OUT_L = col_letter_to_index1("L")
    OUT_Q = col_letter_to_index1("Q")

    # Hardcoded columns
    OUT_A = col_letter_to_index1("A")
    OUT_D = col_letter_to_index1("D")
    OUT_H = col_letter_to_index1("H")
    OUT_J = col_letter_to_index1("J")
    OUT_N = col_letter_to_index1("N")

    # Copy mapping from Inputfile row 2..N → BBV Vorlage row 3..(N+1)
    # Mapping: J->B , B->C , I->G , C->K , E->L , K->Q
    out_row = 3
    for r in range(2, last_row_in + 1):
        # Read source values (allow empty rows; they’ll just write blanks)
        v_J = ws_in.cell(row=r, column=IN_J).value
        v_B = ws_in.cell(row=r, column=IN_B).value
        v_I = ws_in.cell(row=r, column=IN_I).value
        v_C = ws_in.cell(row=r, column=IN_C).value
        v_E = ws_in.cell(row=r, column=IN_E).value
        v_K = ws_in.cell(row=r, column=IN_K).value

        ws_bbv.cell(row=out_row, column=OUT_B, value=v_J)
        ws_bbv.cell(row=out_row, column=OUT_C, value=v_B)
        ws_bbv.cell(row=out_row, column=OUT_G, value=v_I)
        ws_bbv.cell(row=out_row, column=OUT_K, value=v_C)
        ws_bbv.cell(row=out_row, column=OUT_L, value=v_E)
        ws_bbv.cell(row=out_row, column=OUT_Q, value=v_K)

        # Hardcoded values on same row
        ws_bbv.cell(row=out_row, column=OUT_A, value=2098)
        ws_bbv.cell(row=out_row, column=OUT_D, value="SA")
        ws_bbv.cell(row=out_row, column=OUT_H, value="EUR")
        ws_bbv.cell(row=out_row, column=OUT_J, value="S")
        ws_bbv.cell(row=out_row, column=OUT_N, value="V7")

        out_row += 1

    # After that: in BBV Vorlage copy value from column C to column D from row 3 onwards
    # (This will overwrite "SA" in column D as requested.)
    for r in range(3, out_row):
        ws_bbv.cell(row=r, column=OUT_D, value=ws_bbv.cell(row=r, column=OUT_C).value)

    wb.save(dst)

# ---------------- GUI ----------------

def run_with_gui():
    root = tk.Tk()
    root.title("CSV → XLSX Import + BBV Vorlage Post-Processing")

    # Variables
    csv_path_var = tk.StringVar()
    xlsx_path_var = tk.StringVar()

    def browse_csv():
        p = filedialog.askopenfilename(
            title="Select source CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if p:
            csv_path_var.set(p)

    def browse_xlsx():
        # Allow selecting an existing xlsx OR picking a new filename
        p = filedialog.askopenfilename(
            title="Select destination Excel (or Cancel and choose Save As)",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not p:
            # Offer Save As if the user canceled
            p = filedialog.asksaveasfilename(
                title="Choose destination Excel",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )
        if p:
            xlsx_path_var.set(p)

    def run_process():
        try:
            src = Path(csv_path_var.get())
            dst = Path(xlsx_path_var.get())
            if not src or not src.exists():
                messagebox.showerror("Error", "Please choose a valid source CSV file.")
                return
            if not dst:
                messagebox.showerror("Error", "Please choose a destination XLSX file.")
                return

            # Step 1: CSV → XLSX (into sheet "Inputfile")
            csv_to_existing_xlsx(
                src, dst, SHEET_NAME,
                delimiter=FORCE_DELIMITER, encoding=ENCODING, overwrite_sheet=OVERWRITE_SHEET
            )

            # Step 2: Post-processing for "BBV Vorlage"
            post_process_bbv(dst, input_sheet_name=SHEET_NAME, bbv_sheet_name="BBV Vorlage")

            messagebox.showinfo("Done", f"Completed.\n\nWrote CSV to '{dst.name}' sheet '{SHEET_NAME}' and updated 'BBV Vorlage'.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # Layout
    frm = tk.Frame(root, padx=10, pady=10)
    frm.pack(fill="both", expand=True)

    tk.Label(frm, text="Source CSV:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
    tk.Entry(frm, textvariable=csv_path_var, width=60).grid(row=0, column=1, padx=5, pady=5)
    tk.Button(frm, text="Browse…", command=browse_csv).grid(row=0, column=2, padx=5, pady=5)

    tk.Label(frm, text="Destination XLSX:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
    tk.Entry(frm, textvariable=xlsx_path_var, width=60).grid(row=1, column=1, padx=5, pady=5)
    tk.Button(frm, text="Browse…", command=browse_xlsx).grid(row=1, column=2, padx=5, pady=5)

    tk.Button(frm, text="Run", command=run_process, width=15).grid(row=2, column=1, pady=12)

    root.mainloop()

if __name__ == "__main__":
    run_with_gui()
