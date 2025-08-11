import csv
from pathlib import Path
from datetime import date
from openpyxl import Workbook, load_workbook
import tkinter as tk
from tkinter import filedialog, messagebox

# ---------- helpers ----------
def detect_delimiter(path: Path, encoding="utf-8") -> str:
    with path.open("r", encoding=encoding, newline="") as f:
        sample = f.read(4096)
        f.seek(0)
        try:
            return csv.Sniffer().sniff(sample).delimiter
        except Exception:
            for d in [",", ";", "\t", "|"]:
                if d in sample:
                    return d
            return ","  # fallback

def force_number(v):
    """Try to convert to int/float; return None if not possible."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return v
    v = str(v).strip()
    if v == "":
        return None
    v = v.replace(",", ".")
    try:
        if "." in v:
            return float(v)
        else:
            return int(v)
    except ValueError:
        return None

def canonical_for_compare(v):
    """Normalize for equality check: prefer numeric if possible, else trimmed string or None."""
    num = force_number(v)
    if num is not None:
        return num
    if v is None:
        return None
    return str(v).strip()

def value_in_fee_per_bucket(v):
    """True if E belongs to {300,802,803,804,850} (numeric-aware)."""
    codes = {300, 802, 803, 804, 850}
    n = force_number(v)
    if n is not None:
        if isinstance(n, float) and n.is_integer():
            n = int(n)
        return n in codes
    s = str(v).strip() if v is not None else ""
    return s in {str(x) for x in codes}

def get_or_create_sheet(wb, name: str):
    if name in wb.sheetnames:
        ws = wb[name]
        ws.delete_rows(1, ws.max_row or 1)
    else:
        ws = wb.create_sheet(title=name)
    return ws

def copy_csv_to_xlsx_sheet(src: Path, dst: Path, sheet_name: str = "Input-Daten",
                           encoding="utf-8", delimiter=None):
    if not src.exists():
        raise FileNotFoundError(f"Source not found: {src}")

    # Load or create workbook
    if dst.exists():
        wb = load_workbook(dst)
    else:
        wb = Workbook()

    # Get or create target sheet
    ws = get_or_create_sheet(wb, sheet_name)

    # Detect delimiter
    if delimiter is None:
        delimiter = detect_delimiter(src, encoding=encoding)

    # Copy CSV content as text
    with src.open("r", encoding=encoding, errors="replace", newline="") as f:
        reader = csv.reader(f, delimiter=delimiter)
        for row in reader:
            ws.append(["" if v is None else str(v) for v in row])

    # Remove default empty sheet if present
    if "Sheet" in wb.sheetnames and wb["Sheet"] != ws:
        sh = wb["Sheet"]
        if sh.max_row == 1 and sh.max_column == 1 and sh["A1"].value in (None, ""):
            try:
                wb.remove(sh)
            except Exception:
                pass

    # 1) Numeric conversion for A,B,C,E (row 2+)
    numeric_cols = [1, 2, 3, 5]  # A=1, B=2, C=3, E=5
    for col in numeric_cols:
        for row in range(2, ws.max_row + 1):
            val = ws.cell(row=row, column=col).value
            ws.cell(row=row, column=col).value = force_number(val)

    # 2) Mapping update from KSt Mapping (ROW 2 ONLY) -> update A & E
    if "KSt Mapping" in wb.sheetnames:
        ws_map = wb["KSt Mapping"]

        # Build dict from row 2 onward: Column A -> Column B
        mapping_dict = {}
        for r in range(2, ws_map.max_row + 1):  # start row 2
            key_raw = ws_map.cell(row=r, column=1).value   # A
            val_raw = ws_map.cell(row=r, column=2).value   # B
            # Coerce keys to numeric if possible so they match numeric A/E
            key = force_number(key_raw)
            if key is None:
                if key_raw is None or str(key_raw).strip() == "":
                    continue
                key = str(key_raw).strip()
            # Coerce mapped value to numeric when possible
            val = force_number(val_raw)
            if val is None and val_raw not in (None, ""):
                val = str(val_raw).strip()
            mapping_dict[key] = val

        # Apply mapping to columns A and E (row 2+)
        for r in range(2, ws.max_row + 1):
            for col in (1, 5):  # A, E
                old_val = ws.cell(row=r, column=col).value
                key_num = force_number(old_val)
                key = key_num if key_num is not None else (str(old_val).strip() if old_val not in (None, "") else None)
                if key in mapping_dict:
                    ws.cell(row=r, column=col).value = mapping_dict[key]

    # 3) Re-coerce A & E to numeric after mapping (row 2+)
    for col in (1, 5):
        for row in range(2, ws.max_row + 1):
            val = ws.cell(row=row, column=col).value
            num = force_number(val)
            if num is not None:
                ws.cell(row=row, column=col).value = num

    # 4) Delete rows (row 2+) where A == E (numeric-aware, and only if both non-empty)
    for r in range(ws.max_row, 1, -1):  # bottom-up deletion
        a_val = ws.cell(row=r, column=1).value
        e_val = ws.cell(row=r, column=5).value
        a_can = canonical_for_compare(a_val)
        e_can = canonical_for_compare(e_val)
        if a_can is not None and e_can is not None and a_can == e_can:
            ws.delete_rows(r, 1)

    # 5) Split rows into "Fee auswertung per" vs "Fee auswertung"
    #    based on E ∈ {300,802,803,804,850}
    ws_in = ws
    max_col = ws_in.max_column

    # Ensure targets fresh and copy header
    ws_per = get_or_create_sheet(wb, "Fee auswertung per")
    ws_fee = get_or_create_sheet(wb, "Fee auswertung")

    header = [ws_in.cell(row=1, column=c).value for c in range(1, max_col + 1)]
    # Guarantee at least 6 columns so we can write column F later
    while len(header) < 6:
        header.append("")
    ws_per.append(header)
    ws_fee.append(header)

    for r in range(2, ws_in.max_row + 1):
        row_vals = [ws_in.cell(row=r, column=c).value for c in range(1, max_col + 1)]
        # Ensure at least 6 columns
        if len(row_vals) < 6:
            row_vals += [""] * (6 - len(row_vals))
        e_val = ws_in.cell(row=r, column=5).value
        if value_in_fee_per_bucket(e_val):
            ws_per.append(row_vals)
        else:
            ws_fee.append(row_vals)

    # 6) Clear col F (row 2+) then write "Transfer FEE MM/YY" in both fee sheets
    today = date.today()
    label = f"Transfer FEE {today.month:02d}/{today.year % 100:02d}"

    for ws_target in (ws_per, ws_fee):
        # Clear F
        for r in range(2, ws_target.max_row + 1):
            ws_target.cell(row=r, column=6).value = None
        # Write label
        for r in range(2, ws_target.max_row + 1):
            ws_target.cell(row=r, column=6).value = label

    wb.save(dst)

# ---------- GUI ----------
def main():
    root = tk.Tk()
    root.title("CSV → Excel (Input-Daten)")

    src_var = tk.StringVar()
    dst_var = tk.StringVar()
    status_var = tk.StringVar(value="Select files, then click Execute.")

    def browse_src():
        path = filedialog.askopenfilename(
            title="Select Source CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if path:
            src_var.set(path)
            update_execute_state()

    def browse_dst():
        path = filedialog.asksaveasfilename(
            title="Select or Create Destination Excel (.xlsx)",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if path:
            dst_var.set(path)
            update_execute_state()

    def update_execute_state():
        btn_exec.config(state="normal" if (src_var.get() and dst_var.get()) else "disabled")

    def execute():
        src = Path(src_var.get().strip())
        dst = Path(dst_var.get().strip())
        if not src or not dst:
            messagebox.showwarning("Missing", "Please select both files.")
            return
        try:
            status_var.set("Working…")
            root.update_idletasks()
            copy_csv_to_xlsx_sheet(src, dst, sheet_name="Input-Daten", encoding="utf-8", delimiter=None)
            status_var.set(
                f"Done: Input-Daten updated; mapping & cleanup applied; "
                f"split into 'Fee auswertung per' and 'Fee auswertung' with F='Transfer FEE MM/YY'."
            )
            messagebox.showinfo(
                "Success",
                "Finished:\n"
                f"- Copied CSV to: {dst}\n"
                "- Sheet: Input-Daten\n"
                "- Converted A,B,C,E to numeric (row 2+)\n"
                "- Applied KSt Mapping (row 2+)\n"
                "- Deleted rows where A == E (row 2+)\n"
                "- Split to 'Fee auswertung per' (E in {300,802,803,804,850}) and 'Fee auswertung' (others)\n"
                "- Set column F to 'Transfer FEE MM/YY' in both fee sheets"
            )
        except Exception as e:
            status_var.set("Error.")
            messagebox.showerror("Error", str(e))

    # layout
    frm = tk.Frame(root, padx=12, pady=12)
    frm.grid(row=0, column=0, sticky="nsew")

    tk.Label(frm, text="Source CSV:").grid(row=0, column=0, sticky="w")
    tk.Entry(frm, textvariable=src_var, width=60).grid(row=0, column=1, padx=6)
    tk.Button(frm, text="Browse…", command=browse_src).grid(row=0, column=2)

    tk.Label(frm, text="Destination XLSX:").grid(row=1, column=0, sticky="w", pady=(8,0))
    tk.Entry(frm, textvariable=dst_var, width=60).grid(row=1, column=1, padx=6, pady=(8,0))
    tk.Button(frm, text="Browse…", command=browse_dst).grid(row=1, column=2, pady=(8,0))

    btn_exec = tk.Button(frm, text="Execute", command=execute, state="disabled")
    btn_exec.grid(row=2, column=0, columnspan=3, pady=(12,0), sticky="ew")

    tk.Label(frm, textvariable=status_var, anchor="w").grid(row=3, column=0, columnspan=3, sticky="w", pady=(8,0))

    root.resizable(False, False)
    root.mainloop()

if __name__ == "__main__":
    main()
