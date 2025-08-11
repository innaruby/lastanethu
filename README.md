import csv
from pathlib import Path
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

def copy_csv_to_xlsx_sheet(src: Path, dst: Path, sheet_name: str = "Input-Daten",
                           encoding="utf-8", delimiter=None):
    if not src.exists():
        raise FileNotFoundError(f"Source not found: {src}")

    if dst.exists():
        wb = load_workbook(dst)
    else:
        wb = Workbook()

    # get or create target sheet
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        ws.delete_rows(1, ws.max_row or 1)
    else:
        ws = wb.create_sheet(title=sheet_name)

    # detect delimiter
    if delimiter is None:
        delimiter = detect_delimiter(src, encoding=encoding)

    # copy CSV content
    with src.open("r", encoding=encoding, errors="replace", newline="") as f:
        reader = csv.reader(f, delimiter=delimiter)
        for row in reader:
            ws.append(["" if v is None else str(v) for v in row])

    # remove default empty sheet if present
    if "Sheet" in wb.sheetnames and wb["Sheet"] != ws:
        sh = wb["Sheet"]
        if sh.max_row == 1 and sh.max_column == 1 and sh["A1"].value in (None, ""):
            try:
                wb.remove(sh)
            except Exception:
                pass

    # numeric conversion for columns A,B,C,E (row 2+)
    numeric_cols = [1, 2, 3, 5]  # 1-based indexes
    for col in numeric_cols:
        for row in range(2, ws.max_row + 1):
            val = ws.cell(row=row, column=col).value
            num = force_number(val)
            ws.cell(row=row, column=col).value = num

    # --- Mapping update from KSt Mapping (ROW 2 ONLY) ---
    if "KSt Mapping" in wb.sheetnames:
        ws_map = wb["KSt Mapping"]

        # Build dict from row 2 onward: Column A -> Column B
        mapping_dict = {}
        for r in range(2, ws_map.max_row + 1):  # start from row 2 (explicit)
            key_raw = ws_map.cell(row=r, column=1).value  # col A
            val_map = ws_map.cell(row=r, column=2).value  # col B
            # Coerce key to numeric if possible to match Input-Daten numeric columns
            key_num = force_number(key_raw)
            key = key_num if key_num is not None else (key_raw if key_raw not in (None, "") else None)
            if key is not None:
                mapping_dict[key] = val_map

        # Update A and E in Input-Daten (row 2+)
        for r in range(2, ws.max_row + 1):
            for col in [1, 5]:  # A=1, E=5
                old_val = ws.cell(row=r, column=col).value
                if old_val in mapping_dict:
                    ws.cell(row=r, column=col).value = mapping_dict[old_val]

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
            status_var.set(f"Done: copied to '{dst.name}' → 'Input-Daten'. Mapping (row 2+) applied.")
            messagebox.showinfo(
                "Success",
                "Copied data to:\n"
                f"{dst}\n\n"
                "Sheet: Input-Daten\n"
                "Converted A,B,C,E to numeric from row 2.\n"
                "Applied KSt Mapping (rows from 2 only) to columns A & E."
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
