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
        # clear existing content
        ws.delete_rows(1, ws.max_row or 1)
    else:
        ws = wb.create_sheet(title=sheet_name)

    # auto-detect delimiter if not provided
    if delimiter is None:
        delimiter = detect_delimiter(src, encoding=encoding)

    # append rows as text initially
    with src.open("r", encoding=encoding, errors="replace", newline="") as f:
        reader = csv.reader(f, delimiter=delimiter)
        for row in reader:
            ws.append(["" if v is None else str(v) for v in row])

    # remove the default empty sheet if we created a new workbook
    if "Sheet" in wb.sheetnames and wb["Sheet"] != ws:
        sh = wb["Sheet"]
        if sh.max_row == 1 and sh.max_column == 1 and sh["A1"].value in (None, ""):
            try:
                wb.remove(sh)
            except Exception:
                pass

    # --- Explicit numeric conversion for columns A,B,C,E from row 2 ---
    numeric_cols = [1, 2, 3, 5]  # 1-based indexes: A=1, B=2, C=3, E=5
    for col in numeric_cols:
        for row in range(2, ws.max_row + 1):
            val = ws.cell(row=row, column=col).value
            num = force_number(val)
            ws.cell(row=row, column=col).value = num

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
            status_var.set(f"Done: copied to '{dst.name}' → sheet 'Input-Daten'.")
            messagebox.showinfo("Success", f"Copied data to:\n{dst}\nSheet: Input-Daten\n"
                                           f"Columns A,B,C,E converted to numeric from row 2.")
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
