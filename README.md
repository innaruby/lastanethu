import csv
from pathlib import Path
from openpyxl import Workbook, load_workbook

# --- minimal GUI selectors ---
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox

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
            return ","  # safe fallback

def csv_to_existing_xlsx(src: Path, dst: Path, sheet_name: str,
                         delimiter=None, encoding="utf-8", overwrite_sheet=True):
    if not src.exists():
        raise FileNotFoundError(src)

    # Load or create workbook
    if dst.exists():
        wb = load_workbook(dst)
    else:
        wb = Workbook()

    # Prepare sheet (create if needed, otherwise clear it)
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if overwrite_sheet:
            ws.delete_rows(1, ws.max_row or 1)
    else:
        ws = wb.create_sheet(title=sheet_name)

    # Detect delimiter if not given
    if delimiter is None:
        delimiter = detect_delimiter(src, encoding=encoding)

    # Copy rows EXACTLY as text (no type conversions)
    with src.open("r", encoding=encoding, errors="replace", newline="") as f:
        reader = csv.reader(f, delimiter=delimiter)
        for row in reader:
            # Coerce every value to string, keep empty cells as empty strings
            ws.append(["" if v is None else str(v) for v in row])

    # Save to destination workbook
    # If workbook was newly created, consider removing default sheet if empty and not our target
    if "Sheet" in wb.sheetnames and wb["Sheet"].max_row == 1 and wb["Sheet"].max_column == 1 and wb["Sheet"] != ws:
        # Empty default sheet created by Workbook()
        try:
            wb.remove(wb["Sheet"])
        except Exception:
            pass

    wb.save(dst)
    return sheet_name

def main():
    root = tk.Tk()
    root.withdraw()

    messagebox.showinfo(
        "CSV → Excel",
        "Choose the source CSV file, then choose the destination Excel file (.xlsx).\n"
        "Data will be copied as-is into the chosen sheet, overwriting its contents."
    )

    # 1) Pick CSV
    csv_path_str = filedialog.askopenfilename(
        title="Select CSV file",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    if not csv_path_str:
        messagebox.showwarning("Cancelled", "No CSV selected.")
        return
    src_csv = Path(csv_path_str)

    # 2) Pick destination XLSX (existing or new)
    xlsx_path_str = filedialog.asksaveasfilename(
        title="Select or create destination Excel file",
        defaultextension=".xlsx",
        filetypes=[("Excel Workbook", "*.xlsx")]
    )
    if not xlsx_path_str:
        messagebox.showwarning("Cancelled", "No destination Excel selected.")
        return
    dst_xlsx = Path(xlsx_path_str)

    # 3) Sheet name (default “Inputfile”)
    sheet_name = simpledialog.askstring(
        "Sheet name",
        "Enter target sheet name:",
        initialvalue="Inputfile",
        parent=root
    )
    if not sheet_name:
        messagebox.showwarning("Cancelled", "No sheet name provided.")
        return

    try:
        final_sheet = csv_to_existing_xlsx(
            src=src_csv,
            dst=dst_xlsx,
            sheet_name=sheet_name,
            delimiter=None,         # auto-detect
            encoding="utf-8",       # common default; replaces bad bytes
            overwrite_sheet=True    # clear & paste from A1
        )
        messagebox.showinfo(
            "Done",
            f"Copied CSV to '{dst_xlsx.name}' → sheet '{final_sheet}'."
        )
        print(f"Done: wrote CSV to '{dst_xlsx}' sheet '{final_sheet}'.")
    except Exception as e:
        messagebox.showerror("Error", str(e))
        raise

if __name__ == "__main__":
    main()
