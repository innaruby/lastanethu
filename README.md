
    
    
    
    
    
    
    
    
    
    
    
    
    
import csv
from pathlib import Path
from openpyxl import Workbook, load_workbook

# ==== EDIT THESE ====
SRC_CSV = Path(r"U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Pythondateien\ENT_20250613-007 (2).csv")
DST_XLSX = Path(r"U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Umwandlungsexcel Entnahmen05.25_ANLEITUNG.xlsx")
SHEET_NAME = "Inputfile"
FORCE_DELIMITER = None
ENCODING = "utf-8"
OVERWRITE_SHEET = True
NUMERIC_COLUMNS = ["E", "F"]  # force numeric from row 2+
# ====================

def col_letter_to_index(letter: str) -> int:
    letter = letter.upper()
    idx = 0
    for c in letter:
        idx = idx * 26 + (ord(c) - ord("A") + 1)
    return idx - 1

NUMERIC_COL_IDX = {col_letter_to_index(c) for c in NUMERIC_COLUMNS}

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

    wb = load_workbook(dst) if dst.exists() else Workbook()

    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if overwrite_sheet:
            ws.delete_rows(1, ws.max_row or 1)
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
                    if i in NUMERIC_COL_IDX:
                        num = force_number(val)
                        # If it still isn't numeric (e.g., bad value), keep original text
                        new_row.append(num if num is not None else ("" if val is None else str(val)))
                    else:
                        new_row.append(smart_convert(val))
            ws.append(new_row)

    wb.save(dst)
    return sheet_name

if __name__ == "__main__":
    final_sheet = csv_to_existing_xlsx(
        SRC_CSV, DST_XLSX, SHEET_NAME,
        delimiter=FORCE_DELIMITER, encoding=ENCODING, overwrite_sheet=OVERWRITE_SHEET
    )
    print(f"Done: wrote CSV to '{DST_XLSX}' sheet '{final_sheet}' (headers preserved, E/F numeric).")

