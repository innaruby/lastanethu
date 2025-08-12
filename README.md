import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox

import os
from pathlib import Path
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import messagebox

import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox

import csv
from pathlib import Path
import pandas as pd
import tkinter as tk
from tkinter import messagebox


import os
import sys
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import messagebox
import tempfile
import shutil

# ---- CONFIG ----
SUBJECT_KEYWORD = "Risiko- und Eigenmittelkostensätze"  # case-insensitive 'contains'
ATTACHMENT_PREFIXES = [
    "Tab_EM_ICAAP",
    "Tab_Risiko_ICAAP",
    "Tab_RisikoVerlustquote_ICAAP",
    "Tab_EigenmittelVerlustquote_ICAAP",
]
PDF_PREFIX = "Risiko- und Eigenmittelkosten"  # startswith, usually a PDF

TARGET_DIR = r"U:\rlbnas1_rlb_bw_firw_z\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien"

# -------- Red-text detection (from code 2), parameters --------
ZOOM = 2.0           # Increase to 3.0 if tiny red text is missed
RED_THRESHOLD = 160  # Min R value (0–255) to consider red
DELTA = 40           # R must exceed G and B by at least this much
MIN_FRACTION = 1e-5  # Min fraction of red pixels to flag a page
# ---------------------------------------------------------------

def Emailprocessing():
    def get_outlook_inbox():
        try:
            import win32com.client  # type: ignore
        except ImportError as e:
            raise RuntimeError(
                "pywin32 is not installed. Please run: pip install pywin32"
            ) from e

        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        return inbox

    def find_latest_matching_mail(inbox, subject_keyword):
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)
        subject_keyword_lower = subject_keyword.lower()
        for item in items:
            if getattr(item, "Class", None) != 43:  # 43 = MailItem
                continue
            subject = (item.Subject or "").strip()
            if subject_keyword_lower in subject.lower():
                return item
        return None

    def ensure_dir(path_str: str) -> Path:
        p = Path(path_str)
        p.mkdir(parents=True, exist_ok=True)
        return p

    def sanitize_filename(filename: str) -> str:
        sanitized = (
            filename
            .replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
            .replace("Ä", "Ae").replace("Ö", "Oe").replace("Ü", "Ue")
        )
        return sanitized

    def further_simplify_filename(filename: str) -> str:
        simplified = (
            filename
            .replace(" ", "_")
            .replace("-", "_")
            .replace(",", "")
            .replace("'", "")
            .replace("(", "")
            .replace(")", "")
        )
        return simplified

    def normalized_name_for_match(name: str) -> str:
        return further_simplify_filename(sanitize_filename(name)).lower()

    def name_matches_prefixes(filename: str) -> bool:
        lower = filename.lower()
        for p in ATTACHMENT_PREFIXES:
            if lower.startswith(p.lower()):
                return True
        if lower.startswith(PDF_PREFIX.lower()):
            return True
        return False

    def is_veraenderungen_pdf(filename: str) -> bool:
        norm = normalized_name_for_match(filename)
        return ("veraenderungen" in norm) and norm.endswith(".pdf")

    def save_attachment_to_temp(att) -> Path:
        original_name = att.FileName or "attachment"
        sanitized_name = sanitize_filename(original_name)
        simplified_name = further_simplify_filename(sanitized_name)
        local_temp_dir = Path(tempfile.gettempdir())
        local_temp_dir.mkdir(parents=True, exist_ok=True)
        local_temp_path = local_temp_dir / simplified_name
        print(f"[TEMP SAVE] {original_name} -> {local_temp_path}")
        att.SaveAsFile(str(local_temp_path))
        return local_temp_path

    def safe_save_attachment(att, dest_dir: Path) -> Path:
        original_name = att.FileName or "attachment"
        sanitized_name = sanitize_filename(original_name)
        simplified_name = further_simplify_filename(sanitized_name)
        local_temp_dir = Path(tempfile.gettempdir())
        local_temp_path = local_temp_dir / simplified_name
        print(f"[TEMP SAVE] {original_name} -> {local_temp_path}")
        att.SaveAsFile(str(local_temp_path))
        final_path = dest_dir / simplified_name
        print(f"[MOVE] {local_temp_path} -> {final_path}")
        shutil.move(str(local_temp_path), str(final_path))
        return final_path

    def pages_with_red(pdf_path, zoom=ZOOM, red_threshold=RED_THRESHOLD, delta=DELTA, min_fraction=MIN_FRACTION):
        import fitz  # PyMuPDF
        import numpy as np

        doc = fitz.open(pdf_path)
        red_pages = []

        for i, page in enumerate(doc):
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat, alpha=False)  # RGB
            arr = np.frombuffer(pix.samples, dtype=np.uint8)

            if pix.n != 3:
                arr = arr.reshape(pix.h, pix.w, pix.n)[:, :, :3].copy()
            else:
                arr = arr.reshape(pix.h, pix.w, 3)

            r = arr[:, :, 0].astype(np.int16)
            g = arr[:, :, 1].astype(np.int16)
            b = arr[:, :, 2].astype(np.int16)

            red_mask = (r >= red_threshold) & ((r - g) >= delta) & ((r - b) >= delta)
            fraction = red_mask.mean() if red_mask.size else 0.0

            if fraction >= min_fraction:
                red_pages.append(i + 1)  # 1-based
        doc.close()
        return red_pages

    try:
        dest = ensure_dir(TARGET_DIR)

        inbox = get_outlook_inbox()
        mail = find_latest_matching_mail(inbox, SUBJECT_KEYWORD)
        if not mail:
            message = (
                "No email found whose subject contains:\n"
                f" {SUBJECT_KEYWORD}\n\n"
                "Check the Inbox or adjust the keyword."
            )
            messagebox.showwarning("Not found", message)
            return

        attachments = getattr(mail, "Attachments", None)
        if not attachments or attachments.Count == 0:
            messagebox.showinfo("Done", "Email found, but it has no attachments.")
            return

        saved = []
        skipped = []
        veraenderungen_pdf_temp_path = None

        print(f"Total attachments found: {attachments.Count}")
        for i in range(1, attachments.Count + 1):
            att = attachments.Item(i)
            fname = att.FileName or ""
            print(f"Processing attachment: {fname}")

            if is_veraenderungen_pdf(fname):
                try:
                    veraenderungen_pdf_temp_path = save_attachment_to_temp(att)
                    saved.append(f"{Path(veraenderungen_pdf_temp_path).name} (temp only)")
                    print(f"[SPECIAL] Stored Veraenderungen PDF in temp: {veraenderungen_pdf_temp_path}")
                except Exception as e:
                    skipped.append(f"{fname} (error saving to temp: {e})")
                    print(f"Error temp-saving {fname}: {e}")
                continue

            if name_matches_prefixes(fname):
                try:
                    saved_path = safe_save_attachment(att, dest)
                    if saved_path:
                        saved.append(saved_path.name)
                        print(f"Saved attachment: {saved_path.name}")
                    else:
                        skipped.append(fname)
                except Exception as e:
                    skipped.append(f"{fname} (error: {e})")
                    print(f"Error saving attachment {fname}: {e}")
            else:
                skipped.append(fname)
                print(f"Skipped attachment: {fname}")

        detection_msg = ""
        if veraenderungen_pdf_temp_path and Path(veraenderungen_pdf_temp_path).is_file():
            try:
                print(f"[DETECTION] Scanning for red text: {veraenderungen_pdf_temp_path}")
                red_pages = pages_with_red(str(veraenderungen_pdf_temp_path))
                if red_pages:
                    detection_msg = (
                        f"\n\nRed text detected on pages: {', '.join(map(str, red_pages))}."
                    )
                else:
                    detection_msg = "\n\nNo red text detected."
            except Exception as e:
                detection_msg = f"\n\nRed detection failed: {e}"
        summary = (
            f"Email:\n Subject: {mail.Subject}\n Received: {mail.ReceivedTime}\n\n"
            f"Saved ({len(saved)}):\n - " + "\n - ".join(saved or [""])
        )
        if skipped:
            summary += f"\n\nSkipped ({len(skipped)}):\n - " + "\n - ".join(skipped)
        summary += detection_msg

        messagebox.showinfo("Completed", summary)

    except Exception as e:
        messagebox.showerror("Error", str(e))
        print(f"Error in Emailprocessing: {e}")


def RK_CCF():
    SOURCE_DIR = Path(r"U:\rlbnas1_rlb_bw_firw_z\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien")
    OUTPUT_DIR = Path(r"U:\rlbnas1_rlb_bw_firw_z\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Upload_Dateien_Vorkalk")
    OUTPUT_PATH = OUTPUT_DIR / "34000_RK_CCF.csv"

    # ---------- helpers ----------
    def read_any(path: Path) -> pd.DataFrame:
        """Read CSV or Excel. Assume no header in input (like Power Query's Column1..)."""
        if path.suffix.lower() in [".csv", ".txt"]:
            # Try to detect delimiter; default to comma, fall back to semicolon
            try:
                df = pd.read_csv(path, header=None, dtype=str, encoding="utf-8", low_memory=False)
            except UnicodeDecodeError:
                df = pd.read_csv(path, header=None, dtype=str, encoding="latin1", low_memory=False)
            # If it looks like semicolon-separated (few columns but many semicolons), try again:
            if df.shape[1] == 1 and ";" in str(df.iloc[0, 0]):
                df = pd.read_csv(path, header=None, dtype=str, sep=";", encoding=df.encoding if hasattr(df, "encoding") else "utf-8", low_memory=False)
            return df
        else:
            # Excel
            return pd.read_excel(path, header=None, dtype=str, engine="openpyxl")

    def write_same_format(path: Path, df: pd.DataFrame):
        """Overwrite original file, adding headers."""
        if path.suffix.lower() in [".csv", ".txt"]:
            # If original looked semicolon-based, preserve that; else default comma
            sep = ";"
            with open(path, "rb") as f:
                head = f.read(1024)
                if b";" not in head:
                    sep = ","
            df.to_csv(path, index=False, sep=sep, encoding="utf-8")
        else:
            df.to_excel(path, index=False, engine="openpyxl")

    def ensure_column_names(df: pd.DataFrame, prefix="Column") -> pd.DataFrame:
        n = df.shape[1]
        df.columns = [f"{prefix}{i}" for i in range(1, n + 1)]
        return df

    def to_int_safe(s):
        return pd.to_numeric(s, errors="coerce").astype("Int64")

    def to_float_from_en_gb(s):
        """Parse with dot decimal (en-GB style)."""
        if s is None:
            return pd.Series(dtype="float64")
        return pd.to_numeric(s.str.replace(",", ".", regex=False), errors="coerce")

    def replace_dot_with_comma(series: pd.Series) -> pd.Series:
        return series.astype(str).str.replace(".", ",", regex=False)

    # ---------- step A: EM_ICAAP files — insert dummy headers ----------
    def fix_em_icaap_headers():
        for p in SOURCE_DIR.iterdir():
            if p.is_file() and "EM_ICAAP" in p.name:
                df = read_any(p)
                # Insert dummy headers so data starts from row 2
                df = ensure_column_names(df, prefix="Column")
                write_same_format(p, df)

    # ---------- step B: Risiko_ICAAP processing ----------
    def process_risiko_icaap():
        # Collect rows from all matching files (if multiple exist)
        frames = []

        for p in SOURCE_DIR.iterdir():
            if not (p.is_file() and "Risiko_ICAAP" in p.name):
                continue

            df = read_any(p)
            df = ensure_column_names(df, prefix="Column")

            # Ensure we have at least 11 columns referenced in the M-steps
            for i in range(df.shape[1] + 1, 12):
                df[f"Column{i}"] = np.nan # pad missing

            # --- TransformColumnTypes ---
            # Column1,7,8 Int; Column2,9 text; Column10 number (en-GB)
            df["Column1"] = to_int_safe(df["Column1"])
            df["Column7"] = to_int_safe(df["Column7"])
            df["Column8"] = to_int_safe(df["Column8"])
            # Column10 as number with dot decimal
            col10_num = to_float_from_en_gb(df["Column10"].astype(str))
            # Keep original strings too because later we rename it to "Risikokostensatz_Variabel_(in_%)"
            df["Column10"] = col10_num

            # --- RenameColumns ---
            rename_map = {
                "Column1": "BLZ",
                "Column2": "Rating_od_wNote",
                "Column3": "Rating_Kategorie",
                "Column4": "Forderungsklasse",
                "Column5": "Risikokundengruppe",
            }
            df = df.rename(columns=rename_map)

            # --- Duplicate & rename helpers ---
            df["Copy of Rating_od_wNote"] = df["Rating_od_wNote"]
            df = df.rename(columns={"Rating_od_wNote": "Rating_od_wNote_Original"})

            # Reorder to have helpers similar to PQ step
            reorder_cols = [
                "BLZ", "Rating_od_wNote_Original", "Copy of Rating_od_wNote",
                "Rating_Kategorie", "Forderungsklasse", "Risikokundengruppe",
                "Column6", "Column7", "Column8", "Column9", "Column10", "Column11"
            ]
            # Keep any extra columns at the end
            existing = [c for c in reorder_cols if c in df.columns]
            others = [c for c in df.columns if c not in existing]
            df = df[existing + others]

            # Rename the duplicate to Hilfsspalte and drop Column6
            df = df.rename(columns={"Copy of Rating_od_wNote": "Rating_od_wNote_Hilfsspalte"})
            if "Column6" in df.columns:
                df = df.drop(columns=["Column6"])

            # Replace "." -> "," in Hilfsspalte
            df["Rating_od_wNote_Hilfsspalte"] = df["Rating_od_wNote_Hilfsspalte"].astype(str).str.replace(".", ",", regex=False)

            # Add final 'Rating_od_wNote' per condition:
            # if Rating_Kategorie in {10,11,12} OR Forderungsklasse in {1..5} -> use Hilfsspalte else Original
            def use_hilf_or_orig(row):
                rk = str(row.get("Rating_Kategorie", ""))
                fk = str(row.get("Forderungsklasse", ""))
                cond = rk in {"10", "11", "12"} or fk in {"1", "2", "3", "4", "5"}
                return row["Rating_od_wNote_Hilfsspalte"] if cond else row["Rating_od_wNote_Original"]

            df["Rating_od_wNote"] = df.apply(use_hilf_or_orig, axis=1)

            # Rename remaining columns 7..11 to their names
            rename_7_11 = {
                "Column7": "Laufzeit_Von_(in_Tagen)",
                "Column8": "Laufzeit_Bis_(in_Tagen)",
                "Column9": "Risikokostensatz_Fix_(in_%)",
                "Column10": "Risikokostensatz_Variabel_(in_%)",
                "Column11": "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)",
            }
            df = df.rename(columns=rename_7_11)

            # Replace "-2" with "" in Risikokundengruppe
            if "Risikokundengruppe" in df.columns:
                df["Risikokundengruppe"] = df["Risikokundengruppe"].astype(str).replace({"-2": ""})

            # Filter out rows where Rating_od_wNote is '-1,0'/'-1.0'/'-2,0'/'-2.0'
            bad_notes = {"-1,0", "-1.0", "-2,0", "-2.0"}
            df = df[~df["Rating_od_wNote"].astype(str).isin(bad_notes)]
            # Remove helper columns
            df = df.drop(columns=[c for c in ["Rating_od_wNote_Original", "Rating_od_wNote_Hilfsspalte"] if c in df.columns])

            # Reorder to 10 columns (if present)
            target_order = [
                "BLZ", "Rating_Kategorie", "Rating_od_wNote", "Forderungsklasse",
                "Risikokundengruppe", "Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)",
                "Risikokostensatz_Fix_(in_%)", "Risikokostensatz_Variabel_(in_%)",
                "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"
            ]
            existing = [c for c in target_order if c in df.columns]
            others = [c for c in df.columns if c not in existing]
            df = df[existing + others]

            # Remove columns (keep only those needed for CCF)
            drop_cols = [
                "Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)",
                "Risikokundengruppe", "Risikokostensatz_Fix_(in_%)",
                "Risikokostensatz_Variabel_(in_%)"
            ]
            df = df.drop(columns=[c for c in drop_cols if c in df.columns])

            # Rename to Faktor_CCF
            df = df.rename(columns={"Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)": "Faktor_CCF"})

            # Keep rows where Faktor_CCF not 0 / "0"
            # Convert to numeric for filtering (accept both comma and dot decimals)
            fc_num = pd.to_numeric(df["Faktor_CCF"].astype(str).str.replace(",", ".", regex=False), errors="coerce")
            df = df[~(fc_num.fillna(0) == 0)]

            # Drop duplicates
            df = df.drop_duplicates()

            # Keep only BLZ == 34000 and Rating_od_wNote not '-1'/'-2'
            df = df[(df["BLZ"].astype("Int64") == 34000)]
            df = df[~df["Rating_od_wNote"].astype(str).isin({"-1", "-2"})]

            # Replace "." -> "," in Faktor_CCF for output formatting, then also keep numeric for possible re-use
            df["Faktor_CCF"] = df["Faktor_CCF"].astype(str).str.replace(".", ",", regex=False)

            # Add Gueltig_Ab empty string
            df["Gueltig_Ab"] = ""

            # Final reorder
            final_order = ["BLZ", "Gueltig_Ab", "Rating_Kategorie", "Rating_od_wNote", "Forderungsklasse", "Faktor_CCF"]
            existing = [c for c in final_order if c in df.columns]
            others = [c for c in df.columns if c not in existing]
            df = df[existing + others]

            frames.append(df[existing]) # keep just the required output cols

        if not frames:
            raise FileNotFoundError(f"No files containing 'Risiko_ICAAP' were found in {SOURCE_DIR}")

        # If multiple inputs, concatenate and re-deduplicate before save
        out = pd.concat(frames, ignore_index=True).drop_duplicates()

        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        # Save as CSV with semicolon (common in DE) and no index
        out.to_csv(OUTPUT_PATH, index=False, sep=";", encoding="utf-8")

    fix_em_icaap_headers()
    process_risiko_icaap()
    print(f"Done. Wrote: {OUTPUT_PATH}")




def EK_CCF():
    SOURCE_FOLDER = r"U:\rlbnas1_rlb_bw_firw_z\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien"
    TARGET_FOLDER = r"U:\rlbnas1_rlb_bw_firw_z\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Upload_Dateien_Vorkalk"
    OUTPUT_NAME = "34000_EK_CCF.csv"
    KEYWORD = "EM_ICAAP"

    def find_latest_matching_file(folder: str, keyword: str) -> Path:
        p = Path(folder)
        candidates = [f for f in p.iterdir() if f.is_file() and keyword in f.name]
        if not candidates:
            raise FileNotFoundError(f"No file found in {folder} containing '{keyword}'")
        candidates.sort(key=lambda f: f.stat().st_mtime, reverse=True)
        return candidates[0]

    def sniff_delimiter(path: Path, sample_bytes: int = 65536):
        with open(path, "rb") as fh:
            raw = fh.read(sample_bytes)
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                text = raw.decode(enc)
                dialect = csv.Sniffer().sniff(text, delimiters=[",",";","\t","|"])
                return dialect.delimiter, enc
            except Exception:
                continue
        return ",", "utf-8-sig"

    def load_table_no_headers(path: Path) -> pd.DataFrame:
        suffix = path.suffix.lower()
        if suffix in [".xlsx", ".xls"]:
            df = pd.read_excel(path, header=None, dtype=str)
        else:
            delim, enc = sniff_delimiter(path)
            df = pd.read_csv(path, delimiter=delim, encoding=enc, header=None, dtype=str)
        df.columns = [f"Column{i}" for i in range(1, len(df.columns) + 1)]
        return df

    def to_int(series: pd.Series) -> pd.Series:
        return pd.to_numeric(series, errors="coerce").astype("Int64")

    def to_float_from_eu(series: pd.Series) -> pd.Series:
        if series.dtype.kind in "biufc":
            return series.astype(float)
        s = series.astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
        return pd.to_numeric(s, errors="coerce")

    src_path = find_latest_matching_file(SOURCE_FOLDER, KEYWORD)
    df = load_table_no_headers(src_path)
    for col in ["Column1", "Column7", "Column8"]:
        if col in df.columns:
            df[col] = to_int(df[col])
    for col in ["Column2", "Column9"]:
        if col in df.columns:
            df[col] = df[col].astype(str)
    if "Column10" in df.columns:
        df["Column10"] = pd.to_numeric(df["Column10"].astype(str).str.replace(",", "", regex=False), errors="coerce")
    df = df.rename(columns={
        "Column1": "BLZ",
        "Column2": "Rating_od_wNote",
        "Column3": "Rating_Kategorie",
        "Column4": "Forderungsklasse",
        "Column5": "Risikokundengruppe",
    })
    if "Rating_od_wNote" in df.columns:
        df["Copy of Rating_od_wNote"] = df["Rating_od_wNote"]
        df = df.rename(columns={"Rating_od_wNote": "Rating_od_wNote_Original"})
    order_try = ["BLZ","Rating_od_wNote_Original","Copy of Rating_od_wNote","Rating_Kategorie","Forderungsklasse","Risikokundengruppe","Column6","Column7","Column8","Column9","Column10","Column11"]
    df = df[[c for c in order_try if c in df.columns] + [c for c in df.columns if c not in order_try]]
    if "Copy of Rating_od_wNote" in df.columns:
        df = df.rename(columns={"Copy of Rating_od_wNote": "Rating_od_wNote_Hilfsspalte"})
    if "Column6" in df.columns:
        df = df.drop(columns=["Column6"])
    if "Rating_od_wNote_Hilfsspalte" in df.columns:
        df["Rating_od_wNote_Hilfsspalte"] = df["Rating_od_wNote_Hilfsspalte"].astype(str).str.replace(".", ",", regex=False)
    def compute_rating(row):
        rk = str(row.get("Rating_Kategorie", "")).strip()
        fk = str(row.get("Forderungsklasse", "")).strip()
        cond = rk in {"10","11","12"} or fk in {"1","2","3","4","5"}
        return row.get("Rating_od_wNote_Hilfsspalte") if cond else row.get("Rating_od_wNote_Original")
    df["Rating_od_wNote"] = df.apply(compute_rating, axis=1)
    df = df.rename(columns={
        "Column7": "Laufzeit_Von_(in_Tagen)",
        "Column8": "Laufzeit_Bis_(in_Tagen)",
        "Column9": "Risikokostensatz_Fix_(in_%)",
        "Column10": "Risikokostensatz_Variabel_(in_%)",
        "Column11": "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)",
    })
    if "Risikokundengruppe" in df.columns:
        df["Risikokundengruppe"] = df["Risikokundengruppe"].astype(str).str.replace("-2", "", regex=False)
    bad_vals = {"-1,0","-1.0","-2,0","-2.0"}
    if "Rating_od_wNote" in df.columns:
        df = df[~df["Rating_od_wNote"].astype(str).isin(bad_vals)]
    for c in ["Rating_od_wNote_Original","Rating_od_wNote_Hilfsspalte"]:
        if c in df.columns:
            df = df.drop(columns=[c])
    if "Risikokundengruppe" in df.columns:
        df = df.drop(columns=["Risikokundengruppe"])
    df = df.rename(columns={
        "Risikokostensatz_Fix_(in_%)": "Eigenkapitalkosten_Fix_(in_%)",
        "Risikokostensatz_Variabel_(in_%)": "Eigenkapitalkosten_Variabel_(in_%)",
        "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)": "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)",
    })
    for c in ["Laufzeit_Von_(in_Tagen)","Laufzeit_Bis_(in_Tagen)","Eigenkapitalkosten_Fix_(in_%)","Eigenkapitalkosten_Variabel_(in_%)"]:
        if c in df.columns:
            df = df.drop(columns=[c])
    df = df.rename(columns={"Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)": "Faktor_CCF"})
    if "Faktor_CCF" in df.columns:
        fccf_num = to_float_from_eu(df["Faktor_CCF"])
        df = df[fccf_num.fillna(0) != 0]
        df["__Faktor_CCF_num__"] = fccf_num
    df = df.drop_duplicates()
    if "BLZ" in df.columns:
        df["BLZ"] = to_int(df["BLZ"])
        df = df[df["BLZ"] == 34000]
    if "Rating_od_wNote" in df.columns:
        df = df[~df["Rating_od_wNote"].astype(str).isin({"-1","-2"})]
    if "__Faktor_CCF_num__" in df.columns:
        df["Faktor_CCF"] = df["__Faktor_CCF_num__"].map(lambda x: "" if pd.isna(x) else str(x).replace(".", ","))
        df = df.drop(columns=["__Faktor_CCF_num__"])
    elif "Faktor_CCF" in df.columns:
        df["Faktor_CCF"] = df["Faktor_CCF"].astype(str).str.replace(".", ",", regex=False)
    if "Rating_od_wNote" in df.columns:
        df["Rating_od_wNote"] = df["Rating_od_wNote"].astype(str)
    df["Gueltig_Ab"] = ""
    final_order = ["BLZ", "Gueltig_Ab", "Rating_Kategorie", "Rating_od_wNote", "Forderungsklasse", "Faktor_CCF"]
    df = df[[c for c in final_order if c in df.columns]]

    Path(TARGET_FOLDER).mkdir(parents=True, exist_ok=True)
    out_path = Path(TARGET_FOLDER) / OUTPUT_NAME
    df.to_csv(out_path, index=False, sep=";", encoding="utf-8-sig")
    print(f"Saved: {out_path}")
    print(f"Columns ({len(df.columns)}): {list(df.columns)}")


import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import messagebox

# === FILE PATHS ===
FILE1_PATH = Path(r"U:\rlbnas1_rlb_bw_firw_z\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien\Tab_EM_ICAAP.csv")  # no header
FILE2_PATH = Path(r"U:\rlbnas1_rlb_bw_firw_z\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien\Primär-bzw Raiffeisenbanken\EK_Basis_Primaerbanken.csv")  # has header
OUTPUT_CSV = Path(r"U:\rlbnas1_rlb_bw_firw_z\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Upload_Dateien_Vorkalk\34000_EK_Basis.csv")
#OUTPUT_XLSX = OUTPUT_CSV.with_suffix(".xlsx")

DELIMS = [",", ";", "\t", "|"]

def EKBasis():
    def read_csv_flexible(path: Path, has_header: bool, prefer_utf8: bool = True) -> pd.DataFrame:
        encs = ["utf-8-sig", "latin-1"] if prefer_utf8 else ["latin-1", "utf-8-sig"]
        header = 0 if has_header else None
        last_err = None

        for enc in encs:
            # Auto-detect separator
            try:
                df = pd.read_csv(
                    path,
                    header=header,
                    sep=None,
                    engine="python",
                    encoding=enc,
                    on_bad_lines="skip",
                    skipinitialspace=True,
                    na_filter=False,
                )
                return df
            except Exception as e:
                last_err = e
            # Try fixed delimiters
            for sep in DELIMS:
                try:
                    df = pd.read_csv(
                        path,
                        header=header,
                        sep=sep,
                        engine="python",
                        encoding=enc,
                        on_bad_lines="skip",
                        skipinitialspace=True,
                        na_filter=False,
                    )
                    return df
                except Exception as e:
                    last_err = e
        raise RuntimeError(f"Could not read {path}: {last_err}")

    def to_int64(s: pd.Series) -> pd.Series:
        return pd.to_numeric(s, errors="coerce").astype("Int64")

    def to_float_from_maybe_comma(s: pd.Series) -> pd.Series:
        return pd.to_numeric(s.astype(str).str.replace(",", ".", regex=False), errors="coerce")

    def read_tab_em_icaap(file1: Path) -> pd.DataFrame:
        df = read_csv_flexible(file1, has_header=False)
        # Assign dummy column names
        df.columns = [f"Column{i+1}" for i in range(df.shape[1])]

        # Transform column types
        for col in ["Column1", "Column7", "Column8"]:
            if col in df: df[col] = to_int64(df[col])
        for col in ["Column2", "Column9"]:
            if col in df: df[col] = df[col].astype(str)
        if "Column11" in df: df["Column11"] = pd.to_numeric(df["Column11"], errors="coerce")
        if "Column9" in df: df["Column9"] = pd.to_numeric(df["Column9"], errors="coerce")

        # Rename columns
        df = df.rename(columns={
            "Column1": "BLZ",
            "Column2": "Rating_od_wNote",
            "Column3": "Rating_Kategorie",
            "Column4": "Forderungsklasse",
            "Column5": "Risikokundengruppe"
        })

        # Duplicate column and rename
        df["Copy of Rating_od_wNote"] = df["Rating_od_wNote"]
        df = df.rename(columns={"Rating_od_wNote": "Rating_od_wNote_Original"})

        reorder_cols = [
            "BLZ","Rating_od_wNote_Original","Copy of Rating_od_wNote","Rating_Kategorie","Forderungsklasse","Risikokundengruppe",
            "Column6","Column7","Column8","Column9","Column10","Column11"
        ]
        df = df[[c for c in reorder_cols if c in df.columns] + [c for c in df.columns if c not in reorder_cols]]
        df = df.rename(columns={"Copy of Rating_od_wNote": "Rating_od_wNote_Hilfsspalte"})

        if "Column6" in df: df = df.drop(columns=["Column6"])
        if "Rating_od_wNote_Hilfsspalte" in df:
            df["Rating_od_wNote_Hilfsspalte"] = df["Rating_od_wNote_Hilfsspalte"].astype(str).str.replace(".", ",", regex=False)

        # Compute Rating_od_wNote
        def compute_rating(row):
            rk = str(row.get("Rating_Kategorie", ""))
            fk = str(row.get("Forderungsklasse", ""))
            if rk in {"10","11","12"} or fk in {"1","2","3","4","5"}:
                return row.get("Rating_od_wNote_Hilfsspalte", "")
            return row.get("Rating_od_wNote_Original", "")
        df["Rating_od_wNote"] = df.apply(compute_rating, axis=1)

        # Reorder and rename remaining columns
        reorder2 = [
            "BLZ","Rating_od_wNote","Rating_od_wNote_Original","Rating_od_wNote_Hilfsspalte","Rating_Kategorie","Forderungsklasse",
            "Risikokundengruppe","Column7","Column8","Column9","Column10","Column11"
        ]
        df = df[[c for c in reorder2 if c in df.columns] + [c for c in df.columns if c not in reorder2]]
        df = df.rename(columns={
            "Column7": "Laufzeit_Von_(in_Tagen)",
            "Column8": "Laufzeit_Bis_(in_Tagen)",
            "Column9": "Risikokostensatz_Fix_(in_%)",
            "Column10": "Risikokostensatz_Variabel_(in_%)",
            "Column11": "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"
        })

        # Replace values & filter
        if "Risikokundengruppe" in df:
            df["Risikokundengruppe"] = df["Risikokundengruppe"].astype(str).str.replace("-2", "", regex=False)
        df = df[~df["Rating_od_wNote"].astype(str).isin({"-1,0","-1.0","-2,0","-2.0"})]

        # Drop unnecessary
        df = df.drop(columns=["Rating_od_wNote_Original","Rating_od_wNote_Hilfsspalte"], errors="ignore")
        df = df.drop(columns=["Risikokundengruppe"], errors="ignore")

        # Rename cost columns
        df = df.rename(columns={
            "Risikokostensatz_Fix_(in_%)": "Eigenkapitalkosten_Fix_(in_%)",
            "Risikokostensatz_Variabel_(in_%)": "Eigenkapitalkosten_Variabel_(in_%)",
            "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)": "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"
        })

        # Ensure types
        for col in ["BLZ","Laufzeit_Von_(in_Tagen)","Laufzeit_Bis_(in_Tagen)"]:
            if col in df: df[col] = to_int64(df[col])
        for col in ["Eigenkapitalkosten_Fix_(in_%)","Eigenkapitalkosten_Variabel_(in_%)","Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"]:
            if col in df: df[col] = to_float_from_maybe_comma(df[col])

        return df

    def read_ek_basis_primaerbanken(file2: Path) -> pd.DataFrame:
        df = read_csv_flexible(file2, has_header=True)
        df = df[df["BLZ"].astype(str) == "34"]

        for col in ["BLZ", "Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)"]:
            df[col] = to_int64(df[col])
        for col in ["Eigenkapitalkosten_Fix_(in_%)", "Eigenkapitalkosten_Variabel_(in_%)", "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"]:
            df[col] = to_float_from_maybe_comma(df[col])

        return df

    def combine_and_finalize(df1: pd.DataFrame, df2: pd.DataFrame) -> pd.DataFrame:
        df = pd.concat([df1, df2], ignore_index=True)
        df = df[df["BLZ"] != 55000]

        if "Laufzeit_Bis_(in_Tagen)" in df:
            df["Laufzeit_Bis_(in_Tagen)"] = df["Laufzeit_Bis_(in_Tagen)"].replace(365, 366)
        if "Laufzeit_Von_(in_Tagen)" in df:
            df["Laufzeit_Von_(in_Tagen)"] = df["Laufzeit_Von_(in_Tagen)"].replace(366, 367)

        if "Rating_Kategorie" in df:
            df = df[df["Rating_Kategorie"].astype(str) != "9"]

        # Duplicate & adjust cost columns
        if "Eigenkapitalkosten_Variabel_(in_%)" in df:
            df["Eigenkapitalkosten_Variabel_(in_%) - Kopie"] = df["Eigenkapitalkosten_Variabel_(in_%)"]
        df = df.drop(columns=["Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"], errors="ignore")
        if "Eigenkapitalkosten_Variabel_(in_%) - Kopie" in df:
            df = df.rename(columns={"Eigenkapitalkosten_Variabel_(in_%) - Kopie": "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"})

        if "Rating_od_wNote" in df:
            df = df[~df["Rating_od_wNote"].astype(str).isin({"-1", "-2"})]

        return df

    try:
        df1 = read_tab_em_icaap(FILE1_PATH)
        df2 = read_ek_basis_primaerbanken(FILE2_PATH)
        df_final = combine_and_finalize(df1, df2)

        OUTPUT_CSV.parent.mkdir(parents=True, exist_ok=True)
        df_final.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig", sep=";", decimal=",")
        #try:
            #df_final.to_excel(OUTPUT_XLSX, index=False)
        #except Exception:
            #pass
        #messagebox.showinfo("Success", f"Saved CSV: {OUTPUT_CSV}\nSaved XLSX: {OUTPUT_XLSX}")
    except Exception as e:
        messagebox.showerror("Error", str(e))
 
def process_csv():
    input_dir = r"U:\rlbnas1_rlb_bw_firw_z\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien"
    output_dir = r"U:\rlbnas1_rlb_bw_firw_z\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Upload_Dateien_Vorkalk"

    try:
        # Find the file containing the keyword "RisikoVerlustquote"
        for file_name in os.listdir(input_dir):
            if "RisikoVerlustquote" in file_name:
                file_path = os.path.join(input_dir, file_name)
                break
        else:
            messagebox.showerror("Error", "File with keyword 'RisikoVerlustquote' not found.")
            return

        # Define the column names explicitly since the file has no headers
        column_names = ["Column1", "Column2", "Column3", "Column4", "Column5"]

        # Load the CSV file into a DataFrame
        df = pd.read_csv(file_path, delimiter=';', decimal=',', header=None, names=column_names)

        # Rename columns
        df = df.rename(columns={
            "Column5": "LGD",
            "Column4": "Forderungsklasse",
            "Column2": "Sachkontonummer",
            "Column1": "BLZ",
            "Column3": "Rating_Kategorie"
        })

        # Ensure that BLZ column is treated as string for accurate filtering
        df['BLZ'] = df['BLZ'].astype(str)

        # Select rows where BLZ equals "34000"
        df_filtered = df[df["BLZ"] == "34000"]

        # Reorder columns
        df_filtered = df_filtered[["BLZ", "Rating_Kategorie", "Sachkontonummer", "Forderungsklasse", "LGD"]]

        # Replace dots with commas in the "LGD" column
        df_filtered["LGD"] = df_filtered["LGD"].astype(str).str.replace('.', ',', regex=False)

        # Convert the "LGD" column to numeric type (if needed)
        df_filtered["LGD"] = pd.to_numeric(df_filtered["LGD"].str.replace(',', '.'), errors='coerce')

        # Add a new column "Gueltig_Ab" with empty strings
        df_filtered["Gueltig_Ab"] = ""

        # Reorder columns again to include the new column
        df_filtered = df_filtered[["BLZ", "Gueltig_Ab", "Rating_Kategorie", "Sachkontonummer", "Forderungsklasse", "LGD"]]

        # Save the DataFrame to a new CSV file
        output_file_path = os.path.join(output_dir, "34000_RK_LGD.csv")
        df_filtered.to_csv(output_file_path, index=False, sep=';', decimal=',', encoding='utf-8')

        messagebox.showinfo("Success", f"File saved as {output_file_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def process_eigenmittel_verlustquote():
    input_dir = r"U:\rlbnas1_rlb_bw_firw_z\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien"
    output_dir = r"U:\rlbnas1_rlb_bw_firw_z\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Upload_Dateien_Vorkalk"

    try:
        # Find the file containing the keyword "EigenmittelVerlustquote"
        for file_name in os.listdir(input_dir):
            if "EigenmittelVerlustquote" in file_name:
                file_path = os.path.join(input_dir, file_name)
                break
        else:
            messagebox.showerror("Error", "File with keyword 'EigenmittelVerlustquote' not found.")
            return

        # Define the column names explicitly since the file has no headers
        column_names = ["Column1", "Column2", "Column3", "Column4", "Column5"]

        # Load the CSV file into a DataFrame
        df = pd.read_csv(file_path, delimiter=';', decimal=',', header=None, names=column_names)

        # Rename columns
        df = df.rename(columns={
            "Column5": "LGD",
            "Column4": "Forderungsklasse",
            "Column2": "Sachkontonummer",
            "Column1": "BLZ",
            "Column3": "Rating_Kategorie"
        })

        # Ensure that BLZ column is treated as string for accurate filtering
        df['BLZ'] = df['BLZ'].astype(str)

        # Select rows where BLZ equals "34000"
        df_filtered = df[df["BLZ"] == "34000"]

        # Reorder columns
        df_filtered = df_filtered[["BLZ", "Rating_Kategorie", "Sachkontonummer", "Forderungsklasse", "LGD"]]

        # Add a new column "Gueltig_Ab" with empty strings
        df_filtered["Gueltig_Ab"] = ""

        # Reorder columns again to include the new column
        df_filtered = df_filtered[["BLZ", "Gueltig_Ab", "Rating_Kategorie", "Sachkontonummer", "Forderungsklasse", "LGD"]]

        # Replace dots with commas in the "LGD" column
        df_filtered["LGD"] = df_filtered["LGD"].astype(str).str.replace('.', ',', regex=False)

        # Convert the "LGD" column to numeric type (if needed)
        df_filtered["LGD"] = pd.to_numeric(df_filtered["LGD"].str.replace(',', '.'), errors='coerce')

        # Save the DataFrame to a new CSV file
        output_file_path = os.path.join(output_dir, "34000_EK_LGD.csv")
        df_filtered.to_csv(output_file_path, index=False, sep=';', decimal=',', encoding='utf-8')

        messagebox.showinfo("Success", f"File saved as {output_file_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))
def create_gui():
    # Create the main window
    root = tk.Tk()
    root.title("Prozess Vorkalk")
    
    process_button = tk.Button(root, text="Datei hinspeichern", command=Emailprocessing)
    process_button.pack(pady=20)
    # Create a button and link it to the process_csv function
    process_button = tk.Button(root, text="Process RisikoVerlustquote", command=process_csv)
    process_button.pack(pady=20)

    process_button = tk.Button(root, text="Process EigenmittelVerlustquote", command=process_eigenmittel_verlustquote)
    process_button.pack(pady=20)
    
    process_button = tk.Button(root, text="Process 34000_EK_CCF", command=EK_CCF)
    process_button.pack(pady=20)
    
    process_button = tk.Button(root, text="Prozess RK_CCF", command=RK_CCF)
    process_button.pack(pady=20)
    
  
    process_button = tk.Button(root, text="Prozess EK_Basis", command=EKBasis)
    process_button.pack(pady=20)
    
    
    
    # Run the GUI main loop
    root.mainloop()

THIS IS CODE 1. and the following code is code 2. I want to add code 2 to code 1 without any problem. import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import messagebox

# === FILE PATHS ===
FILE1_PATH = Path(r"U:\rlbnas1_rlb_bw_firw_z\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien\Tab_Risiko_ICAAP.csv")  # no header
FILE2_PATH = Path(r"U:\rlbnas1_rlb_bw_firw_z\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien\Primär-bzw Raiffeisenbanken\RK_Basis_Primaerbanken.csv")  # has header
OUTPUT_CSV  = Path(r"U:\rlbnas1_rlb_bw_firw_z\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Upload_Dateien_Vorkalk\34000_RK_Basis.csv")
#OUTPUT_XLSX = Path(r"U:\rlbnas1_rlb_bw_firw_z\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Upload_Dateien_Vorkalk\34000_RK_Basis.xlsx")  # optional

DELIMS = [",", ";", "\t", "|"]

def RKBasis():
    def read_csv_flexible(path: Path, has_header: bool, prefer_utf8: bool = True) -> pd.DataFrame:
        encs = ["utf-8-sig", "latin-1"] if prefer_utf8 else ["latin-1", "utf-8-sig"]
        header = 0 if has_header else None
        last_err = None

        for enc in encs:
            # Auto-detect separator
            try:
                df = pd.read_csv(
                    path,
                    header=header,
                    sep=None,
                    engine="python",
                    encoding=enc,
                    on_bad_lines="skip",
                    skipinitialspace=True,
                    na_filter=False,
                )
                return df
            except Exception as e:
                last_err = e
            # Try fixed delimiters
            for sep in DELIMS:
                try:
                    df = pd.read_csv(
                        path,
                        header=header,
                        sep=sep,
                        engine="python",
                        encoding=enc,
                        on_bad_lines="skip",
                        skipinitialspace=True,
                        na_filter=False,
                    )
                    return df
                except Exception as e:
                    last_err = e
        raise RuntimeError(f"Could not read {path}: {last_err}")

    def to_int64(s: pd.Series) -> pd.Series:
        return pd.to_numeric(s, errors="coerce").astype("Int64")

    def to_float_from_maybe_comma(s: pd.Series) -> pd.Series:
        return pd.to_numeric(s.astype(str).str.replace(",", ".", regex=False), errors="coerce")

    def read_tab_risiko_icaap(file1: Path) -> pd.DataFrame:
        df = read_csv_flexible(file1, has_header=False)
        df.columns = [f"Column{i+1}" for i in range(df.shape[1])]

        for col in ["Column1", "Column7", "Column8"]:
            if col in df: df[col] = to_int64(df[col])
        for col in ["Column2", "Column9"]:
            if col in df: df[col] = df[col].astype(str)

        if "Column11" in df: df["Column11"] = pd.to_numeric(df["Column11"], errors="coerce")
        if "Column9" in df: df["Column9"] = pd.to_numeric(df["Column9"], errors="coerce")

        df = df.rename(columns={
            "Column1": "BLZ",
            "Column2": "Rating_od_wNote",
            "Column3": "Rating_Kategorie",
            "Column4": "Forderungsklasse",
            "Column5": "Risikokundengruppe",
        })

        if "Rating_od_wNote" in df:
            df["Copy of Rating_od_wNote"] = df["Rating_od_wNote"]
        df = df.rename(columns={"Rating_od_wNote": "Rating_od_wNote_Original"})

        reorder_1 = [
            "BLZ","Rating_od_wNote_Original","Copy of Rating_od_wNote","Rating_Kategorie",
            "Forderungsklasse","Risikokundengruppe","Column6","Column7","Column8","Column9","Column10","Column11"
        ]
        cols_existing = [c for c in reorder_1 if c in df.columns]
        df = df[cols_existing + [c for c in df.columns if c not in cols_existing]]
        df = df.rename(columns={"Copy of Rating_od_wNote": "Rating_od_wNote_Hilfsspalte"})

        if "Column6" in df: df = df.drop(columns=["Column6"])

        if "Rating_od_wNote_Hilfsspalte" in df:
            df["Rating_od_wNote_Hilfsspalte"] = df["Rating_od_wNote_Hilfsspalte"].astype(str).str.replace(".", ",", regex=False)

        def compute_rating(row):
            rk = str(row.get("Rating_Kategorie", ""))
            fk = str(row.get("Forderungsklasse", ""))
            if rk in {"10","11","12"} or fk in {"1","2","3","4","5"}:
                return row.get("Rating_od_wNote_Hilfsspalte", "")
            return row.get("Rating_od_wNote_Original", "")

        df["Rating_od_wNote"] = df.apply(compute_rating, axis=1)

        reorder_2 = [
            "BLZ","Rating_od_wNote","Rating_od_wNote_Original","Rating_od_wNote_Hilfsspalte",
            "Rating_Kategorie","Forderungsklasse","Risikokundengruppe","Column7","Column8","Column9","Column10","Column11"
        ]
        cols_existing = [c for c in reorder_2 if c in df.columns]
        df = df[cols_existing + [c for c in df.columns if c not in cols_existing]]

        df = df.rename(columns={
            "Column7": "Laufzeit_Von_(in_Tagen)",
            "Column8": "Laufzeit_Bis_(in_Tagen)",
            "Column9": "Risikokostensatz_Fix_(in_%)",
            "Column10": "Risikokostensatz_Variabel_(in_%)",
            "Column11": "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)",
        })

        if "Risikokundengruppe" in df:
            df["Risikokundengruppe"] = df["Risikokundengruppe"].astype(str).str.replace("-2", "", regex=False)

        if "Rating_od_wNote" in df:
            df = df[~df["Rating_od_wNote"].astype(str).isin({"-1,0","-1.0","-2,0","-2.0"})]

        df = df.drop(columns=[c for c in ["Rating_od_wNote_Original","Rating_od_wNote_Hilfsspalte"] if c in df.columns])

        final_order = [
            "BLZ","Rating_Kategorie","Rating_od_wNote","Forderungsklasse","Risikokundengruppe",
            "Laufzeit_Von_(in_Tagen)","Laufzeit_Bis_(in_Tagen)",
            "Risikokostensatz_Fix_(in_%)","Risikokostensatz_Variabel_(in_%)","Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"
        ]
        df = df[[c for c in final_order if c in df.columns]]

        for col in ["BLZ","Laufzeit_Von_(in_Tagen)","Laufzeit_Bis_(in_Tagen)"]:
            if col in df: df[col] = to_int64(df[col])
        for col in ["Risikokostensatz_Fix_(in_%)","Risikokostensatz_Variabel_(in_%)","Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"]:
            if col in df: df[col] = to_float_from_maybe_comma(df[col])

        return df

    def read_rk_basis_primaerbanken(file2: Path) -> pd.DataFrame:
        df = read_csv_flexible(file2, has_header=True)
        if "BLZ" not in df.columns:
            raise ValueError("Expected column 'BLZ' not found in RK_Basis_Primaerbanken.csv")
        df = df[df["BLZ"].astype(str) == "34"]

        ordered_cols = [
            "BLZ","Rating_od_wNote","Rating_Kategorie","Forderungsklasse","Risikokundengruppe",
            "Laufzeit_Von_(in_Tagen)","Laufzeit_Bis_(in_Tagen)","Risikokostensatz_Fix_(in_%)",
            "Risikokostensatz_Variabel_(in_%)","Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"
        ]
        missing = [c for c in ordered_cols if c not in df.columns]
        if missing:
            raise ValueError(f"Missing expected columns in RK_Basis_Primaerbanken.csv: {missing}")
        df = df[ordered_cols]

        for col in ["BLZ","Laufzeit_Von_(in_Tagen)","Laufzeit_Bis_(in_Tagen)"]:
            df[col] = to_int64(df[col])
        for col in ["Risikokostensatz_Fix_(in_%)","Risikokostensatz_Variabel_(in_%)","Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"]:
            df[col] = to_float_from_maybe_comma(df[col])

        return df

    def combine_and_finalize(df1: pd.DataFrame, df2: pd.DataFrame) -> pd.DataFrame:
        df = pd.concat([df1, df2], ignore_index=True)

        for col in ["BLZ","Laufzeit_Von_(in_Tagen)","Laufzeit_Bis_(in_Tagen)"]:
            if col in df: df[col] = to_int64(df[col])

        df = df[df["BLZ"] != 55000]

        if "Laufzeit_Bis_(in_Tagen)" in df:
            df["Laufzeit_Bis_(in_Tagen)"] = df["Laufzeit_Bis_(in_Tagen)"].replace(365, 366)
        if "Laufzeit_Von_(in_Tagen)" in df:
            df["Laufzeit_Von_(in_Tagen)"] = df["Laufzeit_Von_(in_Tagen)"].replace(366, 367)

        if "Risikokostensatz_Variabel_(in_%)" in df.columns:
            if "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)" in df.columns:
                df = df.drop(columns=["Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"])
            df["Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"] = df["Risikokostensatz_Variabel_(in_%)"]

        if "Rating_od_wNote" in df.columns:
            df = df[~df["Rating_od_wNote"].astype(str).isin({"-1","-2"})]

        df["Gueltig_Ab"] = ""

        final_order = [
            "BLZ","Gueltig_Ab","Rating_Kategorie","Rating_od_wNote","Forderungsklasse",
            "Risikokundengruppe","Laufzeit_Von_(in_Tagen)","Laufzeit_Bis_(in_Tagen)",
            "Risikokostensatz_Fix_(in_%)","Risikokostensatz_Variabel_(in_%)",
            "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"
        ]
        df = df[[c for c in final_order if c in df.columns]]
        return df

    try:
        df1 = read_tab_risiko_icaap(FILE1_PATH)
        df2 = read_rk_basis_primaerbanken(FILE2_PATH)
        df_final = combine_and_finalize(df1, df2)

        OUTPUT_CSV.parent.mkdir(parents=True, exist_ok=True)

        df_final.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig", sep=";", decimal=",")

        #try:
         #   df_final.to_excel(OUTPUT_XLSX, index=False)
        #except Exception:
         #   pass

        #messagebox.showinfo("Success", f"Saved CSV: {OUTPUT_CSV}\nSaved XLSX: {OUTPUT_XLSX}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# === GUI SETUP ===
def create_gui():
    root = tk.Tk()
    root.title("RK Basis Generator")

    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack(padx=10, pady=10)

    button = tk.Button(frame, text="Generate RK Basis", command=RKBasis, padx=10, pady=10)
    button.pack()

    root.mainloop()

create_gui()
# Run the GUI
create_gui()
