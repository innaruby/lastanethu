# -*- coding: utf-8 -*-
"""
Robust version:
- Auto-detects delimiter (;, , |, \t) and encoding (tries UTF-8-SIG, then latin-1).
- Uses engine='python' to handle irregular quoting.
- Tolerates stray/broken rows (on_bad_lines='skip' — switch to 'warn' if you want warnings).
- Then reproduces your Power Query steps 1:1.

Edit the three paths below before running.
"""

import pandas as pd
from pathlib import Path

# ----------------------- EDIT THESE -----------------------
FILE1_PATH = Path(r"U:\...\Originaldateien\Tab_Risiko_ICAAP.csv")  # no header
FILE2_PATH = Path(r"U:\...\Originaldateien\Primär-bzw Raiffeisenbanken\RK_Basis_Primaerbanken.csv")  # has header
OUTPUT_PATH = Path(r"U:\...\Upload_Dateien_Vorkalk\RK_Basis_Final.csv")
# ----------------------------------------------------------

DELIMS = [",", ";", "\t", "|"]

def read_csv_flexible(
    path: Path,
    has_header: bool,
    prefer_utf8: bool = True,
    low_memory: bool = False,
) -> pd.DataFrame:
    """
    Try encodings and delimiters to read messy CSVs robustly.
    """
    encs = ["utf-8-sig", "latin-1"] if prefer_utf8 else ["latin-1", "utf-8-sig"]
    header = 0 if has_header else None

    last_err = None
    for enc in encs:
        # First try automatic sep detection
        try:
            df = pd.read_csv(
                path,
                header=header,
                sep=None,               # let parser detect
                engine="python",
                encoding=enc,
                on_bad_lines="skip",    # or "warn" to see issues
                skipinitialspace=True,
                na_filter=False,
                low_memory=low_memory,
            )
            return df
        except Exception as e:
            last_err = e
        # Then try known delimiters
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
                    low_memory=low_memory,
                )
                return df
            except Exception as e:
                last_err = e
    # If we got here, everything failed
    raise RuntimeError(f"Could not read {path}: {last_err}")

def to_int64(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce").astype("Int64")

def to_float_from_maybe_comma(s: pd.Series) -> pd.Series:
    # Accept both "1,23" and "1.23"
    return pd.to_numeric(s.astype(str).str.replace(",", ".", regex=False), errors="coerce")

def read_tab_risiko_icaap(file1: Path) -> pd.DataFrame:
    # Read without header and add dummy names Column1..ColumnN
    df = read_csv_flexible(file1, has_header=False)
    df.columns = [f"Column{i+1}" for i in range(df.shape[1])]

    # TransformColumnTypes initial
    if "Column1" in df: df["Column1"] = to_int64(df["Column1"])
    if "Column2" in df: df["Column2"] = df["Column2"].astype(str)
    if "Column7" in df: df["Column7"] = to_int64(df["Column7"])
    if "Column8" in df: df["Column8"] = to_int64(df["Column8"])
    if "Column9" in df: df["Column9"] = df["Column9"].astype(str)

    # As numbers with en-US semantics
    if "Column11" in df: df["Column11"] = pd.to_numeric(df["Column11"], errors="coerce")
    if "Column9"  in df: df["Column9"]  = pd.to_numeric(df["Column9"],  errors="coerce")

    # Rename to business names
    rename_map = {
        "Column1": "BLZ",
        "Column2": "Rating_od_wNote",
        "Column3": "Rating_Kategorie",
        "Column4": "Forderungsklasse",
        "Column5": "Risikokundengruppe",
    }
    df = df.rename(columns=rename_map)

    # Duplicate + rename helpers
    if "Rating_od_wNote" in df:
        df["Copy of Rating_od_wNote"] = df["Rating_od_wNote"]
    df = df.rename(columns={"Rating_od_wNote": "Rating_od_wNote_Original"})

    # Reorder (keep only those that exist)
    reorder_1 = [
        "BLZ","Rating_od_wNote_Original","Copy of Rating_od_wNote","Rating_Kategorie",
        "Forderungsklasse","Risikokundengruppe","Column6","Column7","Column8","Column9","Column10","Column11"
    ]
    cols_existing = [c for c in reorder_1 if c in df.columns]
    df = df[cols_existing + [c for c in df.columns if c not in cols_existing]]
    df = df.rename(columns={"Copy of Rating_od_wNote": "Rating_od_wNote_Hilfsspalte"})

    # Remove Column6 if present
    if "Column6" in df: df = df.drop(columns=["Column6"])

    # Replace "." with "," in Hilfsspalte (string)
    if "Rating_od_wNote_Hilfsspalte" in df:
        df["Rating_od_wNote_Hilfsspalte"] = df["Rating_od_wNote_Hilfsspalte"].astype(str).str.replace(".", ",", regex=False)

    # New Rating_od_wNote with conditional
    def compute_rating(row):
        rk = str(row.get("Rating_Kategorie", ""))
        fk = str(row.get("Forderungsklasse", ""))
        if rk in {"10","11","12"} or fk in {"1","2","3","4","5"}:
            return row.get("Rating_od_wNote_Hilfsspalte", "")
        return row.get("Rating_od_wNote_Original", "")

    df["Rating_od_wNote"] = df.apply(compute_rating, axis=1)

    # Reorder again
    reorder_2 = [
        "BLZ","Rating_od_wNote","Rating_od_wNote_Original","Rating_od_wNote_Hilfsspalte",
        "Rating_Kategorie","Forderungsklasse","Risikokundengruppe","Column7","Column8","Column9","Column10","Column11"
    ]
    cols_existing = [c for c in reorder_2 if c in df.columns]
    df = df[cols_existing + [c for c in df.columns if c not in cols_existing]]

    # Rename to final column names
    df = df.rename(columns={
        "Column7": "Laufzeit_Von_(in_Tagen)",
        "Column8": "Laufzeit_Bis_(in_Tagen)",
        "Column9": "Risikokostensatz_Fix_(in_%)",
        "Column10":"Risikokostensatz_Variabel_(in_%)",
        "Column11":"Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)",
    })

    # Replace "-2" with "" in Risikokundengruppe
    if "Risikokundengruppe" in df:
        df["Risikokundengruppe"] = df["Risikokundengruppe"].astype(str).str.replace("-2", "", regex=False)

    # Drop rows with bad ratings
    if "Rating_od_wNote" in df:
        df = df[~df["Rating_od_wNote"].astype(str).isin({"-1,0","-1.0","-2,0","-2.0"})]

    # Drop helper columns
    df = df.drop(columns=[c for c in ["Rating_od_wNote_Original","Rating_od_wNote_Hilfsspalte"] if c in df.columns])

    # Final order
    final_order = [
        "BLZ","Rating_Kategorie","Rating_od_wNote","Forderungsklasse","Risikokundengruppe",
        "Laufzeit_Von_(in_Tagen)","Laufzeit_Bis_(in_Tagen)","Risikokostensatz_Fix_(in_%)",
        "Risikokostensatz_Variabel_(in_%)","Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"
    ]
    df = df[[c for c in final_order if c in df.columns]]

    # Cast likely numerics
    for col in ["BLZ","Laufzeit_Von_(in_Tagen)","Laufzeit_Bis_(in_Tagen)"]:
        if col in df: df[col] = to_int64(df[col])
    for col in ["Risikokostensatz_Fix_(in_%)","Risikokostensatz_Variabel_(in_%)","Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"]:
        if col in df: df[col] = to_float_from_maybe_comma(df[col])

    return df

def read_rk_basis_primaerbanken(file2: Path) -> pd.DataFrame:
    # This file has headers
    df = read_csv_flexible(file2, has_header=True)

    # Filter BLZ == "34"
    if "BLZ" not in df.columns:
        raise ValueError("Expected column 'BLZ' not found in RK_Basis_Primaerbanken.csv")
    df = df[df["BLZ"].astype(str) == "34"]

    # Reorder columns to exact order
    ordered_cols = [
        "BLZ","Rating_od_wNote","Rating_Kategorie","Forderungsklasse","Risikokundengruppe",
        "Laufzeit_Von_(in_Tagen)","Laufzeit_Bis_(in_Tagen)","Risikokostensatz_Fix_(in_%)",
        "Risikokostensatz_Variabel_(in_%)","Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"
    ]
    missing = [c for c in ordered_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Missing expected columns in RK_Basis_Primaerbanken.csv: {missing}")
    df = df[ordered_cols]

    # Cast types as required
    for col in ["BLZ","Laufzeit_Von_(in_Tagen)","Laufzeit_Bis_(in_Tagen)"]:
        df[col] = to_int64(df[col])
    for col in ["Risikokostensatz_Fix_(in_%)","Risikokostensatz_Variabel_(in_%)","Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"]:
        df[col] = to_float_from_maybe_comma(df[col])

    return df

def combine_and_finalize(df1: pd.DataFrame, df2: pd.DataFrame) -> pd.DataFrame:
    df = pd.concat([df1, df2], ignore_index=True)

    # Ensure numeric for the edits
    for col in ["BLZ","Laufzeit_Von_(in_Tagen)","Laufzeit_Bis_(in_Tagen)"]:
        if col in df: df[col] = to_int64(df[col])

    # Filter BLZ != 55000
    df = df[df["BLZ"] != 55000]

    # Replace 365->366 in Laufzeit_Bis; 366->367 in Laufzeit_Von
    if "Laufzeit_Bis_(in_Tagen)" in df:
        df["Laufzeit_Bis_(in_Tagen)"] = df["Laufzeit_Bis_(in_Tagen)"].replace(365, 366)
    if "Laufzeit_Von_(in_Tagen)" in df:
        df["Laufzeit_Von_(in_Tagen)"] = df["Laufzeit_Von_(in_Tagen)"].replace(366, 367)

    # Duplicate variable rate into 'nicht ausgenutzter Rahmen' (remove old first if exists)
    if "Risikokostensatz_Variabel_(in_%)" in df.columns:
        if "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)" in df.columns:
            df = df.drop(columns=["Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"])
        df["Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"] = df["Risikokostensatz_Variabel_(in_%)"]

    # Drop rows with Rating -1 or -2 (as pure values)
    if "Rating_od_wNote" in df.columns:
        df = df[~df["Rating_od_wNote"].astype(str).isin({"-1","-2"})]

    # Replace "."->"," in the two % cols (then back to float)
    for col in ["Risikokostensatz_Variabel_(in_%)","Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"]:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: str(x).replace(".", ",") if pd.notna(x) else x)
            df[col] = to_float_from_maybe_comma(df[col])  # store as float internally

    # Add empty 'Gueltig_Ab'
    df["Gueltig_Ab"] = ""

    # Final column order
    final_order = [
        "BLZ","Gueltig_Ab","Rating_Kategorie","Rating_od_wNote","Forderungsklasse",
        "Risikokundengruppe","Laufzeit_Von_(in_Tagen)","Laufzeit_Bis_(in_Tagen)",
        "Risikokostensatz_Fix_(in_%)","Risikokostensatz_Variabel_(in_%)",
        "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"
    ]
    df = df[[c for c in final_order if c in df.columns]]

    return df

def main():
    df1 = read_tab_risiko_icaap(FILE1_PATH)
    df2 = read_rk_basis_primaerbanken(FILE2_PATH)
    df_final = combine_and_finalize(df1, df2)

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    # Save with UTF-8 BOM for Excel-friendly behavior
    df_final.to_csv(OUTPUT_PATH, index=False, encoding="utf-8-sig")
    print(f"Done. Saved: {OUTPUT_PATH}")

if __name__ == "__main__":
    main()
