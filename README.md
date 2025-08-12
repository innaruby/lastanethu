# build_rk_basis_final.py
# -*- coding: utf-8 -*-
"""
Recreates the given Power Query (M) steps in pandas.

Inputs:
  1) U:\...\Originaldateien\Tab_Risiko_ICAAP.csv         (no header row)
  2) U:\...\Originaldateien\Primär-bzw Raiffeisenbanken\RK_Basis_Primaerbanken.csv  (has header row)

Output:
  U:\...\Upload_Dateien_Vorkalk\RK_Basis_Final.csv

Notes:
- The script is robust to comma/semicolon separators and BOM encodings.
- It follows the exact transformation order you specified.
"""

import os
import sys
import pandas as pd

# ---------------------------
# Config: input/output paths
# ---------------------------
FILE1 = r"U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien\Tab_Risiko_ICAAP.csv"
FILE2 = r"U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien\Primär-bzw Raiffeisenbanken\RK_Basis_Primaerbanken.csv"
OUT  = r"U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Upload_Dateien_Vorkalk\RK_Basis_Final.csv"

# ---------------------------
# Helpers
# ---------------------------
def read_csv_safely(path, has_header=True):
    """
    Tries a couple of encodings; lets pandas sniff the delimiter.
    """
    encodings = ["utf-8-sig", "cp1252", "latin1"]
    last_err = None
    for enc in encodings:
        try:
            if has_header:
                return pd.read_csv(path, sep=None, engine="python", encoding=enc, dtype=str)
            else:
                return pd.read_csv(path, sep=None, engine="python", encoding=enc, dtype=str, header=None)
        except Exception as e:
            last_err = e
    raise last_err

def to_int_nullable(s):
    return pd.to_numeric(s, errors="coerce").astype("Int64")

def parse_en_number(s):
    """
    Parse string with EN (dot-decimal) into float.
    """
    return pd.to_numeric(s, errors="coerce")

def ensure_cols(df, names):
    """
    Ensure a DataFrame has these columns (if missing, create as empty string).
    """
    for c in names:
        if c not in df.columns:
            df[c] = ""
    return df

# ---------------------------
# Step 1: Process Tab_Risiko_ICAAP.csv  (no headers)
# ---------------------------
def process_tab_risiko_icaap(file1):
    # Load without headers; add dummy column names
    df = read_csv_safely(file1, has_header=False)
    df.columns = [f"Column{i+1}" for i in range(df.shape[1])]

    # Transform column types
    # = Table.TransformColumnTypes(Quelle,{{"Column1", Int64.Type}, {"Column2", type text}, {"Column7", Int64.Type}, {"Column8", Int64.Type}, {"Column9", type text}})
    df["Column1"] = to_int_nullable(df.get("Column1"))
    df["Column2"] = df.get("Column2").astype(str)
    if "Column7" in df.columns:
        df["Column7"] = to_int_nullable(df["Column7"])
    if "Column8" in df.columns:
        df["Column8"] = to_int_nullable(df["Column8"])
    if "Column9" in df.columns:
        df["Column9"] = df["Column9"].astype(str)

    # = Table.TransformColumnTypes(#"Geänderter Typ", {{"Column11", type number}, {"Column9", type number}}, "en-US")
    # Interpret Column11 and Column9 as EN numbers (dot decimal)
    for c in ["Column11", "Column9"]:
        if c in df.columns:
            df[c] = parse_en_number(df[c])

    # = Table.RenameColumns(... {"Column1","BLZ"}, {"Column2","Rating_od_wNote"}, {"Column3","Rating_Kategorie"},
    #                        {"Column4","Forderungsklasse"}, {"Column5","Risikokundengruppe"})
    rename_map = {
        "Column1": "BLZ",
        "Column2": "Rating_od_wNote",
        "Column3": "Rating_Kategorie",
        "Column4": "Forderungsklasse",
        "Column5": "Risikokundengruppe",
    }
    df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns}, inplace=True)

    # Ensure columns exist for later order
    df = ensure_cols(df, ["BLZ","Rating_od_wNote","Rating_Kategorie","Forderungsklasse","Risikokundengruppe",
                          "Column6","Column7","Column8","Column9","Column10","Column11"])

    # = Table.DuplicateColumn(..., "Rating_od_wNote", "Copy of Rating_od_wNote")
    df["Copy of Rating_od_wNote"] = df["Rating_od_wNote"]

    # = Table.RenameColumns(... {"Rating_od_wNote","Rating_od_wNote_Original"})
    df.rename(columns={"Rating_od_wNote": "Rating_od_wNote_Original"}, inplace=True)

    # = Table.ReorderColumns(... desired order)
    order1 = ["BLZ", "Rating_od_wNote_Original", "Copy of Rating_od_wNote", "Rating_Kategorie",
              "Forderungsklasse", "Risikokundengruppe", "Column6", "Column7", "Column8",
              "Column9", "Column10", "Column11"]
    df = df[order1]

    # = Table.RenameColumns(... {"Copy of Rating_od_wNote","Rating_od_wNote_Hilfsspalte"})
    df.rename(columns={"Copy of Rating_od_wNote": "Rating_od_wNote_Hilfsspalte"}, inplace=True)

    # = Table.RemoveColumns(... {"Column6"})
    if "Column6" in df.columns:
        df.drop(columns=["Column6"], inplace=True)

    # = Table.ReplaceValue(... ".", ",", Replacer.ReplaceText, {"Rating_od_wNote_Hilfsspalte"})
    df["Rating_od_wNote_Hilfsspalte"] = df["Rating_od_wNote_Hilfsspalte"].astype(str).str.replace(".", ",", regex=False)

    # = Table.AddColumn(... "Rating_od_wNote", each if (Rating_Kategorie in 10..12 OR Forderungsklasse in 1..5)
    #                   then Rating_od_wNote_Hilfsspalte else Rating_od_wNote_Original)
    def new_rating(row):
        if str(row["Rating_Kategorie"]) in {"10","11","12"} or str(row["Forderungsklasse"]) in {"1","2","3","4","5"}:
            return row["Rating_od_wNote_Hilfsspalte"]
        return row["Rating_od_wNote_Original"]

    df["Rating_od_wNote"] = df.apply(new_rating, axis=1)

    # = Table.ReorderColumns(... bring new Rating forward)
    order2 = ["BLZ", "Rating_od_wNote", "Rating_od_wNote_Original", "Rating_od_wNote_Hilfsspalte",
              "Rating_Kategorie", "Forderungsklasse", "Risikokundengruppe", "Column7", "Column8",
              "Column9", "Column10", "Column11"]
    df = df[order2]

    # = Table.RenameColumns(... Column7..11 to proper names)
    df.rename(columns={
        "Column7":  "Laufzeit_Von_(in_Tagen)",
        "Column8":  "Laufzeit_Bis_(in_Tagen)",
        "Column9":  "Risikokostensatz_Fix_(in_%)",
        "Column10": "Risikokostensatz_Variabel_(in_%)",
        "Column11": "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"
    }, inplace=True)

    # = Table.ReplaceValue(... "-2","", {"Risikokundengruppe"})
    df["Risikokundengruppe"] = df["Risikokundengruppe"].astype(str).replace("-2", "")

    # = Table.SelectRows(... Rating_od_wNote not in ["-1,0","-1.0","-2,0","-2.0"])
    bad_vals = {"-1,0", "-1.0", "-2,0", "-2.0"}
    df = df[~df["Rating_od_wNote"].astype(str).isin(bad_vals)]

    # = Table.RemoveColumns(... {"Rating_od_wNote_Original","Rating_od_wNote_Hilfsspalte"})
    df.drop(columns=["Rating_od_wNote_Original", "Rating_od_wNote_Hilfsspalte"], inplace=True)

    # = Table.ReorderColumns(... final order for df1)
    final1 = ["BLZ", "Rating_Kategorie", "Rating_od_wNote", "Forderungsklasse", "Risikokundengruppe",
              "Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)", "Risikokostensatz_Fix_(in_%)",
              "Risikokostensatz_Variabel_(in_%)", "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"]
    df = ensure_cols(df, final1)
    df = df[final1]

    # Ensure integer numeric types on Laufzeit columns (as in the beginning)
    df["BLZ"] = to_int_nullable(df["BLZ"])
    df["Laufzeit_Von_(in_Tagen)"] = to_int_nullable(df["Laufzeit_Von_(in_Tagen)"])
    df["Laufzeit_Bis_(in_Tagen)"] = to_int_nullable(df["Laufzeit_Bis_(in_Tagen)"])

    # Riskokosten columns as numbers (already parsed for fix & var earlier; enforce numeric)
    for c in ["Risikokostensatz_Fix_(in_%)",
              "Risikokostensatz_Variabel_(in_%)",
              "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    return df

# ---------------------------
# Step 2: Process RK_Basis_Primaerbanken.csv (has headers)
# ---------------------------
def process_rk_basis_primaerbanken(file2):
    df = read_csv_safely(file2, has_header=True)

    # = Table.PromoteHeaders(Quelle)  -> already done by reading with header row

    # Ensure required columns exist
    needed = ["BLZ", "Rating_od_wNote", "Rating_Kategorie", "Forderungsklasse", "Risikokundengruppe",
              "Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)", "Risikokostensatz_Fix_(in_%)",
              "Risikokostensatz_Variabel_(in_%)", "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"]
    df = ensure_cols(df, needed)

    # = Table.SelectRows(... each ([BLZ] = "34"))
    df = df[df["BLZ"].astype(str) == "34"]

    # = Table.ReorderColumns(... as given)
    df = df[needed]

    # = Table.TransformColumnTypes(... {"BLZ", Int64}, risk cost numbers, Laufzeit int)
    df["BLZ"] = to_int_nullable(df["BLZ"])
    df["Laufzeit_Von_(in_Tagen)"] = to_int_nullable(df["Laufzeit_Von_(in_Tagen)"])
    df["Laufzeit_Bis_(in_Tagen)"] = to_int_nullable(df["Laufzeit_Bis_(in_Tagen)"])
    for c in ["Risikokostensatz_Fix_(in_%)",
              "Risikokostensatz_Variabel_(in_%)",
              "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"]:
        # Parse EN numbers if present
        df[c] = pd.to_numeric(df[c], errors="coerce")

    return df

# ---------------------------
# Step 3: Combine & final transforms
# ---------------------------
def combine_and_finalize(df1, df2):
    df = pd.concat([df1, df2], ignore_index=True)

    # = Table.SelectRows(Quelle, each ([BLZ] <> 55000))
    # Ensure BLZ is numeric-like for comparison
    # (Coerce to Int64 if not already)
    df["BLZ"] = to_int_nullable(df["BLZ"])
    df = df[df["BLZ"] != 55000]

    # = Table.ReplaceValue(... 365 -> 366, {"Laufzeit_Bis_(in_Tagen)"})
    df["Laufzeit_Bis_(in_Tagen)"] = to_int_nullable(df["Laufzeit_Bis_(in_Tagen)"]).replace(365, 366)

    # = Table.ReplaceValue(... 366 -> 367, {"Laufzeit_Von_(in_Tagen)"})
    df["Laufzeit_Von_(in_Tagen)"] = to_int_nullable(df["Laufzeit_Von_(in_Tagen)"]).replace(366, 367)

    # = Table.DuplicateColumn(... "Risikokostensatz_Variabel_(in_%)" -> copy)
    # = Table.RemoveColumns(... remove "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)")
    # = Table.RenameColumns(... copy -> "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)")
    df["__tmp_var_copy__"] = pd.to_numeric(df["Risikokostensatz_Variabel_(in_%)"], errors="coerce")
    if "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)" in df.columns:
        df.drop(columns=["Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"], inplace=True)
    df.rename(columns={"__tmp_var_copy__": "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"}, inplace=True)

    # = Table.SelectRows(... Rating_od_wNote not in ["-1","-2"])
    df = df[~df["Rating_od_wNote"].astype(str).isin({"-1", "-2"})]

    # = Table.ReplaceValue(... ".", ",", {"Risikokostensatz_Variabel_(in_%)","Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"})
    # Then = Table.TransformColumnTypes(... both as number)
    for c in ["Risikokostensatz_Variabel_(in_%)", "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"]:
        # Perform text replacement for display-like step, then convert back to number
        # (M-code replaced to comma, then turned into numbers; we mimic by round-tripping)
        as_text = df[c].astype(str).str.replace(".", ",", regex=False)
        # Convert to number by reversing comma to dot
        df[c] = pd.to_numeric(as_text.str.replace(",", ".", regex=False), errors="coerce")

    # = Table.AddColumn(... "Gueltig_Ab", each "")
    df["Gueltig_Ab"] = ""

    # = Table.ReorderColumns(... final order)
    final_cols = ["BLZ", "Gueltig_Ab", "Rating_Kategorie", "Rating_od_wNote", "Forderungsklasse",
                  "Risikokundengruppe", "Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)",
                  "Risikokostensatz_Fix_(in_%)", "Risikokostensatz_Variabel_(in_%)",
                  "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"]
    df = df[final_cols]

    return df

def main():
    # Process both inputs
    df1 = process_tab_risiko_icaap(FILE1)
    df2 = process_rk_basis_primaerbanken(FILE2)

    # Combine and finish
    df_final = combine_and_finalize(df1, df2)

    # Ensure output directory exists
    out_dir = os.path.dirname(OUT)
    if out_dir and not os.path.isdir(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    # Save with semicolon sep, as is common in DE setups
    df_final.to_csv(OUT, index=False, sep=";")
    print(f"Saved: {OUT}")
    print(f"Rows: {len(df_final):,} | Columns: {len(df_final.columns)}")

if __name__ == "__main__":
    # Optional: basic guard to ensure pandas is present
    try:
        import pandas as _p  # noqa
    except ImportError:
        print("Please install pandas first:  pip install pandas")
        sys.exit(1)
    main()
