# -*- coding: utf-8 -*-
"""
Process EK files according to specified Power Query steps
"""

import pandas as pd
from pathlib import Path

# ---------------- EDIT THESE ----------------
FILE1_PATH = Path(r"U:\...\Originaldateien\Tab_EM_ICAAP.csv")  # no header
FILE2_PATH = Path(r"U:\...\Originaldateien\Primär-bzw Raiffeisenbanken\EK_Basis_Primaerbanken.csv")  # with header
OUTPUT_CSV  = Path(r"U:\...\Upload_Dateien_Vorkalk\EK_Basis_Final.csv")
OUTPUT_XLSX = Path(r"U:\...\Upload_Dateien_Vorkalk\EK_Basis_Final.xlsx")  # optional
# ---------------------------------------------

def to_int64(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce").astype("Int64")

def to_float_from_maybe_comma(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s.astype(str).str.replace(",", ".", regex=False), errors="coerce")

def read_tab_em_icaap(file1: Path) -> pd.DataFrame:
    # Read no header, add dummy names
    df = pd.read_csv(file1, header=None, dtype=str, na_filter=False)
    df.columns = [f"Column{i+1}" for i in range(df.shape[1])]

    # Transform column types
    for col in ["Column1", "Column7", "Column8"]:
        if col in df: df[col] = to_int64(df[col])
    for col in ["Column2", "Column9"]:
        if col in df: df[col] = df[col].astype(str)

    if "Column11" in df: df["Column11"] = pd.to_numeric(df["Column11"], errors="coerce")
    if "Column9"  in df: df["Column9"]  = pd.to_numeric(df["Column9"],  errors="coerce")

    # Rename main columns
    df = df.rename(columns={
        "Column1": "BLZ",
        "Column2": "Rating_od_wNote",
        "Column3": "Rating_Kategorie",
        "Column4": "Forderungsklasse",
        "Column5": "Risikokundengruppe",
    })

    # Duplicate + rename helper columns
    df["Copy of Rating_od_wNote"] = df["Rating_od_wNote"]
    df = df.rename(columns={"Rating_od_wNote": "Rating_od_wNote_Original"})

    # Reorder
    reorder1 = [
        "BLZ","Rating_od_wNote_Original","Copy of Rating_od_wNote","Rating_Kategorie",
        "Forderungsklasse","Risikokundengruppe","Column6","Column7","Column8","Column9","Column10","Column11"
    ]
    df = df[[c for c in reorder1 if c in df.columns] + [c for c in df.columns if c not in reorder1]]

    df = df.rename(columns={"Copy of Rating_od_wNote": "Rating_od_wNote_Hilfsspalte"})

    # Remove Column6
    if "Column6" in df: df = df.drop(columns=["Column6"])

    # Replace "." with "," in Hilfsspalte
    df["Rating_od_wNote_Hilfsspalte"] = df["Rating_od_wNote_Hilfsspalte"].astype(str).str.replace(".", ",", regex=False)

    # Add conditional Rating_od_wNote
    def compute_rating(row):
        rk = str(row.get("Rating_Kategorie", ""))
        fk = str(row.get("Forderungsklasse", ""))
        if rk in {"10","11","12"} or fk in {"1","2","3","4","5"}:
            return row.get("Rating_od_wNote_Hilfsspalte", "")
        return row.get("Rating_od_wNote_Original", "")
    df["Rating_od_wNote"] = df.apply(compute_rating, axis=1)

    # Reorder
    reorder2 = [
        "BLZ","Rating_od_wNote","Rating_od_wNote_Original","Rating_od_wNote_Hilfsspalte",
        "Rating_Kategorie","Forderungsklasse","Risikokundengruppe","Column7","Column8","Column9","Column10","Column11"
    ]
    df = df[[c for c in reorder2 if c in df.columns] + [c for c in df.columns if c not in reorder2]]

    # Rename Column7..Column11
    df = df.rename(columns={
        "Column7": "Laufzeit_Von_(in_Tagen)",
        "Column8": "Laufzeit_Bis_(in_Tagen)",
        "Column9": "Risikokostensatz_Fix_(in_%)",
        "Column10":"Risikokostensatz_Variabel_(in_%)",
        "Column11":"Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)",
    })

    # Replace "-2" in Risikokundengruppe
    df["Risikokundengruppe"] = df["Risikokundengruppe"].astype(str).str.replace("-2", "", regex=False)

    # Filter unwanted ratings
    df = df[~df["Rating_od_wNote"].astype(str).isin({"-1,0","-1.0","-2,0","-2.0"})]

    # Remove helper cols
    df = df.drop(columns=["Rating_od_wNote_Original", "Rating_od_wNote_Hilfsspalte"])

    # Reorder final
    final_order = [
        "BLZ","Rating_Kategorie","Rating_od_wNote","Forderungsklasse","Risikokundengruppe",
        "Laufzeit_Von_(in_Tagen)","Laufzeit_Bis_(in_Tagen)","Risikokostensatz_Fix_(in_%)",
        "Risikokostensatz_Variabel_(in_%)","Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"
    ]
    df = df[[c for c in final_order if c in df.columns]]

    # Remove 'Risikokundengruppe'
    if "Risikokundengruppe" in df: df = df.drop(columns=["Risikokundengruppe"])

    # Rename risk cost cols to EK cost cols
    df = df.rename(columns={
        "Risikokostensatz_Fix_(in_%)": "Eigenkapitalkosten_Fix_(in_%)",
        "Risikokostensatz_Variabel_(in_%)": "Eigenkapitalkosten_Variabel_(in_%)",
        "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)": "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)",
    })

    return df

def read_ek_basis_primaerbanken(file2: Path) -> pd.DataFrame:
    df = pd.read_csv(file2, dtype=str, na_filter=False)

    df = df[df["BLZ"] == "34"]

    for col in ["Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)", "BLZ"]:
        if col in df: df[col] = to_int64(df[col])
    for col in ["Eigenkapitalkosten_Fix_(in_%)","Eigenkapitalkosten_Variabel_(in_%)","Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"]:
        if col in df: df[col] = to_float_from_maybe_comma(df[col])

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

    df = df[df["Rating_Kategorie"].astype(str) != "9"]

    if "Eigenkapitalkosten_Variabel_(in_%)" in df:
        df["__copy_var"] = df["Eigenkapitalkosten_Variabel_(in_%)"]
        if "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)" in df:
            df = df.drop(columns=["Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"])
        df = df.rename(columns={"__copy_var": "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"})

    df = df[~df["Rating_od_wNote"].astype(str).isin({"-1","-2"})]

    return df

def main():
    df1 = read_tab_em_icaap(FILE1_PATH)
    df2 = read_ek_basis_primaerbanken(FILE2_PATH)
    df_final = combine_and_finalize(df1, df2)

    OUTPUT_CSV.parent.mkdir(parents=True, exist_ok=True)
    df_final.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig", sep=";", decimal=",")
    try:
        df_final.to_excel(OUTPUT_XLSX, index=False)
    except Exception:
        pass

    print(f"Saved CSV:  {OUTPUT_CSV}")
    print(f"Saved XLSX: {OUTPUT_XLSX}")

if __name__ == "__main__":
    main()
