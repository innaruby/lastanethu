import pandas as pd
from pathlib import Path

# ===== File Paths =====
FILE1_PATH = Path(r"U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien\Tab_EM_ICAAP.csv")
FILE2_PATH = Path(r"U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien\Primär-bzw Raiffeisenbanken\EK_Basis_Primaerbanken.csv")
OUTPUT_CSV  = Path(r"U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Upload_Dateien_Vorkalk\EK_Basis_Final.csv")
OUTPUT_XLSX = Path(r"U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Upload_Dateien_Vorkalk\EK_Basis_Final.xlsx")

DELIMS = [",", ";", "\t", "|"]

# ===== Generic CSV Reader =====
def read_csv_flexible(path: Path, header_row: bool, prefer_utf8: bool = True) -> pd.DataFrame:
    encodings = ["utf-8-sig", "latin-1"] if prefer_utf8 else ["latin-1", "utf-8-sig"]
    header = 0 if header_row else None
    last_err = None

    for enc in encodings:
        try:
            df = pd.read_csv(path, sep=None, engine="python", encoding=enc, header=header, skipinitialspace=True, na_filter=False)
            return df
        except Exception as e:
            last_err = e
        for sep in DELIMS:
            try:
                df = pd.read_csv(path, sep=sep, engine="python", encoding=enc, header=header, skipinitialspace=True, na_filter=False)
                return df
            except Exception as e:
                last_err = e
    raise RuntimeError(f"Could not read {path}: {last_err}")

def to_int64(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce").astype("Int64")

def to_float_from_maybe_comma(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s.astype(str).str.replace(",", ".", regex=False), errors="coerce")

# ===== Read and Transform Tab_EM_ICAAP =====
def read_tab_em_icaap(file1: Path) -> pd.DataFrame:
    df = read_csv_flexible(file1, header_row=False)
    df.columns = [f"Column{i+1}" for i in range(df.shape[1])]

    # Type conversions
    for col in ["Column1", "Column7", "Column8"]:
        if col in df: df[col] = to_int64(df[col])
    if "Column2" in df: df["Column2"] = df["Column2"].astype(str)
    if "Column9" in df: df["Column9"] = pd.to_numeric(df["Column9"], errors="coerce")
    if "Column11" in df: df["Column11"] = pd.to_numeric(df["Column11"], errors="coerce")

    # Rename
    df = df.rename(columns={
        "Column1": "BLZ",
        "Column2": "Rating_od_wNote",
        "Column3": "Rating_Kategorie",
        "Column4": "Forderungsklasse",
        "Column5": "Risikokundengruppe"
    })

    # Duplicate and reorder
    if "Rating_od_wNote" in df:
        df["Copy of Rating_od_wNote"] = df["Rating_od_wNote"]
    df = df.rename(columns={"Rating_od_wNote": "Rating_od_wNote_Original"})
    reorder_cols = ["BLZ", "Rating_od_wNote_Original", "Copy of Rating_od_wNote", "Rating_Kategorie", "Forderungsklasse", "Risikokundengruppe", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11"]
    df = df[[c for c in reorder_cols if c in df.columns] + [c for c in df.columns if c not in reorder_cols]]
    df = df.rename(columns={"Copy of Rating_od_wNote": "Rating_od_wNote_Hilfsspalte"})

    # Drop Column6
    if "Column6" in df: df = df.drop(columns=["Column6"])

    # Replace dot with comma in Hilfsspalte
    if "Rating_od_wNote_Hilfsspalte" in df:
        df["Rating_od_wNote_Hilfsspalte"] = df["Rating_od_wNote_Hilfsspalte"].astype(str).str.replace(".", ",", regex=False)

    # Compute new Rating_od_wNote
    def compute_rating(row):
        if str(row.get("Rating_Kategorie", "")) in {"10", "11", "12"} or str(row.get("Forderungsklasse", "")) in {"1", "2", "3", "4", "5"}:
            return row.get("Rating_od_wNote_Hilfsspalte", "")
        return row.get("Rating_od_wNote_Original", "")
    df["Rating_od_wNote"] = df.apply(compute_rating, axis=1)

    # Rename final columns
    df = df.rename(columns={
        "Column7": "Laufzeit_Von_(in_Tagen)",
        "Column8": "Laufzeit_Bis_(in_Tagen)",
        "Column9": "Risikokostensatz_Fix_(in_%)",
        "Column10": "Risikokostensatz_Variabel_(in_%)",
        "Column11": "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"
    })

    # Replace "-2" in Risikokundengruppe
    if "Risikokundengruppe" in df:
        df["Risikokundengruppe"] = df["Risikokundengruppe"].astype(str).str.replace("-2", "", regex=False)

    # Remove unwanted ratings
    df = df[~df["Rating_od_wNote"].astype(str).isin({"-1,0", "-1.0", "-2,0", "-2.0"})]

    # Drop helper cols
    df = df.drop(columns=[c for c in ["Rating_od_wNote_Original", "Rating_od_wNote_Hilfsspalte"] if c in df.columns])

    # Final reorder
    final_order = ["BLZ", "Rating_Kategorie", "Rating_od_wNote", "Forderungsklasse", "Risikokundengruppe", "Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)", "Risikokostensatz_Fix_(in_%)", "Risikokostensatz_Variabel_(in_%)", "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"]
    df = df[[c for c in final_order if c in df.columns]]

    # Drop Risikokundengruppe
    if "Risikokundengruppe" in df: df = df.drop(columns=["Risikokundengruppe"])

    # Rename to Eigenkapitalkosten
    df = df.rename(columns={
        "Risikokostensatz_Fix_(in_%)": "Eigenkapitalkosten_Fix_(in_%)",
        "Risikokostensatz_Variabel_(in_%)": "Eigenkapitalkosten_Variabel_(in_%)",
        "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)": "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"
    })

    return df

# ===== Read and Transform EK_Basis_Primaerbanken =====
def read_ek_basis_primaerbanken(file2: Path) -> pd.DataFrame:
    df = read_csv_flexible(file2, header_row=True)
    if "BLZ" not in df.columns:
        raise ValueError("Expected column 'BLZ' not found in EK_Basis_Primaerbanken.csv")

    df = df[df["BLZ"].astype(str) == "34"]
    for col in ["BLZ", "Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)"]:
        if col in df: df[col] = to_int64(df[col])
    for col in ["Eigenkapitalkosten_Fix_(in_%)", "Eigenkapitalkosten_Variabel_(in_%)", "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"]:
        if col in df: df[col] = to_float_from_maybe_comma(df[col])

    return df

# ===== Combine and Finalize =====
def combine_and_finalize(df1: pd.DataFrame, df2: pd.DataFrame) -> pd.DataFrame:
    df = pd.concat([df1, df2], ignore_index=True)

    # Filter out BLZ 55000
    df = df[df["BLZ"] != 55000]

    # Replace values
    if "Laufzeit_Bis_(in_Tagen)" in df: df["Laufzeit_Bis_(in_Tagen)"] = df["Laufzeit_Bis_(in_Tagen)"].replace(365, 366)
    if "Laufzeit_Von_(in_Tagen)" in df: df["Laufzeit_Von_(in_Tagen)"] = df["Laufzeit_Von_(in_Tagen)"].replace(366, 367)

    # Remove Rating_Kategorie 9
    if "Rating_Kategorie" in df: df = df[df["Rating_Kategorie"] != "9"]

    # Duplicate and adjust columns
    if "Eigenkapitalkosten_Variabel_(in_%)" in df:
        df["Eigenkapitalkosten_Variabel_(in_%) - Kopie"] = df["Eigenkapitalkosten_Variabel_(in_%)"]
        if "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)" in df:
            df = df.drop(columns=["Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"])
        df = df.rename(columns={"Eigenkapitalkosten_Variabel_(in_%) - Kopie": "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"})

    # Remove certain Rating_od_wNote
    if "Rating_od_wNote" in df: df = df[~df["Rating_od_wNote"].astype(str).isin({"-1", "-2"})]

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

    print(f"Saved CSV: {OUTPUT_CSV}")
    print(f"Saved XLSX: {OUTPUT_XLSX}")

if __name__ == "__main__":
    main()
