

import pandas as pd
from pathlib import Path

# ---------------- EDIT THESE ----------------
FILE1_PATH = Path(r"U:\...\Originaldateien\Tab_EM_ICAAP.csv")  # no header
FILE2_PATH = Path(r"U:\...\Originaldateien\Primär-bzw Raiffeisenbanken\EK_Basis_Primaerbanken.csv")  # with header
OUTPUT_CSV  = Path(r"U:\...\Upload_Dateien_Vorkalk\EK_Basis_Final.csv")
OUTPUT_XLSX = Path(r"U:\...\Upload_Dateien_Vorkalk\EK_Basis_Final.xlsx")  # optional
# ---------------------------------------------

DELIMS = [",", ";", "\t", "|"]

def read_csv_flexible(path: Path, has_header: bool, prefer_utf8: bool = True) -> pd.DataFrame:
    encs = ["utf-8-sig", "latin-1"] if prefer_utf8 else ["latin-1", "utf-8-sig"]
    header = 0 if has_header else None
    last_err = None
    for enc in encs:
        # 1) sep=None (sniffer)
        try:
            return pd.read_csv(
                path, header=header, sep=None, engine="python",
                encoding=enc, on_bad_lines="skip", skipinitialspace=True, na_filter=False
            )
        except Exception as e:
            last_err = e
        # 2) try a few common delimiters
        for sep in DELIMS:
            try:
                return pd.read_csv(
                    path, header=header, sep=sep, engine="python",
                    encoding=enc, on_bad_lines="skip", skipinitialspace=True, na_filter=False
                )
            except Exception as e:
                last_err = e
    raise RuntimeError(f"Could not read {path}: {last_err}")

def to_int64(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce").astype("Int64")

def to_float_from_maybe_comma(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s.astype(str).str.replace(",", ".", regex=False), errors="coerce")

def ensure_required_columns(df: pd.DataFrame, required_raw_cols: list, context: str):
    missing = [c for c in required_raw_cols if c not in df.columns]
    if missing:
        # Give a helpful error instead of a KeyError later
        raise ValueError(
            f"{context}: expected the following columns but they are missing: {missing}\n"
            f"Detected columns: {list(df.columns)}\n"
            "Tip: This usually means the delimiter/encoding was different "
            "or the file has fewer fields than expected."
        )

def read_tab_em_icaap(file1: Path) -> pd.DataFrame:
    # Read WITHOUT header; add dummy names Column1..ColumnN
    df = read_csv_flexible(file1, has_header=False)
    df.columns = [f"Column{i+1}" for i in range(df.shape[1])]

    # We need at least up to Column11 for this logic
    required_raw = ["Column1", "Column2", "Column3", "Column4", "Column5",
                    "Column7", "Column8", "Column9", "Column10", "Column11"]
    ensure_required_columns(df, required_raw, "Tab_EM_ICAAP.csv")

    # Transform types (matching your M steps)
    for col in ["Column1", "Column7", "Column8"]:
        df[col] = to_int64(df[col])
    for col in ["Column2", "Column9"]:
        df[col] = df[col].astype(str)

    df["Column11"] = pd.to_numeric(df["Column11"], errors="coerce")
    df["Column9"]  = pd.to_numeric(df["Column9"],  errors="coerce")

    # Rename to business names
    df = df.rename(columns={
        "Column1": "BLZ",
        "Column2": "Rating_od_wNote",
        "Column3": "Rating_Kategorie",
        "Column4": "Forderungsklasse",
        "Column5": "Risikokundengruppe",
    })

    # Duplicate + rename helper
    df["Copy of Rating_od_wNote"] = df["Rating_od_wNote"]  # safe now because we validated Column2 earlier
    df = df.rename(columns={"Rating_od_wNote": "Rating_od_wNote_Original"})

    # Reorder
    reorder1 = [
        "BLZ","Rating_od_wNote_Original","Copy of Rating_od_wNote","Rating_Kategorie",
        "Forderungsklasse","Risikokundengruppe","Column6","Column7","Column8","Column9","Column10","Column11"
    ]
    cols_existing = [c for c in reorder1 if c in df.columns]
    df = df[cols_existing + [c for c in df.columns if c not in cols_existing]]

    # Rename copy -> Hilfsspalte
    df = df.rename(columns={"Copy of Rating_od_wNote": "Rating_od_wNote_Hilfsspalte"})

    # Remove Column6 if present
    if "Column6" in df:
        df = df.drop(columns=["Column6"])

    # Replace "." -> "," in Hilfsspalte (string)
    df["Rating_od_wNote_Hilfsspalte"] = df["Rating_od_wNote_Hilfsspalte"].astype(str).str.replace(".", ",", regex=False)

    # Conditional new Rating_od_wNote
    def compute_rating(row):
        rk = str(row.get("Rating_Kategorie", ""))
        fk = str(row.get("Forderungsklasse", ""))
        if rk in {"10", "11", "12"} or fk in {"1", "2", "3", "4", "5"}:
            return row.get("Rating_od_wNote_Hilfsspalte", "")
        return row.get("Rating_od_wNote_Original", "")

    df["Rating_od_wNote"] = df.apply(compute_rating, axis=1)

    # Reorder again
    reorder2 = [
        "BLZ","Rating_od_wNote","Rating_od_wNote_Original","Rating_od_wNote_Hilfsspalte",
        "Rating_Kategorie","Forderungsklasse","Risikokundengruppe","Column7","Column8","Column9","Column10","Column11"
    ]
    cols_existing = [c for c in reorder2 if c in df.columns]
    df = df[cols_existing + [c for c in df.columns if c not in cols_existing]]

    # Rename columns 7..11 to final names
    df = df.rename(columns={
        "Column7": "Laufzeit_Von_(in_Tagen)",
        "Column8": "Laufzeit_Bis_(in_Tagen)",
        "Column9": "Risikokostensatz_Fix_(in_%)",
        "Column10": "Risikokostensatz_Variabel_(in_%)",
        "Column11": "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)",
    })

    # Replace "-2" with "" in Risikokundengruppe
    df["Risikokundengruppe"] = df["Risikokundengruppe"].astype(str).str.replace("-2", "", regex=False)

    # Remove unwanted ratings
    df = df[~df["Rating_od_wNote"].astype(str).isin({"-1,0", "-1.0", "-2,0", "-2.0"})]

    # Remove helper columns
    df = df.drop(columns=["Rating_od_wNote_Original", "Rating_od_wNote_Hilfsspalte"], errors="ignore")

    # Final order (before further removals/renames)
    final_order = [
        "BLZ","Rating_Kategorie","Rating_od_wNote","Forderungsklasse","Risikokundengruppe",
        "Laufzeit_Von_(in_Tagen)","Laufzeit_Bis_(in_Tagen)",
        "Risikokostensatz_Fix_(in_%)","Risikokostensatz_Variabel_(in_%)","Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"
    ]
    df = df[[c for c in final_order if c in df.columns]]

    # Remove 'Risikokundengruppe'
    if "Risikokundengruppe" in df:
        df = df.drop(columns=["Risikokundengruppe"])

    # Rename risk → EK columns
    df = df.rename(columns={
        "Risikokostensatz_Fix_(in_%)": "Eigenkapitalkosten_Fix_(in_%)",
        "Risikokostensatz_Variabel_(in_%)": "Eigenkapitalkosten_Variabel_(in_%)",
        "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)": "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)",
    })

    return df

def read_ek_basis_primaerbanken(file2: Path) -> pd.DataFrame:
    df = read_csv_flexible(file2, has_header=True)

    if "BLZ" not in df.columns:
        raise ValueError("EK_Basis_Primaerbanken.csv: expected column 'BLZ' but it is missing.")

    df = df[df["BLZ"].astype(str) == "34"]

    for col in ["Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)", "BLZ"]:
        if col in df: df[col] = to_int64(df[col])
    for col in [
        "Eigenkapitalkosten_Fix_(in_%)",
        "Eigenkapitalkosten_Variabel_(in_%)",
        "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"
    ]:
        if col in df: df[col] = to_float_from_maybe_comma(df[col])

    return df

def combine_and_finalize(df1: pd.DataFrame, df2: pd.DataFrame) -> pd.DataFrame:
    df = pd.concat([df1, df2], ignore_index=True)

    for col in ["BLZ", "Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)"]:
        if col in df: df[col] = to_int64(df[col])

    # BLZ != 55000
    df = df[df["BLZ"] != 55000]

    # 365->366 (Bis), 366->367 (Von)
    if "Laufzeit_Bis_(in_Tagen)" in df:
        df["Laufzeit_Bis_(in_Tagen)"] = df["Laufzeit_Bis_(in_Tagen)"].replace(365, 366)
    if "Laufzeit_Von_(in_Tagen)" in df:
        df["Laufzeit_Von_(in_Tagen)"] = df["Laufzeit_Von_(in_Tagen)"].replace(366, 367)

    # Filter Rating_Kategorie != "9"
    if "Rating_Kategorie" in df:
        df = df[df["Rating_Kategorie"].astype(str) != "9"]

    # Duplicate variable EK cost into 'nicht ausgenutzter Rahmen'
    if "Eigenkapitalkosten_Variabel_(in_%)" in df.columns:
        if "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)" in df.columns:
            df = df.drop(columns=["Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"])
        df["Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"] = df["Eigenkapitalkosten_Variabel_(in_%)"]

    # Keep rows with Rating_od_wNote not -1/-2 (only if column exists)
    if "Rating_od_wNote" in df.columns:
        df = df[~df["Rating_od_wNote"].astype(str).isin({"-1", "-2"})]

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
