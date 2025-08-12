

import pandas as pd
from pathlib import Path




FILE1_PATH = Path(r"U:\...\Originaldateien\Tab_Risiko_ICAAP.csv")
FILE2_PATH = Path(r"U:\...\Originaldateien\Primär-bzw Raiffeisenbanken\RK_Basis_Primaerbanken.csv")
OUTPUT_PATH = Path(r"U:\...\Upload_Dateien_Vorkalk\RK_Basis_Final.csv")


def read_tab_risiko_icaap(file1: Path) -> pd.DataFrame:
    # Read without header, add dummy column names
    df = pd.read_csv(file1, header=None, dtype=str, na_filter=False)
    df.columns = [f"Column{i+1}" for i in range(df.shape[1])]

    # = Table.TransformColumnTypes(Quelle, {Column1:Int64, Column2:text, Column7:Int64, Column8:Int64, Column9:text})
    # Cast with care; keep others as-is
    def to_int_series(s: pd.Series) -> pd.Series:
        return pd.to_numeric(s, errors="coerce").astype("Int64")

    def to_text_series(s: pd.Series) -> pd.Series:
        return s.astype(str)

    for col in ["Column1", "Column7", "Column8"]:
        if col in df.columns:
            df[col] = to_int_series(df[col])
    for col in ["Column2", "Column9"]:
        if col in df.columns:
            df[col] = to_text_series(df[col])

    # = Table.TransformColumnTypes(..., {Column11:number, Column9:number}, "en-US")
    # Interpreting as US decimals (dot). to_numeric handles this.
    for col in ["Column11", "Column9"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # = Table.RenameColumns(..., Column1->BLZ, Column2->Rating_od_wNote, Column3->Rating_Kategorie,
    #                        Column4->Forderungsklasse, Column5->Risikokundengruppe)
    rename_map = {
        "Column1": "BLZ",
        "Column2": "Rating_od_wNote",
        "Column3": "Rating_Kategorie",
        "Column4": "Forderungsklasse",
        "Column5": "Risikokundengruppe",
    }
    df = df.rename(columns=rename_map)

    # = Table.DuplicateColumn(..., "Rating_od_wNote", "Copy of Rating_od_wNote")
    if "Rating_od_wNote" in df.columns:
        df["Copy of Rating_od_wNote"] = df["Rating_od_wNote"]

    # = Table.RenameColumns(..., Rating_od_wNote->Rating_od_wNote_Original)
    df = df.rename(columns={"Rating_od_wNote": "Rating_od_wNote_Original"})

    # = Table.ReorderColumns(..., ["BLZ","Rating_od_wNote_Original","Copy of Rating_od_wNote","Rating_Kategorie",
    #                               "Forderungsklasse","Risikokundengruppe","Column6","Column7","Column8","Column9","Column10","Column11"])
    reorder_cols = [
        "BLZ", "Rating_od_wNote_Original", "Copy of Rating_od_wNote", "Rating_Kategorie",
        "Forderungsklasse", "Risikokundengruppe", "Column6", "Column7", "Column8", "Column9",
        "Column10", "Column11"
    ]
    # Keep only those that exist, then append any remaining columns (stable)
    cols_existing = [c for c in reorder_cols if c in df.columns]
    df = df[cols_existing + [c for c in df.columns if c not in cols_existing]]

    # = Table.RenameColumns(..., "Copy of Rating_od_wNote"->"Rating_od_wNote_Hilfsspalte")
    df = df.rename(columns={"Copy of Rating_od_wNote": "Rating_od_wNote_Hilfsspalte"})

    # = Table.RemoveColumns(..., {"Column6"})
    if "Column6" in df.columns:
        df = df.drop(columns=["Column6"])

    # = Table.ReplaceValue(..., ".", ",", {"Rating_od_wNote_Hilfsspalte"})
    if "Rating_od_wNote_Hilfsspalte" in df.columns:
        df["Rating_od_wNote_Hilfsspalte"] = df["Rating_od_wNote_Hilfsspalte"].astype(str).str.replace(".", ",", regex=False)

    # = Table.AddColumn(..., "Rating_od_wNote",
    #     each if Rating_Kategorie in {"10","11","12"} or Forderungsklasse in {"1","2","3","4","5"}
    #     then Rating_od_wNote_Hilfsspalte else Rating_od_wNote_Original)
    def compute_rating(row):
        rk = str(row.get("Rating_Kategorie", ""))
        fk = str(row.get("Forderungsklasse", ""))
        if rk in {"10", "11", "12"} or fk in {"1", "2", "3", "4", "5"}:
            return row.get("Rating_od_wNote_Hilfsspalte", "")
        return row.get("Rating_od_wNote_Original", "")

    df["Rating_od_wNote"] = df.apply(compute_rating, axis=1)

    # = Table.ReorderColumns(... place "Rating_od_wNote" etc.)
    reorder2 = [
        "BLZ", "Rating_od_wNote", "Rating_od_wNote_Original", "Rating_od_wNote_Hilfsspalte",
        "Rating_Kategorie", "Forderungsklasse", "Risikokundengruppe", "Column7", "Column8",
        "Column9", "Column10", "Column11"
    ]
    cols_existing = [c for c in reorder2 if c in df.columns]
    df = df[cols_existing + [c for c in df.columns if c not in cols_existing]]

    # = Table.RenameColumns(..., Column7->Laufzeit_Von_(in_Tagen), Column8->Laufzeit_Bis_(in_Tagen),
    #                                Column9->Risikokostensatz_Fix_(in_%),
    #                                Column10->Risikokostensatz_Variabel_(in_%),
    #                                Column11->Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%))
    df = df.rename(columns={
        "Column7": "Laufzeit_Von_(in_Tagen)",
        "Column8": "Laufzeit_Bis_(in_Tagen)",
        "Column9": "Risikokostensatz_Fix_(in_%)",
        "Column10": "Risikokostensatz_Variabel_(in_%)",
        "Column11": "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)",
    })

    # = Table.ReplaceValue(..., "-2"->"" in "Risikokundengruppe")
    if "Risikokundengruppe" in df.columns:
        df["Risikokundengruppe"] = df["Risikokundengruppe"].astype(str).str.replace("-2", "", regex=False)

    # = Table.SelectRows(..., Rating_od_wNote not in {"-1,0","-1.0","-2,0","-2.0"})
    drop_bad = {"-1,0", "-1.0", "-2,0", "-2.0"}
    if "Rating_od_wNote" in df.columns:
        df = df[~df["Rating_od_wNote"].astype(str).isin(drop_bad)]

    # = Table.RemoveColumns(..., {"Rating_od_wNote_Original","Rating_od_wNote_Hilfsspalte"})
    df = df.drop(columns=[c for c in ["Rating_od_wNote_Original", "Rating_od_wNote_Hilfsspalte"] if c in df.columns])

    # = Table.ReorderColumns(...) final order
    final_order = [
        "BLZ", "Rating_Kategorie", "Rating_od_wNote", "Forderungsklasse", "Risikokundengruppe",
        "Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)", "Risikokostensatz_Fix_(in_%)",
        "Risikokostensatz_Variabel_(in_%)", "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"
    ]
    cols_existing = [c for c in final_order if c in df.columns]
    df = df[cols_existing]

    # Cast some obvious numeric columns
    for col in ["BLZ", "Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").astype("Int64")

    for col in ["Risikokostensatz_Fix_(in_%)", "Risikokostensatz_Variabel_(in_%)", "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"]:
        if col in df.columns:
            # Values may be in either comma or dot format already; normalize to dot to parse, keep numeric
            s = df[col].astype(str).str.replace(",", ".", regex=False)
            df[col] = pd.to_numeric(s, errors="coerce")

    return df


def read_rk_basis_primaerbanken(file2: Path) -> pd.DataFrame:
    # = Table.PromoteHeaders(Quelle) -> CSV already has headers
    df = pd.read_csv(file2, dtype=str, na_filter=False)

    # = Table.SelectRows(..., [BLZ] = "34")
    if "BLZ" not in df.columns:
        raise ValueError("Expected column 'BLZ' not found in RK_Basis_Primaerbanken.csv")
    df = df[df["BLZ"] == "34"]

    # = Table.ReorderColumns(...) to exact order
    ordered_cols = [
        "BLZ", "Rating_od_wNote", "Rating_Kategorie", "Forderungsklasse", "Risikokundengruppe",
        "Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)", "Risikokostensatz_Fix_(in_%)",
        "Risikokostensatz_Variabel_(in_%)", "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"
    ]
    missing = [c for c in ordered_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Missing expected columns in RK_Basis_Primaerbanken.csv: {missing}")
    df = df[ordered_cols]

    # = Table.TransformColumnTypes(... BLZ:Int64, risk rates:number, Laufzeit_Von/Bis:Int64)
    for col in ["BLZ", "Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").astype("Int64")

    for col in ["Risikokostensatz_Fix_(in_%)", "Risikokostensatz_Variabel_(in_%)", "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"]:
        # numbers might contain comma or dot; normalize to dot before parsing
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", ".", regex=False), errors="coerce")

    return df


def combine_and_finalize(df1: pd.DataFrame, df2: pd.DataFrame) -> pd.DataFrame:
    # Combine
    df = pd.concat([df1, df2], ignore_index=True)

    # Ensure numeric for operations
    for col in ["BLZ", "Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").astype("Int64")

    # = Table.SelectRows(Quelle, each ([BLZ] <> 55000))
    df = df[df["BLZ"] != 55000]

    # = Table.ReplaceValue(..., 365->366, {"Laufzeit_Bis_(in_Tagen)"})
    if "Laufzeit_Bis_(in_Tagen)" in df.columns:
        df["Laufzeit_Bis_(in_Tagen)"] = df["Laufzeit_Bis_(in_Tagen)"].replace(365, 366)

    # = Table.ReplaceValue(..., 366->367, {"Laufzeit_Von_(in_Tagen)"})
    if "Laufzeit_Von_(in_Tagen)" in df.columns:
        df["Laufzeit_Von_(in_Tagen)"] = df["Laufzeit_Von_(in_Tagen)"].replace(366, 367)

    # = Table.DuplicateColumn(... "Risikokostensatz_Variabel_(in_%)" -> copy)
    # = Table.RemoveColumns(... {"Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"})
    # = Table.RenameColumns(... copy -> "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)")
    if "Risikokostensatz_Variabel_(in_%)" in df.columns:
        df["__copy_var"] = df["Risikokostensatz_Variabel_(in_%)"]
        if "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)" in df.columns:
            df = df.drop(columns=["Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"])
        df = df.rename(columns={"__copy_var": "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"})

    # = Table.SelectRows(..., Rating_od_wNote not in {"-1","-2"})
    if "Rating_od_wNote" in df.columns:
        df = df[~df["Rating_od_wNote"].astype(str).isin({"-1", "-2"})]

    # = Table.ReplaceValue(..., "."->"," in two % columns)
    for col in ["Risikokostensatz_Variabel_(in_%)", "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"]:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: str(x).replace(".", ",") if pd.notna(x) else x)

    # = Table.TransformColumnTypes(... those two back to number)
    # Interpret the comma as decimal separator, store as float
    for col in ["Risikokostensatz_Variabel_(in_%)", "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"]:
        if col in df.columns:
            s = df[col].astype(str).str.replace(",", ".", regex=False)
            df[col] = pd.to_numeric(s, errors="coerce")

    # = Table.AddColumn(..., "Gueltig_Ab", each "")
    df["Gueltig_Ab"] = ""

    # = Table.ReorderColumns(... final order)
    final_order = [
        "BLZ", "Gueltig_Ab", "Rating_Kategorie", "Rating_od_wNote", "Forderungsklasse",
        "Risikokundengruppe", "Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)",
        "Risikokostensatz_Fix_(in_%)", "Risikokostensatz_Variabel_(in_%)",
        "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"
    ]
    # Keep only existing columns in that order
    final_cols = [c for c in final_order if c in df.columns]
    df = df[final_cols]

    return df


def main():
    df1 = read_tab_risiko_icaap(FILE1_PATH)
    df2 = read_rk_basis_primaerbanken(FILE2_PATH)
    df_final = combine_and_finalize(df1, df2)

    # Save with UTF-8 BOM to be Excel-friendly
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    df_final.to_csv(OUTPUT_PATH, index=False, encoding="utf-8-sig")

    print(f"Done. Saved: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
