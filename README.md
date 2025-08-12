# -*- coding: utf-8 -*-
"""


from pathlib import Path
import pandas as pd



PATH_TAB_EM_ICAAP = Path(
    r"U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien\Tab_EM_ICAAP.csv"
)
PATH_EK_BASIS_PRIM = Path(
    r"U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien\Primär-bzw Raiffeisenbanken\EK_Basis_Primaerbanken.csv"
)
OUTPUT_DIR = Path(
    r"U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Upload_Dateien_Vorkalk"
)
OUTPUT_FILE = OUTPUT_DIR / "EK_Basis_Final.csv"


def process_tab_em_icaap(path_tab: Path) -> pd.DataFrame:
    """Replicates the Power Query steps for Tab_EM_ICAAP.csv (no headers)."""

    # Read with no headers -> create dummy Column1..ColumnN
    df = pd.read_csv(path_tab, header=None, dtype=str, encoding="utf-8", engine="python")
    df.columns = [f"Column{i+1}" for i in range(df.shape[1])]

    # Table.TransformColumnTypes(Quelle, {Column1:int, Column2:text, Column7:int, Column8:int, Column9:text})
    # We keep everything as str for safety, then coerce needed columns.
    for col in ["Column1", "Column7", "Column8"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").astype("Int64")

    if "Column2" in df.columns:
        df["Column2"] = df["Column2"].astype(str)

    # Column9 initially treated as text in first step; second step converts Column11, Column9 to number (en-US)
    # -> Make sure '.' as decimal separator works.
    for col in ["Column11", "Column9"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Rename columns
    rename_map = {
        "Column1": "BLZ",
        "Column2": "Rating_od_wNote",
        "Column3": "Rating_Kategorie",
        "Column4": "Forderungsklasse",
        "Column5": "Risikokundengruppe",
    }
    df = df.rename(columns=rename_map)

    # Duplicate & rename to create Original/Hilfsspalte
    if "Rating_od_wNote" in df.columns:
        df["Copy of Rating_od_wNote"] = df["Rating_od_wNote"]
        df = df.rename(columns={"Rating_od_wNote": "Rating_od_wNote_Original"})

    # Reorder to match PQ step (tolerant if some columns are missing)
    order_cols = [
        "BLZ",
        "Rating_od_wNote_Original",
        "Copy of Rating_od_wNote",
        "Rating_Kategorie",
        "Forderungsklasse",
        "Risikokundengruppe",
        "Column6",
        "Column7",
        "Column8",
        "Column9",
        "Column10",
        "Column11",
    ]
    df = df.reindex(columns=[c for c in order_cols if c in df.columns])

    # Rename duplicated to Hilfsspalte
    df = df.rename(columns={"Copy of Rating_od_wNote": "Rating_od_wNote_Hilfsspalte"})

    # Remove Column6 if exists
    if "Column6" in df.columns:
        df = df.drop(columns=["Column6"])

    # Replace "." -> "," in Hilfsspalte (string operation)
    if "Rating_od_wNote_Hilfsspalte" in df.columns:
        df["Rating_od_wNote_Hilfsspalte"] = (
            df["Rating_od_wNote_Hilfsspalte"].astype(str).str.replace(".", ",", regex=False)
        )

    # Add new Rating_od_wNote based on conditions
    def choose_rating(row):
        rk = str(row.get("Rating_Kategorie", ""))
        fk = str(row.get("Forderungsklasse", ""))
        if rk in {"10", "11", "12"} or fk in {"1", "2", "3", "4", "5"}:
            return row.get("Rating_od_wNote_Hilfsspalte")
        return row.get("Rating_od_wNote_Original")

    df["Rating_od_wNote"] = df.apply(choose_rating, axis=1)

    # Reorder again
    order_cols2 = [
        "BLZ",
        "Rating_od_wNote",
        "Rating_od_wNote_Original",
        "Rating_od_wNote_Hilfsspalte",
        "Rating_Kategorie",
        "Forderungsklasse",
        "Risikokundengruppe",
        "Column7",
        "Column8",
        "Column9",
        "Column10",
        "Column11",
    ]
    df = df.reindex(columns=[c for c in order_cols2 if c in df.columns])

    # Rename Column7..11
    rename_costs = {
        "Column7": "Laufzeit_Von_(in_Tagen)",
        "Column8": "Laufzeit_Bis_(in_Tagen)",
        "Column9": "Risikokostensatz_Fix_(in_%)",
        "Column10": "Risikokostensatz_Variabel_(in_%)",
        "Column11": "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)",
    }
    df = df.rename(columns=rename_costs)

    # Replace "-2" -> "" in Risikokundengruppe
    if "Risikokundengruppe" in df.columns:
        df["Risikokundengruppe"] = df["Risikokundengruppe"].astype(str).str.replace("-2", "", regex=False)

    # Remove rows where Rating_od_wNote is one of "-1,0", "-1.0", "-2,0", "-2.0"
    if "Rating_od_wNote" in df.columns:
        bad_vals = {"-1,0", "-1.0", "-2,0", "-2.0"}
        df = df[~df["Rating_od_wNote"].astype(str).isin(bad_vals)]

    # Drop helper columns
    for c in ["Rating_od_wNote_Original", "Rating_od_wNote_Hilfsspalte"]:
        if c in df.columns:
            df = df.drop(columns=[c])

    # Reorder
    final_order_1 = [
        "BLZ",
        "Rating_Kategorie",
        "Rating_od_wNote",
        "Forderungsklasse",
        "Risikokundengruppe",
        "Laufzeit_Von_(in_Tagen)",
        "Laufzeit_Bis_(in_Tagen)",
        "Risikokostensatz_Fix_(in_%)",
        "Risikokostensatz_Variabel_(in_%)",
        "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)",
    ]
    df = df.reindex(columns=[c for c in final_order_1 if c in df.columns])

    # Remove Risikokundengruppe
    if "Risikokundengruppe" in df.columns:
        df = df.drop(columns=["Risikokundengruppe"])

    # Rename risk cost columns to eigenkapitalkosten
    df = df.rename(
        columns={
            "Risikokostensatz_Fix_(in_%)": "Eigenkapitalkosten_Fix_(in_%)",
            "Risikokostensatz_Variabel_(in_%)": "Eigenkapitalkosten_Variabel_(in_%)",
            "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)": "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)",
        }
    )

    # Ensure numeric types where sensible (but keep strings where PQ kept strings)
    for col in ["BLZ", "Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").astype("Int64")

    for col in [
        "Eigenkapitalkosten_Fix_(in_%)",
        "Eigenkapitalkosten_Variabel_(in_%)",
        "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)",
    ]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


def process_ek_basis_primaerbanken(path_prim: Path) -> pd.DataFrame:
    """Replicates the Power Query steps for EK_Basis_Primaerbanken.csv (has headers)."""
    df = pd.read_csv(path_prim, dtype=str, encoding="utf-8", engine="python")

    # Table.PromoteHeaders -> already used headers
    # Table.SelectRows(..., [BLZ] = "34")
    if "BLZ" not in df.columns:
        raise ValueError("EK_Basis_Primaerbanken.csv is missing 'BLZ' column.")
    df = df[df["BLZ"].astype(str) == "34"].copy()

    # Convert types:
    # {"Laufzeit_Von_(in_Tagen)": int, "Laufzeit_Bis_(in_Tagen)": int,
    #  "Eigenkapitalkosten_Fix_(in_%)": number, "Eigenkapitalkosten_Variabel_(in_%)": number,
    #  "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)": number, "BLZ": int}
    for col in ["Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)", "BLZ"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").astype("Int64")

    for col in [
        "Eigenkapitalkosten_Fix_(in_%)",
        "Eigenkapitalkosten_Variabel_(in_%)",
        "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)",
    ]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Ensure PQ-common columns exist to align with the first dataframe
    # Some columns like Rating_Kategorie / Rating_od_wNote / Forderungsklasse might be absent in this file.
    for needed in ["Rating_Kategorie", "Rating_od_wNote", "Forderungsklasse"]:
        if needed not in df.columns:
            df[needed] = pd.NA

    # Keep the shared final column set
    final_cols = [
        "BLZ",
        "Rating_Kategorie",
        "Rating_od_wNote",
        "Forderungsklasse",
        "Laufzeit_Von_(in_Tagen)",
        "Laufzeit_Bis_(in_Tagen)",
        "Eigenkapitalkosten_Fix_(in_%)",
        "Eigenkapitalkosten_Variabel_(in_%)",
        "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)",
    ]
    df = df.reindex(columns=[c for c in final_cols if c in df.columns])

    return df


def combine_and_postprocess(df1: pd.DataFrame, df2: pd.DataFrame) -> pd.DataFrame:
    """Combine the two processed dataframes and apply the final transformation steps."""
    combined = pd.concat([df1, df2], ignore_index=True, sort=False)

    # = Table.SelectRows(Quelle, each ([BLZ] <> 55000))
    if "BLZ" in combined.columns:
        combined = combined[combined["BLZ"].astype("Int64") != 55000]

    # = Replace 365 -> 366 in Laufzeit_Bis_(in_Tagen)
    if "Laufzeit_Bis_(in_Tagen)" in combined.columns:
        combined["Laufzeit_Bis_(in_Tagen)"] = combined["Laufzeit_Bis_(in_Tagen)"].replace(365, 366)

    # = Replace 366 -> 367 in Laufzeit_Von_(in_Tagen)
    if "Laufzeit_Von_(in_Tagen)" in combined.columns:
        combined["Laufzeit_Von_(in_Tagen)"] = combined["Laufzeit_Von_(in_Tagen)"].replace(366, 367)

    # = Table.SelectRows(..., each [Rating_Kategorie] <> "9")
    if "Rating_Kategorie" in combined.columns:
        combined = combined[combined["Rating_Kategorie"].astype(str) != "9"]

    # = Duplicate Eigenkapitalkosten_Variabel_(in_%) -> copy, drop _nicht_ausgenutzter_, rename copy to _nicht_ausgenutzter_
    var_col = "Eigenkapitalkosten_Variabel_(in_%)"
    frame_col = "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"
    if var_col in combined.columns:
        combined[f"{var_col} - Kopie"] = combined[var_col]
        if frame_col in combined.columns:
            combined = combined.drop(columns=[frame_col])
        combined = combined.rename(columns={f"{var_col} - Kopie": frame_col})

    # = Table.SelectRows(NaRahmenSatz, each ([Rating_od_wNote] <> "-1" and [Rating_od_wNote] <> "-2"))
    if "Rating_od_wNote" in combined.columns:
        combined = combined[~combined["Rating_od_wNote"].astype(str).isin({"-1", "-2"})]

    # Enforce final column order (keep any extras at the end)
    final_order = [
        "BLZ",
        "Rating_Kategorie",
        "Rating_od_wNote",
        "Forderungsklasse",
        "Laufzeit_Von_(in_Tagen)",
        "Laufzeit_Bis_(in_Tagen)",
        "Eigenkapitalkosten_Fix_(in_%)",
        "Eigenkapitalkosten_Variabel_(in_%)",
        "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)",
    ]
    ordered = [c for c in final_order if c in combined.columns]
    remaining = [c for c in combined.columns if c not in ordered]
    combined = combined[ordered + remaining]

    return combined


def main():
    # Safety: create output dir if missing
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # Process both inputs
    df_icaap = process_tab_em_icaap(PATH_TAB_EM_ICAAP)
    df_prim = process_ek_basis_primaerbanken(PATH_EK_BASIS_PRIM)

    # Combine + postprocess
    final_df = combine_and_postprocess(df_icaap, df_prim)

    # Save CSV (UTF-8 with BOM for Excel friendliness)
    final_df.to_csv(OUTPUT_FILE, index=False, encoding="utf-8-sig")

    print(f"Done. Rows: {len(final_df):,} | Saved to: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
