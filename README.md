



import pandas as pd
from pathlib import Path

FILE1_PATH = Path(r"U:\rlbnas1_rlb_bw_firw_z\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien\Tab_Risiko_ICAAP.csv")  # no header
FILE2_PATH = Path(r"U:\rlbnas1_rlb_bw_firw_z\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien\Primär-bzw Raiffeisenbanken\RK_Basis_Primaerbanken.csv")  # has header
OUTPUT_CSV  = Path(r"U:\rlbnas1_rlb_bw_firw_z\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Upload_Dateien_Vorkalk\RK_Basis_Final.csv")
OUTPUT_XLSX = Path(r"U:\rlbnas1_rlb_bw_firw_z\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Upload_Dateien_Vorkalk\RK_Basis_Final.xlsx")  # optional


DELIMS = [",", ";", "\t", "|"]

def read_csv_flexible(path: Path, has_header: bool, prefer_utf8: bool = True) -> pd.DataFrame:
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
                sep=None,            # let parser detect
                engine="python",
                encoding=enc,
                on_bad_lines="skip", # or "warn" to inspect issues
                skipinitialspace=True,
                na_filter=False,
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
    if "Column9"  in df: df["Column9"]  = pd.to_numeric(df["Column9"],  errors="coerce")

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
        "Column10":"Risikokostensatz_Variabel_(in_%)",
        "Column11":"Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)",
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

    # Keep floats internally; we’ll format at export
    df["Gueltig_Ab"] = ""

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

    OUTPUT_CSV.parent.mkdir(parents=True, exist_ok=True)

    # Save CSV for German/Austrian Excel: semicolon delimiter and comma decimal
    df_final.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig", sep=";", decimal=",")

    # Optional: also save XLSX (always displays as proper columns)
    try:
        df_final.to_excel(OUTPUT_XLSX, index=False)
    except Exception:
        pass

    print(f"Saved CSV:  {OUTPUT_CSV}")
    print(f"Saved XLSX: {OUTPUT_XLSX}")

if __name__ == "__main__":
    main()


this is code 1 . Similar to Code 1 construct code 2 according to the following logic. 
I need python code for the following power querys. At first open the file whose name is Tab_EM_ICAAP.csv          which is in directory 


U: \Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien

Please do the following transformation steps. Please note that in this file in row 1 there are no column names .so please add some dummy names in rows so that further processing gets smoothy.


= Table.TransformColumnTypes(Quelle,{{"Column1", Int64.Type}, {"Column2", type text}, {"Column7", Int64.Type}, {"Column8", Int64.Type}, {"Column9", type text}})

= Table.TransformColumnTypes(#"Geänderter Typ", {{"Column11", type number}, {"Column9", type number}}, "en-US")


= Table.RenameColumns(#"Geänderter Typ Zahl aus Englischem Zellenformat CCF",{{"Column1", "BLZ"}, {"Column2", "Rating_od_wNote"}, {"Column3", "Rating_Kategorie"}, {"Column4", "Forderungsklasse"}, {"Column5", "Risikokundengruppe"}})

= Table.DuplicateColumn(#"Umbenannte Spalten", "Rating_od_wNote", "Copy of Rating_od_wNote")

= Table.RenameColumns(#"Hilfsspalte einfuegen",{{"Rating_od_wNote", "Rating_od_wNote_Original"}})

= Table.ReorderColumns(#"Umbenannte Spalten2",{"BLZ", "Rating_od_wNote_Original", "Copy of Rating_od_wNote", "Rating_Kategorie", "Forderungsklasse", "Risikokundengruppe", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11"})

= Table.RenameColumns(#"Neu angeordnete Spalten",{{"Copy of Rating_od_wNote", "Rating_od_wNote_Hilfsspalte"}})

= Table.RemoveColumns(#"Umbenannte Spalten3",{"Column6"})

= Table.ReplaceValue(#"Entfernte Spalten",".",",",Replacer.ReplaceText,{"Rating_od_wNote_Hilfsspalte"})

= Table.AddColumn(#"Punkt Durch Komma in Hilfsspalte", "Rating_od_wNote", each if ([Rating_Kategorie]="10" or [Rating_Kategorie]="11" or [Rating_Kategorie]="12" or [Forderungsklasse] ="1" or [Forderungsklasse] ="2" or [Forderungsklasse] ="3" or [Forderungsklasse] ="4" or [Forderungsklasse] ="5") then [Rating_od_wNote_Hilfsspalte] else [Rating_od_wNote_Original])

= Table.ReorderColumns(#"Bei Ratingkategorie 10 bis 12 Werte aus Hilfsspalte",{"BLZ", "Rating_od_wNote", "Rating_od_wNote_Original", "Rating_od_wNote_Hilfsspalte", "Rating_Kategorie", "Forderungsklasse", "Risikokundengruppe", "Column7", "Column8", "Column9", "Column10", "Column11"})

= Table.RenameColumns(#"Neue Ratingspalte nach vorne schieben",{{"Column7", "Laufzeit_Von_(in_Tagen)"}, {"Column8", "Laufzeit_Bis_(in_Tagen)"}, {"Column9", "Risikokostensatz_Fix_(in_%)"}, {"Column10", "Risikokostensatz_Variabel_(in_%)"}, {"Column11", "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"}})


= Table.ReplaceValue(#"Spalten umbenennen","-2","",Replacer.ReplaceText,{"Risikokundengruppe"})

= Table.SelectRows(#"Ersetzter Wert", each ([Rating_od_wNote] <> "-1,0" and [Rating_od_wNote] <> "-1.0" and [Rating_od_wNote] <> "-2,0" and [Rating_od_wNote] <> "-2.0"))

= Table.RemoveColumns(#"Unnoetige Ratings entfernen",{"Rating_od_wNote_Original", "Rating_od_wNote_Hilfsspalte"})

= Table.ReorderColumns(#"Hilfsspalten fürs Rating löschen",{"BLZ", "Rating_Kategorie", "Rating_od_wNote", "Forderungsklasse", "Risikokundengruppe", "Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)", "Risikokostensatz_Fix_(in_%)", "Risikokostensatz_Variabel_(in_%)", "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"})

= Table.RemoveColumns(#"Kategorie vor Rating",{"Risikokundengruppe"})


= Table.RenameColumns(#"Entfernte Spalten1",{{"Risikokostensatz_Fix_(in_%)", "Eigenkapitalkosten_Fix_(in_%)"}, {"Risikokostensatz_Variabel_(in_%)", "Eigenkapitalkosten_Variabel_(in_%)"}, {"Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)", "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"}})


As next step open the file with name EK_Basis_Primaerbanken.csv   in the following directory

U:\ Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien\Primär-bzw Raiffeisenbanken
Please do the following transformation steps. Please note that in this file in row 1 there are  column names .


= Table.PromoteHeaders(Quelle)


= Table.SelectRows(#"Erste Zeile als Header", each ([BLZ] = "34"))


= Table.TransformColumnTypes(#"Gefilterte Zeilen",{{"Laufzeit_Von_(in_Tagen)", Int64.Type}, {"Laufzeit_Bis_(in_Tagen)", Int64.Type}, {"Eigenkapitalkosten_Fix_(in_%)", type number}, {"Eigenkapitalkosten_Variabel_(in_%)", type number}, {"Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)", type number}, {"BLZ", Int64.Type}})


Then as next step, the processed two files , I need to combine and on to the combined file I want to perform following transformations

= Table.SelectRows(Quelle, each ([BLZ] <> 55000))

= Table.ReplaceValue(#"nur 34 und 34000",365,366,Replacer.ReplaceValue,{"Laufzeit_Bis_(in_Tagen)"})

= Table.ReplaceValue(#"365_durch 366 ersetzen",366,367,Replacer.ReplaceValue,{"Laufzeit_Von_(in_Tagen)"})

= Table.SelectRows(#"Ersetzter Wert", each ([Rating_Kategorie] <> "9"))

= Table.DuplicateColumn(#"Ratingkategorie 9 wegfiltern", "Eigenkapitalkosten_Variabel_(in_%)", "Eigenkapitalkosten_Variabel_(in_%) - Kopie")

= Table.RemoveColumns(#"Duplizierte Spalte",{"Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"})

= Table.RenameColumns(#"Entfernte Spalten",{{"Eigenkapitalkosten_Variabel_(in_%) - Kopie", "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"}})

= Table.SelectRows(NaRahmenSatz, each ([Rating_od_wNote] <> "-1" and [Rating_od_wNote] <> "-2"))






Finally save the file in name    EK_Basis_Final in the following directory

U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Upload_Dateien_Vorkalk
