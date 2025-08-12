import pandas as pd
import numpy as np
import os

# === Paths ===
path_tab_em_icaap = r"U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien\Tab_EM_ICAAP.csv"
path_ek_basis_prim = r"U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien\Primär-bzw Raiffeisenbanken\EK_Basis_Primaerbanken.csv"
save_path = r"U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Upload_Dateien_Vorkalk\EK_Basis_Final.csv"

# === 1. Load Tab_EM_ICAAP.csv (no headers) ===
df1 = pd.read_csv(path_tab_em_icaap, header=None, dtype=str)
df1.columns = [f"Column{i+1}" for i in range(df1.shape[1])]

# Convert types
df1["Column1"] = pd.to_numeric(df1["Column1"], errors="coerce").astype("Int64")
df1["Column2"] = df1["Column2"].astype(str)
df1["Column7"] = pd.to_numeric(df1["Column7"], errors="coerce").astype("Int64")
df1["Column8"] = pd.to_numeric(df1["Column8"], errors="coerce").astype("Int64")

# Column9 and Column11 as numbers (en-US decimal)
df1["Column9"] = pd.to_numeric(df1["Column9"], errors="coerce")
df1["Column11"] = pd.to_numeric(df1["Column11"], errors="coerce")

# Rename columns
df1.rename(columns={
    "Column1": "BLZ",
    "Column2": "Rating_od_wNote",
    "Column3": "Rating_Kategorie",
    "Column4": "Forderungsklasse",
    "Column5": "Risikokundengruppe"
}, inplace=True)

# Duplicate and rename
df1["Copy of Rating_od_wNote"] = df1["Rating_od_wNote"]
df1.rename(columns={"Rating_od_wNote": "Rating_od_wNote_Original"}, inplace=True)

# Reorder columns
order_cols = [
    "BLZ", "Rating_od_wNote_Original", "Copy of Rating_od_wNote", "Rating_Kategorie",
    "Forderungsklasse", "Risikokundengruppe", "Column6", "Column7", "Column8",
    "Column9", "Column10", "Column11"
]
df1 = df1[order_cols]

# Rename copy column
df1.rename(columns={"Copy of Rating_od_wNote": "Rating_od_wNote_Hilfsspalte"}, inplace=True)

# Remove Column6
df1.drop(columns=["Column6"], inplace=True)

# Replace . with , in Hilfsspalte
df1["Rating_od_wNote_Hilfsspalte"] = df1["Rating_od_wNote_Hilfsspalte"].str.replace(".", ",", regex=False)

# Conditional Rating_od_wNote
conditions = (
    df1["Rating_Kategorie"].isin(["10", "11", "12"]) |
    df1["Forderungsklasse"].isin(["1", "2", "3", "4", "5"])
)
df1["Rating_od_wNote"] = np.where(conditions, df1["Rating_od_wNote_Hilfsspalte"], df1["Rating_od_wNote_Original"])

# Reorder again
df1 = df1[[
    "BLZ", "Rating_od_wNote", "Rating_od_wNote_Original", "Rating_od_wNote_Hilfsspalte",
    "Rating_Kategorie", "Forderungsklasse", "Risikokundengruppe",
    "Column7", "Column8", "Column9", "Column10", "Column11"
]]

# Rename remaining columns
df1.rename(columns={
    "Column7": "Laufzeit_Von_(in_Tagen)",
    "Column8": "Laufzeit_Bis_(in_Tagen)",
    "Column9": "Risikokostensatz_Fix_(in_%)",
    "Column10": "Risikokostensatz_Variabel_(in_%)",
    "Column11": "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"
}, inplace=True)

# Replace -2 in Risikokundengruppe
df1["Risikokundengruppe"] = df1["Risikokundengruppe"].replace("-2", "")

# Remove unwanted ratings
df1 = df1[~df1["Rating_od_wNote"].isin(["-1,0", "-1.0", "-2,0", "-2.0"])]

# Remove helper cols
df1.drop(columns=["Rating_od_wNote_Original", "Rating_od_wNote_Hilfsspalte"], inplace=True)

# Reorder & drop Risikokundengruppe
df1 = df1[[
    "BLZ", "Rating_Kategorie", "Rating_od_wNote", "Forderungsklasse",
    "Risikokundengruppe", "Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)",
    "Risikokostensatz_Fix_(in_%)", "Risikokostensatz_Variabel_(in_%)",
    "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"
]]
df1.drop(columns=["Risikokundengruppe"], inplace=True)

# Rename Risikokostensatz columns
df1.rename(columns={
    "Risikokostensatz_Fix_(in_%)": "Eigenkapitalkosten_Fix_(in_%)",
    "Risikokostensatz_Variabel_(in_%)": "Eigenkapitalkosten_Variabel_(in_%)",
    "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)": "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"
}, inplace=True)


# === 2. Load EK_Basis_Primaerbanken.csv ===
df2 = pd.read_csv(path_ek_basis_prim, dtype=str)
df2.columns = df2.columns.str.strip()

# Filter BLZ == "34"
df2 = df2[df2["BLZ"] == "34"]

# Convert types
df2["Laufzeit_Von_(in_Tagen)"] = pd.to_numeric(df2["Laufzeit_Von_(in_Tagen)"], errors="coerce").astype("Int64")
df2["Laufzeit_Bis_(in_Tagen)"] = pd.to_numeric(df2["Laufzeit_Bis_(in_Tagen)"], errors="coerce").astype("Int64")
df2["Eigenkapitalkosten_Fix_(in_%)"] = pd.to_numeric(df2["Eigenkapitalkosten_Fix_(in_%)"], errors="coerce")
df2["Eigenkapitalkosten_Variabel_(in_%)"] = pd.to_numeric(df2["Eigenkapitalkosten_Variabel_(in_%)"], errors="coerce")
df2["Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"] = pd.to_numeric(df2["Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"], errors="coerce")
df2["BLZ"] = pd.to_numeric(df2["BLZ"], errors="coerce").astype("Int64")


# === 3. Combine ===
combined = pd.concat([df1, df2], ignore_index=True)

# Filter BLZ != 55000
combined = combined[combined["BLZ"] != 55000]

# Replace values
combined["Laufzeit_Bis_(in_Tagen)"] = combined["Laufzeit_Bis_(in_Tagen)"].replace(365, 366)
combined["Laufzeit_Von_(in_Tagen)"] = combined["Laufzeit_Von_(in_Tagen)"].replace(366, 367)

# Remove Rating_Kategorie == "9"
combined = combined[combined["Rating_Kategorie"] != "9"]

# Duplicate & replace column
combined["Eigenkapitalkosten_Variabel_(in_%) - Kopie"] = combined["Eigenkapitalkosten_Variabel_(in_%)"]
combined.drop(columns=["Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"], inplace=True)
combined.rename(columns={
    "Eigenkapitalkosten_Variabel_(in_%) - Kopie": "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"
}, inplace=True)

# Remove more unwanted ratings
combined = combined[~combined["Rating_od_wNote"].isin(["-1", "-2"])]

# === 4. Save result ===
combined.to_csv(save_path, index=False, encoding="utf-8-sig")

print(f"Final file saved to: {save_path}")
