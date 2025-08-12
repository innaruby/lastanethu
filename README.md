import pandas as pd
import numpy as np

# ==== 1. Read and process Tab_EM_ICAAP.csv ====

# Paths (adjust if necessary, ensure raw strings for backslashes)
path_tab_em = r"U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien\Tab_EM_ICAAP.csv"
path_ek_basis = r"U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien\Primär-bzw Raiffeisenbanken\EK_Basis_Primaerbanken.csv"
path_output = r"U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Upload_Dateien_Vorkalk\EK_Basis_Final.csv"

# Read without headers, assign dummy names
df1 = pd.read_csv(path_tab_em, header=None, dtype=str)
df1.columns = [f"Column{i+1}" for i in range(df1.shape[1])]

# Type conversions
df1["Column1"] = pd.to_numeric(df1["Column1"], errors="coerce").astype("Int64")
df1["Column2"] = df1["Column2"].astype(str)
df1["Column7"] = pd.to_numeric(df1["Column7"], errors="coerce").astype("Int64")
df1["Column8"] = pd.to_numeric(df1["Column8"], errors="coerce").astype("Int64")
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

# Duplicate & rename for Hilfsspalte logic
df1["Copy of Rating_od_wNote"] = df1["Rating_od_wNote"]
df1.rename(columns={"Rating_od_wNote": "Rating_od_wNote_Original"}, inplace=True)

# Reorder columns
df1 = df1[[
    "BLZ", "Rating_od_wNote_Original", "Copy of Rating_od_wNote", "Rating_Kategorie",
    "Forderungsklasse", "Risikokundengruppe", "Column6", "Column7", "Column8",
    "Column9", "Column10", "Column11"
]]

# Rename Hilfsspalte
df1.rename(columns={"Copy of Rating_od_wNote": "Rating_od_wNote_Hilfsspalte"}, inplace=True)

# Remove Column6
df1.drop(columns=["Column6"], inplace=True)

# Replace "." with "," in Hilfsspalte
df1["Rating_od_wNote_Hilfsspalte"] = df1["Rating_od_wNote_Hilfsspalte"].astype(str).str.replace(".", ",", regex=False)

# Add calculated column
df1["Rating_od_wNote"] = df1.apply(
    lambda row: row["Rating_od_wNote_Hilfsspalte"]
    if row["Rating_Kategorie"] in ["10", "11", "12"] or row["Forderungsklasse"] in ["1", "2", "3", "4", "5"]
    else row["Rating_od_wNote_Original"], axis=1
)

# Reorder again
df1 = df1[[
    "BLZ", "Rating_od_wNote", "Rating_od_wNote_Original", "Rating_od_wNote_Hilfsspalte",
    "Rating_Kategorie", "Forderungsklasse", "Risikokundengruppe", "Column7", "Column8",
    "Column9", "Column10", "Column11"
]]

# Rename last columns
df1.rename(columns={
    "Column7": "Laufzeit_Von_(in_Tagen)",
    "Column8": "Laufzeit_Bis_(in_Tagen)",
    "Column9": "Risikokostensatz_Fix_(in_%)",
    "Column10": "Risikokostensatz_Variabel_(in_%)",
    "Column11": "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"
}, inplace=True)

# Replace -2 with "" in Risikokundengruppe
df1["Risikokundengruppe"] = df1["Risikokundengruppe"].replace("-2", "")

# Remove unwanted Rating_od_wNote
df1 = df1[~df1["Rating_od_wNote"].isin(["-1,0", "-1.0", "-2,0", "-2.0"])]

# Drop helper columns
df1.drop(columns=["Rating_od_wNote_Original", "Rating_od_wNote_Hilfsspalte"], inplace=True)

# Final reorder
df1 = df1[[
    "BLZ", "Rating_Kategorie", "Rating_od_wNote", "Forderungsklasse", "Risikokundengruppe",
    "Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)",
    "Risikokostensatz_Fix_(in_%)", "Risikokostensatz_Variabel_(in_%)",
    "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"
]]

# Remove Risikokundengruppe
df1.drop(columns=["Risikokundengruppe"], inplace=True)

# Rename to Eigenkapitalkosten
df1.rename(columns={
    "Risikokostensatz_Fix_(in_%)": "Eigenkapitalkosten_Fix_(in_%)",
    "Risikokostensatz_Variabel_(in_%)": "Eigenkapitalkosten_Variabel_(in_%)",
    "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)": "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"
}, inplace=True)


# ==== 2. Read and process EK_Basis_Primaerbanken.csv ====

df2 = pd.read_csv(path_ek_basis, dtype=str)
df2.columns = df2.columns.str.strip()  # clean header

# Filter BLZ == "34"
df2 = df2[df2["BLZ"] == "34"]

# Convert types
df2["Laufzeit_Von_(in_Tagen)"] = pd.to_numeric(df2["Laufzeit_Von_(in_Tagen)"], errors="coerce").astype("Int64")
df2["Laufzeit_Bis_(in_Tagen)"] = pd.to_numeric(df2["Laufzeit_Bis_(in_Tagen)"], errors="coerce").astype("Int64")
for col in ["Eigenkapitalkosten_Fix_(in_%)", "Eigenkapitalkosten_Variabel_(in_%)", "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"]:
    df2[col] = pd.to_numeric(df2[col], errors="coerce")
df2["BLZ"] = pd.to_numeric(df2["BLZ"], errors="coerce").astype("Int64")


# ==== 3. Combine and apply final transformations ====

df_combined = pd.concat([df1, df2], ignore_index=True)

# Filter out BLZ 55000
df_combined = df_combined[df_combined["BLZ"] != 55000]

# Replace values
df_combined["Laufzeit_Bis_(in_Tagen)"] = df_combined["Laufzeit_Bis_(in_Tagen)"].replace(365, 366)
df_combined["Laufzeit_Von_(in_Tagen)"] = df_combined["Laufzeit_Von_(in_Tagen)"].replace(366, 367)

# Remove Rating_Kategorie == "9"
df_combined = df_combined[df_combined["Rating_Kategorie"] != "9"]

# Duplicate and rename columns
df_combined["Eigenkapitalkosten_Variabel_(in_%) - Kopie"] = df_combined["Eigenkapitalkosten_Variabel_(in_%)"]
df_combined.drop(columns=["Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"], inplace=True)
df_combined.rename(columns={
    "Eigenkapitalkosten_Variabel_(in_%) - Kopie": "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"
}, inplace=True)

# Remove rows where Rating_od_wNote == "-1" or "-2"
df_combined = df_combined[~df_combined["Rating_od_wNote"].isin(["-1", "-2"])]

# ==== 4. Save final file ====
df_combined.to_csv(path_output, index=False, encoding="utf-8-sig")

print(f"✅ Transformation complete. File saved to:\n{path_output}")
