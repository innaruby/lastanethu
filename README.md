import pandas as pd
import os

# ----------------------------
# File paths
# ----------------------------
file1_path = r"U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien\Tab_EM_ICAAP.csv"
file2_path = r"U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien\Primär-bzw Raiffeisenbanken\EK_Basis_Primaerbanken.csv"
output_path = r"U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Upload_Dateien_Vorkalk\EK_Basis_Final.csv"

# ---------- utility: read CSV with flexible delimiter and no headers ----------
def read_csv_flexible(path, header=None, min_cols=None):
    """
    Read a csv trying to sniff the delimiter. Guarantee at least `min_cols` columns by padding.
    """
    # Try automatic delimiter detection first
    df = pd.read_csv(path, header=header, dtype=str, sep=None, engine="python", encoding="utf-8-sig")
    # If we got only 1 column, try common alternates
    if df.shape[1] == 1:
        for sep_try in [";", "\t", "|"]:
            df_try = pd.read_csv(path, header=header, dtype=str, sep=sep_try, encoding="utf-8-sig")
            if df_try.shape[1] > 1:
                df = df_try
                break

    # Assign dummy names if no headers
    if header is None:
        df.columns = [f"Column{i+1}" for i in range(df.shape[1])]

    # Ensure minimum number of columns by padding empty ones
    if min_cols is not None and df.shape[1] < min_cols:
        for i in range(df.shape[1]+1, min_cols+1):
            df[f"Column{i}"] = pd.NA

    return df

# ----------------------------
# Process first file (Tab_EM_ICAAP.csv) — file has NO headers
# ----------------------------
df1 = read_csv_flexible(file1_path, header=None, min_cols=11)

# --- types: mimic Power Query steps ---
# Column1,7,8 as integers if possible
for col in ["Column1", "Column7", "Column8"]:
    df1[col] = pd.to_numeric(df1[col], errors="coerce").astype("Int64")

# Column2 text
if "Column2" not in df1.columns:
    df1["Column2"] = pd.NA
df1["Column2"] = df1["Column2"].astype(str)

# Column9 and Column11 as numbers (en-US decimal)
for col in ["Column9", "Column11"]:
    # Replace commas with dots just in case, then parse
    df1[col] = pd.to_numeric(
        df1[col].astype(str).str.replace(",", ".", regex=False),
        errors="coerce"
    )

# Rename to business names
df1 = df1.rename(columns={
    "Column1": "BLZ",
    "Column2": "Rating_od_wNote",
    "Column3": "Rating_Kategorie",
    "Column4": "Forderungsklasse",
    "Column5": "Risikokundengruppe"
})

# Duplicate + rename to create helper/original
df1["Copy of Rating_od_wNote"] = df1["Rating_od_wNote"]
df1 = df1.rename(columns={"Rating_od_wNote": "Rating_od_wNote_Original"})

# Make sure Column6..Column11 exist before reordering (pad if needed)
for col in [f"Column{i}" for i in range(6, 12)]:
    if col not in df1.columns:
        df1[col] = pd.NA

df1 = df1[[
    "BLZ", "Rating_od_wNote_Original", "Copy of Rating_od_wNote", "Rating_Kategorie",
    "Forderungsklasse", "Risikokundengruppe", "Column6", "Column7", "Column8",
    "Column9", "Column10", "Column11"
]]

# Rename copy to helper
df1 = df1.rename(columns={"Copy of Rating_od_wNote": "Rating_od_wNote_Hilfsspalte"})

# Remove Column6 if present
if "Column6" in df1.columns:
    df1 = df1.drop(columns=["Column6"])

# Replace "." with "," in helper text (keep as text)
df1["Rating_od_wNote_Hilfsspalte"] = df1["Rating_od_wNote_Hilfsspalte"].astype(str).str.replace(".", ",", regex=False)

# Compute final Rating_od_wNote
def choose_rating(row):
    if str(row.get("Rating_Kategorie")) in {"10", "11", "12"} or str(row.get("Forderungsklasse")) in {"1", "2", "3", "4", "5"}:
        return row.get("Rating_od_wNote_Hilfsspalte")
    return row.get("Rating_od_wNote_Original")

df1["Rating_od_wNote"] = df1.apply(choose_rating, axis=1)

# Reorder
df1 = df1[[
    "BLZ", "Rating_od_wNote", "Rating_od_wNote_Original", "Rating_od_wNote_Hilfsspalte",
    "Rating_Kategorie", "Forderungsklasse", "Risikokundengruppe", "Column7",
    "Column8", "Column9", "Column10", "Column11"
]]

# Rename technicals to final names
df1 = df1.rename(columns={
    "Column7": "Laufzeit_Von_(in_Tagen)",
    "Column8": "Laufzeit_Bis_(in_Tagen)",
    "Column9": "Risikokostensatz_Fix_(in_%)",
    "Column10": "Risikokostensatz_Variabel_(in_%)",
    "Column11": "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"
})

# Replace "-2" -> "" in Risikokundengruppe
df1["Risikokundengruppe"] = df1["Risikokundengruppe"].replace("-2", "")

# Remove unwanted ratings (treat as strings)
df1 = df1[~df1["Rating_od_wNote"].astype(str).isin(["-1,0", "-1.0", "-2,0", "-2.0"])]

# Remove helper cols
df1 = df1.drop(columns=["Rating_od_wNote_Original", "Rating_od_wNote_Hilfsspalte"])

# Reorder then drop Risikokundengruppe per PQ
df1 = df1[[
    "BLZ", "Rating_Kategorie", "Rating_od_wNote", "Forderungsklasse", "Risikokundengruppe",
    "Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)",
    "Risikokostensatz_Fix_(in_%)", "Risikokostensatz_Variabel_(in_%)",
    "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"
]]
df1 = df1.drop(columns=["Risikokundengruppe"])

# Rename to Eigenkapitalkosten
df1 = df1.rename(columns={
    "Risikokostensatz_Fix_(in_%)": "Eigenkapitalkosten_Fix_(in_%)",
    "Risikokostensatz_Variabel_(in_%)": "Eigenkapitalkosten_Variabel_(in_%)",
    "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)": "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"
})

# ----------------------------
# Process second file (EK_Basis_Primaerbanken.csv) — file HAS headers
# ----------------------------
# Let pandas read headers; also sniff delimiter
df2 = read_csv_flexible(file2_path, header=0)
df2.columns = df2.columns.str.strip()

# Power Query: PromoteHeaders already done by header=0

# Filter BLZ == "34" (accept both string or numeric)
if "BLZ" not in df2.columns:
    raise ValueError("Expected BLZ column in EK_Basis_Primaerbanken.csv, but it was not found.")

df2 = df2[df2["BLZ"].astype(str).str.strip() == "34"]

# Type conversions
int_cols = ["Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)"]
for c in int_cols:
    if c in df2.columns:
        df2[c] = pd.to_numeric(df2[c], errors="coerce").astype("Int64")
    else:
        df2[c] = pd.NA

num_cols = [
    "Eigenkapitalkosten_Fix_(in_%)",
    "Eigenkapitalkosten_Variabel_(in_%)",
    "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"
]
for c in num_cols:
    if c in df2.columns:
        # ensure dot-decimal
        df2[c] = pd.to_numeric(df2[c].astype(str).str.replace(",", ".", regex=False), errors="coerce")
    else:
        df2[c] = pd.NA

# BLZ as Int64 where possible (keep text compatibility later)
df2["BLZ"] = pd.to_numeric(df2["BLZ"], errors="coerce").astype("Int64")

# If df2 lacks some columns present in df1, add them so concat works cleanly
for col in df1.columns:
    if col not in df2.columns:
        df2[col] = pd.NA
# Also ensure df1 has any missing df2 cols (rare)
for col in df2.columns:
    if col not in df1.columns:
        df1[col] = pd.NA

# Align column order
df2 = df2[df1.columns]

# ----------------------------
# Combine
# ----------------------------
df_combined = pd.concat([df1, df2], ignore_index=True)

# Filter BLZ != 55000 (compare as string and numeric safely)
df_combined = df_combined[~df_combined["BLZ"].astype(str).str.fullmatch(r"\s*55000\s*")]

# Replace 365 -> 366 in Laufzeit_Bis_(in_Tagen)
df_combined["Laufzeit_Bis_(in_Tagen)"] = pd.to_numeric(df_combined["Laufzeit_Bis_(in_Tagen)"], errors="coerce").astype("Int64").replace(365, 366)

# Replace 366 -> 367 in Laufzeit_Von_(in_Tagen)
df_combined["Laufzeit_Von_(in_Tagen)"] = pd.to_numeric(df_combined["Laufzeit_Von_(in_Tagen)"], errors="coerce").astype("Int64").replace(366, 367)

# Filter Rating_Kategorie != "9"
df_combined = df_combined[df_combined["Rating_Kategorie"].astype(str).str.strip() != "9"]

# Duplicate variable cost to 'nicht ausgenutzter Rahmen', then drop old 'nicht ausgenutzt'
df_combined["Eigenkapitalkosten_Variabel_(in_%) - Kopie"] = pd.to_numeric(
    df_combined["Eigenkapitalkosten_Variabel_(in_%)"], errors="coerce"
)
if "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)" in df_combined.columns:
    df_combined = df_combined.drop(columns=["Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"])
df_combined = df_combined.rename(columns={
    "Eigenkapitalkosten_Variabel_(in_%) - Kopie": "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"
})

# Final rating filter (remove -1 and -2)
df_combined = df_combined[~df_combined["Rating_od_wNote"].astype(str).isin(["-1", "-2"])]

# ----------------------------
# Save
# ----------------------------
os.makedirs(os.path.dirname(output_path), exist_ok=True)
df_combined.to_csv(output_path, index=False, encoding="utf-8-sig")
print("Processing completed. File saved to:", output_path)
