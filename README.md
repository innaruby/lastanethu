import pandas as pd
import numpy as np
import os

# === Paths ===
path_tab_em_icaap = r"U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien\Tab_EM_ICAAP.csv"
path_ek_basis_prim = r"U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien\Primär-bzw Raiffeisenbanken\EK_Basis_Primaerbanken.csv"
save_path = r"U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Upload_Dateien_Vorkalk\EK_Basis_Final.csv"

def read_no_header_autosep(path, min_cols=1):
    """Read a CSV that has no header, but unknown delimiter. Return dtype=str DataFrame."""
    # Try common separators first
    for sep in [',', ';', '\t', '|']:
        df_try = pd.read_csv(
            path, header=None, sep=sep, dtype=str, encoding='utf-8-sig',
            engine='python', skip_blank_lines=False
        )
        if df_try.shape[1] >= min_cols:
            return df_try
    # Fallback: let pandas sniff
    df_auto = pd.read_csv(
        path, header=None, sep=None, dtype=str, encoding='utf-8-sig',
        engine='python', skip_blank_lines=False
    )
    return df_auto

def ensure_dummy_headers(df):
    """Assign Column1..ColumnN headers for however many columns are present."""
    df = df.copy()
    df.columns = [f"Column{i+1}" for i in range(df.shape[1])]
    return df

def ensure_columns(df, required):
    """Make sure required columns exist; create missing ones with empty string."""
    df = df.copy()
    for col in required:
        if col not in df.columns:
            df[col] = pd.NA
    return df

# === 1) Tab_EM_ICAAP.csv (no headers) ===
df1_raw = read_no_header_autosep(path_tab_em_icaap, min_cols=1)
df1 = ensure_dummy_headers(df1_raw)

# We will need up to Column11 later; create missing ones safely
df1 = ensure_columns(df1, [f"Column{i}" for i in range(1, 12)])

# Types similar to Power Query
df1["Column1"] = pd.to_numeric(df1["Column1"], errors="coerce").astype("Int64")
df1["Column2"] = df1["Column2"].astype(str)
df1["Column7"] = pd.to_numeric(df1["Column7"], errors="coerce").astype("Int64")
df1["Column8"] = pd.to_numeric(df1["Column8"], errors="coerce").astype("Int64")

# en-US numbers for Column9 and Column11
df1["Column9"]  = pd.to_numeric(df1["Column9"],  errors="coerce")
df1["Column11"] = pd.to_numeric(df1["Column11"], errors="coerce")

# Rename columns per spec
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

# Reorder (tolerant)
order_cols = [
    "BLZ", "Rating_od_wNote_Original", "Copy of Rating_od_wNote", "Rating_Kategorie",
    "Forderungsklasse", "Risikokundengruppe", "Column6", "Column7", "Column8",
    "Column9", "Column10", "Column11"
]
df1 = ensure_columns(df1, order_cols)[order_cols]

# Rename copy column
df1.rename(columns={"Copy of Rating_od_wNote": "Rating_od_wNote_Hilfsspalte"}, inplace=True)

# Remove Column6 (ignore if missing)
df1.drop(columns=["Column6"], errors="ignore", inplace=True)

# Replace dot with comma in helper col (string op; keep original commas if any)
df1["Rating_od_wNote_Hilfsspalte"] = df1["Rating_od_wNote_Hilfsspalte"].astype(str).str.replace(".", ",", regex=False)

# Conditional selection for Rating_od_wNote
cond = (
    df1["Rating_Kategorie"].isin(["10", "11", "12"]) |
    df1["Forderungsklasse"].isin(["1", "2", "3", "4", "5"])
)
df1["Rating_od_wNote"] = np.where(cond, df1["Rating_od_wNote_Hilfsspalte"], df1["Rating_od_wNote_Original"])

# Reorder again (tolerant)
re_cols = [
    "BLZ", "Rating_od_wNote", "Rating_od_wNote_Original", "Rating_od_wNote_Hilfsspalte",
    "Rating_Kategorie", "Forderungsklasse", "Risikokundengruppe",
    "Column7", "Column8", "Column9", "Column10", "Column11"
]
df1 = ensure_columns(df1, re_cols)[re_cols]

# Rename technical columns to business names
df1.rename(columns={
    "Column7": "Laufzeit_Von_(in_Tagen)",
    "Column8": "Laufzeit_Bis_(in_Tagen)",
    "Column9": "Risikokostensatz_Fix_(in_%)",
    "Column10": "Risikokostensatz_Variabel_(in_%)",
    "Column11": "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"
}, inplace=True)

# Replace "-2" in Risikokundengruppe with empty
if "Risikokundengruppe" in df1.columns:
    df1["Risikokundengruppe"] = df1["Risikokundengruppe"].replace("-2", "")

# Filter out unwanted rating text variants
df1 = df1[~df1["Rating_od_wNote"].isin(["-1,0", "-1.0", "-2,0", "-2.0"])]

# Remove helper cols
df1.drop(columns=["Rating_od_wNote_Original", "Rating_od_wNote_Hilfsspalte"], errors="ignore", inplace=True)

# Final reorder, then drop Risikokundengruppe
final_cols = [
    "BLZ", "Rating_Kategorie", "Rating_od_wNote", "Forderungsklasse",
    "Risikokundengruppe", "Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)",
    "Risikokostensatz_Fix_(in_%)", "Risikokostensatz_Variabel_(in_%)",
    "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"
]
df1 = ensure_columns(df1, final_cols)[final_cols]
df1.drop(columns=["Risikokundengruppe"], errors="ignore", inplace=True)

# Rename to Eigenkapitalkosten*
df1.rename(columns={
    "Risikokostensatz_Fix_(in_%)": "Eigenkapitalkosten_Fix_(in_%)",
    "Risikokostensatz_Variabel_(in_%)": "Eigenkapitalkosten_Variabel_(in_%)",
    "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)": "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"
}, inplace=True)

# === 2) EK_Basis_Primaerbanken.csv (headers present; auto-detect sep) ===
df2 = pd.read_csv(
    path_ek_basis_prim, sep=None, engine='python', dtype=str, encoding='utf-8-sig'
)
df2.columns = df2.columns.str.strip()

# Filter BLZ == "34"
if "BLZ" not in df2.columns:
    raise ValueError("Expected column 'BLZ' not found in EK_Basis_Primaerbanken.csv")
df2 = df2[df2["BLZ"].astype(str).str.strip() == "34"]

# Robust numeric conversions (nullable)
for c in ["Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)"]:
    if c in df2.columns:
        df2[c] = pd.to_numeric(df2[c], errors="coerce").astype("Int64")
for c in ["Eigenkapitalkosten_Fix_(in_%)", "Eigenkapitalkosten_Variabel_(in_%)", "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"]:
    if c in df2.columns:
        df2[c] = pd.to_numeric(df2[c], errors="coerce")
df2["BLZ"] = pd.to_numeric(df2["BLZ"], errors="coerce").astype("Int64")

# Ensure df2 has the same essential columns as df1 before concatenation
essential = [
    "BLZ", "Rating_Kategorie", "Rating_od_wNote", "Forderungsklasse",
    "Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)",
    "Eigenkapitalkosten_Fix_(in_%)", "Eigenkapitalkosten_Variabel_(in_%)",
    "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"
]
df2 = ensure_columns(df2, essential)[essential]

# === 3) Combine ===
combined = pd.concat([df1, df2], ignore_index=True, sort=False)

# Filter BLZ != 55000 (BLZ may be Int64 or string; make comparable)
# Convert to Int64 if possible; otherwise compare as string
if combined["BLZ"].dtype.name != "Int64":
    combined["BLZ"] = pd.to_numeric(combined["BLZ"], errors="coerce").astype("Int64")
combined = combined[combined["BLZ"] != 55000]

# Replace values 365->366, then 366->367 (nullable-safe)
for col, old_val, new_val in [
    ("Laufzeit_Bis_(in_Tagen)", 365, 366),
    ("Laufzeit_Von_(in_Tagen)", 366, 367),
]:
    if col in combined.columns:
        combined[col] = combined[col].where(combined[col] != old_val, new_val)

# Remove Rating_Kategorie == "9"
combined["Rating_Kategorie"] = combined["Rating_Kategorie"].astype(str)
combined = combined[combined["Rating_Kategorie"] != "9"]

# Duplicate & replace column (variabel -> nicht_ausgenutzter_Rahmen)
combined["Eigenkapitalkosten_Variabel_(in_%) - Kopie"] = combined["Eigenkapitalkosten_Variabel_(in_%)"]
combined.drop(columns=["Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"], errors="ignore", inplace=True)
combined.rename(columns={
    "Eigenkapitalkosten_Variabel_(in_%) - Kopie": "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"
}, inplace=True)

# Remove rows with Rating_od_wNote in {"-1","-2"}
combined["Rating_od_wNote"] = combined["Rating_od_wNote"].astype(str)
combined = combined[~combined["Rating_od_wNote"].isin(["-1", "-2"])]

# === 4) Save result ===
combined.to_csv(save_path, index=False, encoding="utf-8-sig")
print(f"Final file saved to: {save_path}")
