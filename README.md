import pandas as pd
import numpy as np
from pathlib import Path

# =========================
# Paths
# =========================
path_tab_em = r"U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien\Tab_EM_ICAAP.csv"
path_ek_basis = r"U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien\Primär-bzw Raiffeisenbanken\EK_Basis_Primaerbanken.csv"
out_base = r"U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Upload_Dateien_Vorkalk\EK_Basis_Final"

# =========================
# Helpers
# =========================
def robust_read_no_header(path: str) -> pd.DataFrame:
    """
    Read a CSV that has NO header row (row 1 contains data).
    Try delimiter sniffing, then ';', then ','; handle UTF-8 BOM / cp1252.
    """
    # 1) Try sniffing + utf-8-sig
    try:
        df = pd.read_csv(path, header=None, sep=None, engine="python", dtype=str, encoding="utf-8-sig")
    except Exception:
        df = None
    # 2) If one column, try semicolon
    if df is None or df.shape[1] == 1:
        try:
            df = pd.read_csv(path, header=None, sep=";", dtype=str, encoding="utf-8-sig")
        except Exception:
            df = None
    # 3) Fall back to cp1252 + sniffing
    if df is None or df.shape[1] == 1:
        try:
            df = pd.read_csv(path, header=None, sep=None, engine="python", dtype=str, encoding="cp1252")
        except Exception:
            df = None
    # 4) Final fallback to comma + cp1252
    if df is None or df.shape[1] == 1:
        df = pd.read_csv(path, header=None, sep=",", dtype=str, encoding="cp1252")
    return df

def robust_read_with_header(path: str) -> pd.DataFrame:
    """
    Read a CSV that DOES have a header row.
    Try delimiter sniffing, then ';', handle BOM and cp1252 fallback. Trim headers.
    """
    # 1) Try sniffing + utf-8-sig
    try:
        df = pd.read_csv(path, sep=None, engine="python", dtype=str, encoding="utf-8-sig")
    except Exception:
        df = None
    # 2) If one column, try semicolon
    if df is None or df.shape[1] == 1:
        try:
            df = pd.read_csv(path, sep=";", dtype=str, encoding="utf-8-sig")
        except Exception:
            df = None
    # 3) Fall back to cp1252 + sniffing
    if df is None or df.shape[1] == 1:
        try:
            df = pd.read_csv(path, sep=None, engine="python", dtype=str, encoding="cp1252")
        except Exception:
            df = None
    # 4) Final fallback to semicolon + cp1252
    if df is None or df.shape[1] == 1:
        df = pd.read_csv(path, sep=";", dtype=str, encoding="cp1252")

    # Clean header names
    df.columns = (
        df.columns.astype(str)
          .str.replace("\ufeff", "", regex=False)  # remove BOM if present
          .str.strip()
    )
    return df

def ensure_min_columns(df: pd.DataFrame, min_cols: int, context: str):
    if df.shape[1] < min_cols:
        raise ValueError(
            f"{context} has only {df.shape[1]} columns after parsing; expected at least {min_cols}. "
            "Check the delimiter/encoding or open the file to verify its structure."
        )

def to_us_number(s: pd.Series) -> pd.Series:
    """
    Normalize numbers written with either comma or dot decimals into Python float.
    Keeps NaN on errors.
    """
    s = s.astype(str).str.replace(" ", "", regex=False)
    s = s.str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce")

# =========================
# 1) Process Tab_EM_ICAAP.csv (NO headers; add dummy names)
# =========================
df1 = robust_read_no_header(path_tab_em)
df1.columns = [f"Column{i+1}" for i in range(df1.shape[1])]
ensure_min_columns(df1, 11, "Tab_EM_ICAAP.csv")

# TransformColumnTypes
df1["Column1"] = pd.to_numeric(df1["Column1"], errors="coerce").astype("Int64")
df1["Column2"] = df1["Column2"].astype(str)
df1["Column7"] = pd.to_numeric(df1["Column7"], errors="coerce").astype("Int64")
df1["Column8"] = pd.to_numeric(df1["Column8"], errors="coerce").astype("Int64")
df1["Column11"] = to_us_number(df1["Column11"])
df1["Column9"]  = to_us_number(df1["Column9"])

# RenameColumns
df1 = df1.rename(columns={
    "Column1": "BLZ",
    "Column2": "Rating_od_wNote",
    "Column3": "Rating_Kategorie",
    "Column4": "Forderungsklasse",
    "Column5": "Risikokundengruppe"
})

# Duplicate & rename for Hilfsspalte
df1["Copy of Rating_od_wNote"] = df1["Rating_od_wNote"]
df1 = df1.rename(columns={"Rating_od_wNote": "Rating_od_wNote_Original"})

# Reorder (guard missing columns)
cols_reorder = [
    "BLZ", "Rating_od_wNote_Original", "Copy of Rating_od_wNote", "Rating_Kategorie",
    "Forderungsklasse", "Risikokundengruppe", "Column6", "Column7", "Column8",
    "Column9", "Column10", "Column11"
]
df1 = df1[[c for c in cols_reorder if c in df1.columns]]

# Rename Hilfsspalte
df1 = df1.rename(columns={"Copy of Rating_od_wNote": "Rating_od_wNote_Hilfsspalte"})

# Remove Column6
if "Column6" in df1.columns:
    df1 = df1.drop(columns=["Column6"])

# Replace "." -> "," in Hilfsspalte
df1["Rating_od_wNote_Hilfsspalte"] = df1["Rating_od_wNote_Hilfsspalte"].astype(str).str.replace(".", ",", regex=False)

# Add calculated column
def choose_rating(row):
    if (str(row.get("Rating_Kategorie", "")) in {"10", "11", "12"}) or (str(row.get("Forderungsklasse", "")) in {"1","2","3","4","5"}):
        return row.get("Rating_od_wNote_Hilfsspalte")
    return row.get("Rating_od_wNote_Original")

df1["Rating_od_wNote"] = df1.apply(choose_rating, axis=1)

# Reorder again
cols_reorder2 = [
    "BLZ", "Rating_od_wNote", "Rating_od_wNote_Original", "Rating_od_wNote_Hilfsspalte",
    "Rating_Kategorie", "Forderungsklasse", "Risikokundengruppe", "Column7", "Column8",
    "Column9", "Column10", "Column11"
]
df1 = df1[[c for c in cols_reorder2 if c in df1.columns]]

# Rename tail columns
rename_tail = {}
if "Column7" in df1.columns:  rename_tail["Column7"]  = "Laufzeit_Von_(in_Tagen)"
if "Column8" in df1.columns:  rename_tail["Column8"]  = "Laufzeit_Bis_(in_Tagen)"
if "Column9" in df1.columns:  rename_tail["Column9"]  = "Risikokostensatz_Fix_(in_%)"
if "Column10" in df1.columns: rename_tail["Column10"] = "Risikokostensatz_Variabel_(in_%)"
if "Column11" in df1.columns: rename_tail["Column11"] = "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"
df1 = df1.rename(columns=rename_tail)

# Replace -2 -> "" in Risikokundengruppe
if "Risikokundengruppe" in df1.columns:
    df1["Risikokundengruppe"] = df1["Risikokundengruppe"].replace("-2", "")

# Filter out unwanted Rating_od_wNote
df1 = df1[~df1["Rating_od_wNote"].isin(["-1,0", "-1.0", "-2,0", "-2.0"])]

# Drop helper columns
for col in ["Rating_od_wNote_Original", "Rating_od_wNote_Hilfsspalte"]:
    if col in df1.columns:
        df1 = df1.drop(columns=[col])

# Remove Risikokundengruppe
if "Risikokundengruppe" in df1.columns:
    df1 = df1.drop(columns=["Risikokundengruppe"])

# Rename Risikokostensatz* -> Eigenkapitalkosten*
df1 = df1.rename(columns={
    "Risikokostensatz_Fix_(in_%)": "Eigenkapitalkosten_Fix_(in_%)",
    "Risikokostensatz_Variabel_(in_%)": "Eigenkapitalkosten_Variabel_(in_%)",
    "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)": "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"
})

# Ensure numeric where needed
for c in ["Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)"]:
    if c in df1.columns:
        df1[c] = pd.to_numeric(df1[c], errors="coerce").astype("Int64")
for c in ["Eigenkapitalkosten_Fix_(in_%)", "Eigenkapitalkosten_Variabel_(in_%)", "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"]:
    if c in df1.columns:
        df1[c] = to_us_number(df1[c])

# =========================
# 2) Process EK_Basis_Primaerbanken.csv (HAS headers)
# =========================
df2 = robust_read_with_header(path_ek_basis)

# Find BLZ column safely (case-insensitive)
blz_col = next((c for c in df2.columns if c.strip().lower() == "blz"), None)
if blz_col is None:
    raise ValueError(f"Could not find a 'BLZ' column in EK_Basis_Primaerbanken.csv. Found columns: {list(df2.columns)}")

# Filter BLZ == "34" (string compare, trimmed)
df2 = df2[df2[blz_col].astype(str).str.strip() == "34"]

# Normalize BLZ column name to 'BLZ'
if blz_col != "BLZ":
    df2 = df2.rename(columns={blz_col: "BLZ"})

# Convert types (Power Query)
if "Laufzeit_Von_(in_Tagen)" in df2.columns:
    df2["Laufzeit_Von_(in_Tagen)"] = pd.to_numeric(df2["Laufzeit_Von_(in_Tagen)"], errors="coerce").astype("Int64")
if "Laufzeit_Bis_(in_Tagen)" in df2.columns:
    df2["Laufzeit_Bis_(in_Tagen)"] = pd.to_numeric(df2["Laufzeit_Bis_(in_Tagen)"], errors="coerce").astype("Int64")

for col in ["Eigenkapitalkosten_Fix_(in_%)", "Eigenkapitalkosten_Variabel_(in_%)", "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"]:
    if col in df2.columns:
        df2[col] = to_us_number(df2[col])

# =========================
# 3) Combine and final transformations
# =========================
# Align columns (outer union), then concat
all_cols = sorted(set(df1.columns).union(df2.columns))
df1a = df1.reindex(columns=all_cols)
df2a = df2.reindex(columns=all_cols)
df_combined = pd.concat([df1a, df2a], ignore_index=True)

# Filter out BLZ 55000
if "BLZ" in df_combined.columns:
    blz_numeric = pd.to_numeric(df_combined["BLZ"], errors="coerce")
    mask_not_55000 = ~((blz_numeric == 55000) | (df_combined["BLZ"].astype(str).str.strip() == "55000"))
    df_combined = df_combined[mask_not_55000]

# 365 -> 366 for Laufzeit_Bis
if "Laufzeit_Bis_(in_Tagen)" in df_combined.columns:
    df_combined["Laufzeit_Bis_(in_Tagen)"] = pd.to_numeric(df_combined["Laufzeit_Bis_(in_Tagen)"], errors="coerce").astype("Int64")
    df_combined["Laufzeit_Bis_(in_Tagen)"] = df_combined["Laufzeit_Bis_(in_Tagen)"].replace(365, 366)

# 366 -> 367 for Laufzeit_Von
if "Laufzeit_Von_(in_Tagen)" in df_combined.columns:
    df_combined["Laufzeit_Von_(in_Tagen)"] = pd.to_numeric(df_combined["Laufzeit_Von_(in_Tagen)"], errors="coerce").astype("Int64")
    df_combined["Laufzeit_Von_(in_Tagen)"] = df_combined["Laufzeit_Von_(in_Tagen)"].replace(366, 367)

# Remove Rating_Kategorie == "9"
if "Rating_Kategorie" in df_combined.columns:
    df_combined = df_combined[df_combined["Rating_Kategorie"].astype(str) != "9"]

# Duplicate/rename cost columns
if "Eigenkapitalkosten_Variabel_(in_%)" in df_combined.columns:
    df_combined["Eigenkapitalkosten_Variabel_(in_%) - Kopie"] = df_combined["Eigenkapitalkosten_Variabel_(in_%)"]
if "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)" in df_combined.columns:
    df_combined = df_combined.drop(columns=["Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"])
if "Eigenkapitalkosten_Variabel_(in_%) - Kopie" in df_combined.columns:
    df_combined = df_combined.rename(columns={
        "Eigenkapitalkosten_Variabel_(in_%) - Kopie": "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"
    })

# Remove rows where Rating_od_wNote exactly "-1" or "-2"
if "Rating_od_wNote" in df_combined.columns:
    df_combined = df_combined[~df_combined["Rating_od_wNote"].isin(["-1", "-2"])]

# =========================
# 4) Enforce final column order and save (Excel + CSV with semicolon)
# =========================
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
# Ensure all exist
for c in final_cols:
    if c not in df_combined.columns:
        df_combined[c] = pd.NA
df_final = df_combined.reindex(columns=final_cols)

# Create output dir
out_dir = Path(out_base).parent
out_dir.mkdir(parents=True, exist_ok=True)

# Excel (recommended to avoid "all data in column A")
xlsx_path = out_base + ".xlsx"
with pd.ExcelWriter(xlsx_path, engine="xlsxwriter") as writer:
    df_final.to_excel(writer, sheet_name="EK_Basis_Final", index=False)

# CSV with semicolon (opens correctly in German Excel as columns)
csv_path = out_base + ".csv"
df_final.to_csv(csv_path, index=False, sep=";", encoding="utf-8-sig")

print("✅ Saved:")
print("  • Excel:", xlsx_path)
print("  • CSV  :", csv_path)
