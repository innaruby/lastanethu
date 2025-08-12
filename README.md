import pandas as pd
import numpy as np
from pathlib import Path
import re

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
    """Read CSV without headers, trying different delimiters/encodings."""
    try:
        df = pd.read_csv(path, header=None, sep=None, engine="python", dtype=str, encoding="utf-8-sig")
    except Exception:
        df = None
    if df is None or df.shape[1] == 1:
        try:
            df = pd.read_csv(path, header=None, sep=";", dtype=str, encoding="utf-8-sig")
        except Exception:
            df = None
    if df is None or df.shape[1] == 1:
        try:
            df = pd.read_csv(path, header=None, sep=None, engine="python", dtype=str, encoding="cp1252")
        except Exception:
            df = None
    if df is None or df.shape[1] == 1:
        df = pd.read_csv(path, header=None, sep=",", dtype=str, encoding="cp1252")
    return df

def robust_read_with_header(path: str) -> pd.DataFrame:
    """Read CSV with headers, trying different delimiters/encodings."""
    try:
        df = pd.read_csv(path, sep=None, engine="python", dtype=str, encoding="utf-8-sig")
    except Exception:
        df = None
    if df is None or df.shape[1] == 1:
        try:
            df = pd.read_csv(path, sep=";", dtype=str, encoding="utf-8-sig")
        except Exception:
            df = None
    if df is None or df.shape[1] == 1:
        try:
            df = pd.read_csv(path, sep=None, engine="python", dtype=str, encoding="cp1252")
        except Exception:
            df = None
    if df is None or df.shape[1] == 1:
        df = pd.read_csv(path, sep=";", dtype=str, encoding="cp1252")
    df.columns = df.columns.astype(str).str.replace("\ufeff", "", regex=False).str.strip()
    return df

def ensure_min_columns(df: pd.DataFrame, min_cols: int, context: str):
    if df.shape[1] < min_cols:
        raise ValueError(f"{context} has only {df.shape[1]} columns, expected >= {min_cols}")

def to_us_number(s: pd.Series) -> pd.Series:
    """Convert strings with comma or dot decimal to float."""
    s = s.astype(str).str.replace(" ", "", regex=False)
    s = s.str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce")

# =========================
# 1) Process Tab_EM_ICAAP.csv
# =========================
df1 = robust_read_no_header(path_tab_em)
df1.columns = [f"Column{i+1}" for i in range(df1.shape[1])]
ensure_min_columns(df1, 11, "Tab_EM_ICAAP.csv")

df1["Column1"] = pd.to_numeric(df1["Column1"], errors="coerce").astype("Int64")
df1["Column2"] = df1["Column2"].astype(str)
df1["Column7"] = pd.to_numeric(df1["Column7"], errors="coerce").astype("Int64")
df1["Column8"] = pd.to_numeric(df1["Column8"], errors="coerce").astype("Int64")
df1["Column11"] = to_us_number(df1["Column11"])
df1["Column9"]  = to_us_number(df1["Column9"])

df1 = df1.rename(columns={
    "Column1": "BLZ",
    "Column2": "Rating_od_wNote",
    "Column3": "Rating_Kategorie",
    "Column4": "Forderungsklasse",
    "Column5": "Risikokundengruppe"
})
df1["Copy of Rating_od_wNote"] = df1["Rating_od_wNote"]
df1 = df1.rename(columns={"Rating_od_wNote": "Rating_od_wNote_Original"})

df1 = df1[[c for c in [
    "BLZ", "Rating_od_wNote_Original", "Copy of Rating_od_wNote", "Rating_Kategorie",
    "Forderungsklasse", "Risikokundengruppe", "Column6", "Column7", "Column8",
    "Column9", "Column10", "Column11"
] if c in df1.columns]]

df1 = df1.rename(columns={"Copy of Rating_od_wNote": "Rating_od_wNote_Hilfsspalte"})
if "Column6" in df1.columns:
    df1 = df1.drop(columns=["Column6"])
df1["Rating_od_wNote_Hilfsspalte"] = df1["Rating_od_wNote_Hilfsspalte"].astype(str).str.replace(".", ",", regex=False)

def choose_rating(row):
    if (str(row.get("Rating_Kategorie", "")) in {"10", "11", "12"}) or (str(row.get("Forderungsklasse", "")) in {"1","2","3","4","5"}):
        return row.get("Rating_od_wNote_Hilfsspalte")
    return row.get("Rating_od_wNote_Original")

df1["Rating_od_wNote"] = df1.apply(choose_rating, axis=1)

df1 = df1.rename(columns={
    "Column7": "Laufzeit_Von_(in_Tagen)",
    "Column8": "Laufzeit_Bis_(in_Tagen)",
    "Column9": "Risikokostensatz_Fix_(in_%)",
    "Column10": "Risikokostensatz_Variabel_(in_%)",
    "Column11": "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"
})
if "Risikokundengruppe" in df1.columns:
    df1["Risikokundengruppe"] = df1["Risikokundengruppe"].replace("-2", "")
df1 = df1[~df1["Rating_od_wNote"].isin(["-1,0", "-1.0", "-2,0", "-2.0"])]
df1 = df1.drop(columns=["Rating_od_wNote_Original", "Rating_od_wNote_Hilfsspalte"], errors="ignore")
if "Risikokundengruppe" in df1.columns:
    df1 = df1.drop(columns=["Risikokundengruppe"])
df1 = df1.rename(columns={
    "Risikokostensatz_Fix_(in_%)": "Eigenkapitalkosten_Fix_(in_%)",
    "Risikokostensatz_Variabel_(in_%)": "Eigenkapitalkosten_Variabel_(in_%)",
    "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)": "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"
})

# =========================
# 2) Process EK_Basis_Primaerbanken.csv
# =========================
df2 = robust_read_with_header(path_ek_basis)
blz_col = next((c for c in df2.columns if c.strip().lower() == "blz"), None)
if blz_col is None:
    raise ValueError(f"No BLZ column found. Found: {df2.columns.tolist()}")
df2 = df2[df2[blz_col].astype(str).str.strip() == "34"]
if blz_col != "BLZ":
    df2 = df2.rename(columns={blz_col: "BLZ"})

for col in ["Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)"]:
    if col in df2.columns:
        df2[col] = pd.to_numeric(df2[col], errors="coerce").astype("Int64")
for col in ["Eigenkapitalkosten_Fix_(in_%)", "Eigenkapitalkosten_Variabel_(in_%)", "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"]:
    if col in df2.columns:
        df2[col] = to_us_number(df2[col])

# =========================
# 3) Combine & final transformations
# =========================
all_cols = sorted(set(df1.columns).union(df2.columns))
df_combined = pd.concat([df1.reindex(columns=all_cols), df2.reindex(columns=all_cols)], ignore_index=True)

if "BLZ" in df_combined.columns:
    mask = ~((pd.to_numeric(df_combined["BLZ"], errors="coerce") == 55000) | (df_combined["BLZ"].astype(str).str.strip() == "55000"))
    df_combined = df_combined[mask]

if "Laufzeit_Bis_(in_Tagen)" in df_combined.columns:
    df_combined["Laufzeit_Bis_(in_Tagen)"] = df_combined["Laufzeit_Bis_(in_Tagen)"].replace(365, 366)
if "Laufzeit_Von_(in_Tagen)" in df_combined.columns:
    df_combined["Laufzeit_Von_(in_Tagen)"] = df_combined["Laufzeit_Von_(in_Tagen)"].replace(366, 367)
if "Rating_Kategorie" in df_combined.columns:
    df_combined = df_combined[df_combined["Rating_Kategorie"].astype(str) != "9"]

if "Eigenkapitalkosten_Variabel_(in_%)" in df_combined.columns:
    df_combined["Eigenkapitalkosten_Variabel_(in_%) - Kopie"] = df_combined["Eigenkapitalkosten_Variabel_(in_%)"]
if "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)" in df_combined.columns:
    df_combined = df_combined.drop(columns=["Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"])
if "Eigenkapitalkosten_Variabel_(in_%) - Kopie" in df_combined.columns:
    df_combined = df_combined.rename(columns={"Eigenkapitalkosten_Variabel_(in_%) - Kopie": "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"})

if "Rating_od_wNote" in df_combined.columns:
    df_combined = df_combined[~df_combined["Rating_od_wNote"].isin(["-1", "-2"])]

# =========================
# 4) Enforce final order & remove header-like rows
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
for c in final_cols:
    if c not in df_combined.columns:
        df_combined[c] = pd.NA
df_final = df_combined.reindex(columns=final_cols)

def is_headerish_row(row, columns, min_hits=3):
    hits = row.astype(str).str.match(r'(?i)^column\d+$', na=False).sum()
    equals_header = sum((str(row[col]).strip().casefold() == col.strip().casefold()) for col in columns)
    return (hits + equals_header) >= min_hits

mask_headerish = df_final.apply(lambda r: is_headerish_row(r, df_final.columns), axis=1)
df_final = df_final[~mask_headerish].reset_index(drop=True)

# =========================
# 5) Save Excel + CSV
# =========================
out_dir = Path(out_base).parent
out_dir.mkdir(parents=True, exist_ok=True)

xlsx_path = out_base + ".xlsx"
with pd.ExcelWriter(xlsx_path, engine="xlsxwriter") as writer:
    df_final.to_excel(writer, sheet_name="EK_Basis_Final", index=False)

csv_path = out_base + ".csv"
df_final.to_csv(csv_path, index=False, sep=";", encoding="utf-8-sig")

print("✅ Saved:")
print("  • Excel:", xlsx_path)
print("  • CSV  :", csv_path)
