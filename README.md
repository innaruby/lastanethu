def read_tab_em(path):
    # Try automatic delimiter detection
    try:
        df = pd.read_csv(path, header=None, sep=None, engine="python", dtype=str)
    except Exception:
        df = None

    # If still only 1 column, try semicolon explicitly
    if df is None or df.shape[1] == 1:
        df = pd.read_csv(path, header=None, sep=";", dtype=str)

    # If still only 1 column, fall back to comma
    if df.shape[1] == 1:
        df = pd.read_csv(path, header=None, sep=",", dtype=str)

    return df

df1 = read_tab_em(path_tab_em)
df1.columns = [f"Column{i+1}" for i in range(df1.shape[1])]
