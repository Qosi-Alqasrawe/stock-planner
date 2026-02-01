# src/io.py  (كامل)
import pandas as pd

def read_excel_any(file) -> pd.DataFrame:
    name = getattr(file, "name", "").lower()
    if name.endswith(".xls"):
        return pd.read_excel(file, engine="xlrd")
    return pd.read_excel(file, engine="openpyxl")

def normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def _digits_only(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return "".join(ch for ch in s if ch.isdigit())

def format_item_no(x, pad_len: int = 11) -> str:
    d = _digits_only(x)
    if not d:
        return ""
    return d.zfill(pad_len)

def _to_key_last10(x) -> str:
    d = _digits_only(x)
    return d[-10:] if d else ""

def merge_items_stock(stock_df, items_df, stock_item_col, items_item_col, keep_items_cols=None):
    s = stock_df.copy()
    i = items_df.copy()

    s["_key"] = s[stock_item_col].apply(_to_key_last10)
    i["_key"] = i[items_item_col].apply(_to_key_last10)

    if keep_items_cols:
        keep = ["_key"] + [c for c in keep_items_cols if c in i.columns]
        i = i[keep]
    else:
        i = i[["_key"]]

    out = s.merge(i, on="_key", how="left")
    out.drop(columns=["_key"], inplace=True)
    return out
