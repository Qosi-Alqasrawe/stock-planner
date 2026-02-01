# src/alerts.py
import pandas as pd


# =========================================================
# CUST Alert Exclusions (Daily pickup items)
# These items always have CUST stock = 0 and are supplied daily from MPI,
# so we remove them from Customer Alert by blanking CUST Alert.
# Match is done using digits-only Item No. without leading zeros.
# =========================================================
CUST_ALERT_EXCLUDE_ITEMNOS = {
    "1002010105",
    "1002010107",
    "1002010301",
    "1002010302",
    "1002010303",
    "1002010304",
    "1002010305",
    "1002010308",
    "1002010401",
    "1002010405",
    "1002010410",
    "1002010411",
    "1002010412",
    "1002010413",
    "1002010414",
    "1002020101",
    "1002080101",
    "1002080102",
    "1002080103",
    "1002080104",
    "1002080201",
    "1002080202",
    "1002080203",
    "1002080204",
    "1002080302",
    "1002080303",
    "1002080304",
    "1002080401",
    "1002080402",
    "1002080501",
    "1002080502",
    "1002080601",
    "1002080701",
    "1002080801",
    "1002080901",
    "1002090101",
    "1002080301",
    "1002100102",
    "1001060606",
    "1001060607",
    "1001090305",
}


def _digits_only(x) -> str:
    """Keep digits only, and remove leading zeros."""
    if x is None:
        return ""
    if isinstance(x, float) and pd.isna(x):
        return ""
    s = str(x)
    digits = "".join(ch for ch in s if ch.isdigit())
    digits = digits.lstrip("0")
    return digits if digits else ""


def _product_alert(days: float) -> str:
    if pd.isna(days):
        return ""
    if days < 7:
        return "RED"
    if days < 21:
        return "ORANGE"
    if days < 60:
        return "YELLOW"
    return "GREEN"


def _cust_alert(days: float) -> str:
    if pd.isna(days):
        return ""
    if days < 5:
        return "RED"
    if days < 10:
        return "ORANGE"
    if days < 20:
        return "YELLOW"
    return "GREEN"


def add_alerts(df: pd.DataFrame) -> pd.DataFrame:
    """
    Adds:
      - Product Alert (based on Total Coverage Days)
      - CUST Alert   (based on CUST Stock Days)
    Then applies exclusion list so excluded items do NOT show in Customer Alert tab/export.
    """
    if df is None:
        return None

    out = df.copy()

    out["Product Alert"] = out["Total Coverage Days"].apply(_product_alert)
    out["CUST Alert"] = out["CUST Stock Days"].apply(_cust_alert)

    # Apply exclusions by Item No.
    if "Item No." in out.columns:
        item_norm = out["Item No."].apply(_digits_only)
        mask_excluded = item_norm.isin(CUST_ALERT_EXCLUDE_ITEMNOS)
        out.loc[mask_excluded, "CUST Alert"] = ""  # مهم: يمسح التنبيه فقط

    return out
