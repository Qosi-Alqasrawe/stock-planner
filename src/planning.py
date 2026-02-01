# src/planning.py  (كامل)
import pandas as pd
import numpy as np

def add_planning_columns(df: pd.DataFrame, start_date, working_days: int) -> pd.DataFrame:
    out = df.copy()

    if "Plan Qty Input" not in out.columns:
        out["Plan Qty Input"] = np.nan
    if "Plan Months Input" not in out.columns:
        out["Plan Months Input"] = np.nan

    out["Months Covered by Plan Qty"] = (pd.to_numeric(out["Plan Qty Input"], errors="coerce") / out["Monthly Demand"]).round(1)
    out["Months Covered by Plan Qty"] = out["Months Covered by Plan Qty"].replace([np.inf, -np.inf], np.nan).fillna(0)

    out["Qty Needed for Plan Months"] = (pd.to_numeric(out["Plan Months Input"], errors="coerce") * out["Monthly Demand"]).round(0)
    out["Qty Needed for Plan Months"] = pd.to_numeric(out["Qty Needed for Plan Months"], errors="coerce").fillna(0).astype(int)

    out["Start Date"] = pd.to_datetime(start_date)
    days_int = out["Total Coverage Days"].round(0).fillna(0).astype(int)
    out["End Date"] = out["Start Date"] + pd.to_timedelta(days_int, unit="D")

    out["Start Date"] = out["Start Date"].dt.date
    out["End Date"] = out["End Date"].dt.date
    return out
