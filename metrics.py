# src/metrics.py  (كامل)
import pandas as pd
import numpy as np

def compute_metrics_qty_based(
    df: pd.DataFrame,
    working_days: int,
    monthly_demand_col: str,
    mpi_qty_col: str,
    cust_qty_col: str,
    total_qty_col: str,
    mpi_days_col: str,
    cust_days_col: str,
) -> pd.DataFrame:
    out = df.copy()

    # Monthly/Daily demand
    out["Monthly Demand"] = pd.to_numeric(out[monthly_demand_col], errors="coerce").fillna(0).round(0).astype(int)
    out["Daily Demand"] = (out["Monthly Demand"] / float(working_days)).round(0).astype(int)

    # Quantities (given)
    out["MPI Stock Qty"] = pd.to_numeric(out[mpi_qty_col], errors="coerce").fillna(0).round(0).astype(int)
    out["CUST Stock Qty"] = pd.to_numeric(out[cust_qty_col], errors="coerce").fillna(0).round(0).astype(int)
    out["Stock"] = pd.to_numeric(out[total_qty_col], errors="coerce").fillna(0).round(0).astype(int)

    # Days (given)
    out["MPI Stock days"] = pd.to_numeric(out[mpi_days_col], errors="coerce").fillna(0).round(1)
    out["CUST Stock Days"] = pd.to_numeric(out[cust_days_col], errors="coerce").fillna(0).round(1)

    # Coverage
    out["Total Coverage Days"] = (out["MPI Stock days"] + out["CUST Stock Days"]).round(1)
    out["Total Coverage Month"] = (out["Total Coverage Days"] / float(working_days)).round(1)

    return out
