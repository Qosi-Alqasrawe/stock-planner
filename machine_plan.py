# src/machine_plan.py
import math
import pandas as pd


def _classify_by_percentiles(monthly_demand: pd.Series) -> pd.Series:
    """
    5 buckets by percentiles on Monthly Demand:
    VERY_HIGH / HIGH / MEDIUM / LOW / VERY_LOW
    """
    s = pd.to_numeric(monthly_demand, errors="coerce").fillna(0.0)
    nz = s[s > 0]
    if len(nz) == 0:
        return pd.Series(["VERY_LOW"] * len(s), index=s.index)

    p90 = float(nz.quantile(0.90))
    p70 = float(nz.quantile(0.70))
    p40 = float(nz.quantile(0.40))
    p15 = float(nz.quantile(0.15))

    def cls(v):
        v = float(v)
        if v >= p90:
            return "VERY_HIGH"
        if v >= p70:
            return "HIGH"
        if v >= p40:
            return "MEDIUM"
        if v >= p15:
            return "LOW"
        return "VERY_LOW"

    return s.apply(cls)


def build_production_plan(
    master: pd.DataFrame,
    machine_col: str,
    working_days: int = 26,
    demand_col: str = "Monthly Demand",
    daily_col: str = "Daily Demand",
    coverage_days_col: str = "Total Coverage Days",
    alert_col: str = "Product Alert",
    itemno_col: str = "Item No.",
    target_months_map: dict | None = None,
    safety_days_map: dict | None = None,
    min_batch_months_map: dict | None = None,
    batch_round_to: int = 0,
    # optional overrides if you later add columns in Excel
    target_months_input_col: str = "Target Months Input",
    safety_days_input_col: str = "Safety Days Input",
) -> pd.DataFrame:
    df = master.copy()

    for col in [machine_col, demand_col, daily_col, coverage_days_col]:
        if col not in df.columns:
            raise ValueError(f"Missing required column: {col}")

    # Defaults if not provided
    if target_months_map is None:
        target_months_map = {"VERY_HIGH": 3, "HIGH": 4, "MEDIUM": 6, "LOW": 9, "VERY_LOW": 12}
    if safety_days_map is None:
        safety_days_map = {"VERY_HIGH": 14, "HIGH": 10, "MEDIUM": 7, "LOW": 5, "VERY_LOW": 3}
    if min_batch_months_map is None:
        min_batch_months_map = {"LOW": 4, "VERY_LOW": 6}

    # ============ 1) explode machines ============
    df[machine_col] = df[machine_col].astype(str)
    df["Machine"] = df[machine_col].str.split("-")
    df = df.explode("Machine")
    df["Machine"] = df["Machine"].astype(str).str.strip()
    df.loc[df["Machine"].isin(["nan", "None", ""]), "Machine"] = "UNSPECIFIED"

    # ============ 2) demand class ============
    df["Demand Class"] = _classify_by_percentiles(df[demand_col])

    # ============ 3) auto target months + safety days ============
    df["Target Months (Auto)"] = df["Demand Class"].map(target_months_map).fillna(6).astype(float)
    df["Safety Days (Auto)"] = df["Demand Class"].map(safety_days_map).fillna(5).astype(float)

    # overrides (optional columns)
    if target_months_input_col in df.columns:
        tm_in = pd.to_numeric(df[target_months_input_col], errors="coerce")
        df["Target Months (Final)"] = df["Target Months (Auto)"]
        df.loc[tm_in.notna() & (tm_in > 0), "Target Months (Final)"] = tm_in
    else:
        df["Target Months (Final)"] = df["Target Months (Auto)"]

    if safety_days_input_col in df.columns:
        sd_in = pd.to_numeric(df[safety_days_input_col], errors="coerce")
        df["Safety Days (Final)"] = df["Safety Days (Auto)"]
        df.loc[sd_in.notna() & (sd_in >= 0), "Safety Days (Final)"] = sd_in
    else:
        df["Safety Days (Final)"] = df["Safety Days (Auto)"]

    # ============ 4) plan math ============
    current_cov = pd.to_numeric(df[coverage_days_col], errors="coerce").fillna(0.0)
    daily = pd.to_numeric(df[daily_col], errors="coerce").fillna(0.0)
    monthly = pd.to_numeric(df[demand_col], errors="coerce").fillna(0.0)

    df["Target Coverage Days"] = (df["Target Months (Final)"] * float(working_days)).round(1)
    df["Plan Coverage Days"] = (df["Target Coverage Days"] + df["Safety Days (Final)"]).round(1)

    # ===== Minimum Batch Months for slow movers (forces bigger plan coverage) =====
    min_months = df["Demand Class"].map(min_batch_months_map).fillna(0).astype(float)
    min_plan_days = (min_months * float(working_days)).round(1)
    df["Plan Coverage Days"] = pd.concat([df["Plan Coverage Days"], min_plan_days.rename("MinPlan")], axis=1).max(axis=1).round(1)

    # Gap Days
    df["Gap Days"] = (df["Plan Coverage Days"] - current_cov).clip(lower=0).round(1)

    # Proposed Qty (ceil)
    df["Proposed Production Qty"] = (df["Gap Days"] * daily).apply(lambda x: int(math.ceil(max(0.0, x))))

    # ===== Round to batch size =====
    if batch_round_to and int(batch_round_to) > 0:
        b = int(batch_round_to)
        df["Proposed Production Qty"] = df["Proposed Production Qty"].apply(lambda q: int(math.ceil(q / b) * b) if q > 0 else 0)

    # ============ 5) priority + rank ============
    priority_map = {"RED": 0, "ORANGE": 1, "YELLOW": 2, "GREEN": 3}
    df["_AlertPriority"] = df.get(alert_col, "").map(priority_map).fillna(9).astype(int)

    df["_MonthlyDemandNum"] = monthly

    df = df.sort_values(
        by=["Machine", "_AlertPriority", "Gap Days", "Proposed Production Qty", "_MonthlyDemandNum"],
        ascending=[True, True, False, False, False],
    )

    df["Production Rank"] = df.groupby("Machine").cumcount() + 1

    df = df.drop(columns=["_AlertPriority", "_MonthlyDemandNum"], errors="ignore")

    if itemno_col in df.columns:
        df[itemno_col] = df[itemno_col].astype(str)

    return df


def make_machine_sheets(plan_df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    sheets = {}
    machines = sorted(plan_df["Machine"].dropna().unique().tolist())
    for m in machines:
        df_m = plan_df[plan_df["Machine"] == m].copy()
        if df_m.empty:
            continue

        safe_name = (
            str(m)[:31]
            .replace("/", "-").replace("\\", "-").replace(":", "-").replace("*", "-")
            .replace("?", "").replace("[", "(").replace("]", ")")
        )
        sheets[safe_name] = df_m

    return sheets
