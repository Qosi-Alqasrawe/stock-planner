# app.py
import streamlit as st
import pandas as pd
import numpy as np

from src.io import read_excel_any, normalize_cols, merge_items_stock, format_item_no
from src.metrics import compute_metrics_qty_based
from src.alerts import add_alerts
from src.planning import add_planning_columns
from src.machine_plan import build_production_plan, make_machine_sheets
from src.export import to_excel_bytes, to_excel_bytes_multi_sheets
from src.config import PLAN_DEFAULTS, WORKING_DAYS, DEMAND_CLASSES
from src.production_plan_export import fill_qty_in_client_orders
from src.report_export import (
    export_management_report_excel_bytes,
    export_management_report_pdf_bytes,
)

st.set_page_config(page_title="Stock Planner", layout="wide")
st.title("Stock Planner")

START_DATE_DEFAULT = "2026-01-30"

# ========= Upload =========
st.sidebar.header("Upload")
items_file = st.sidebar.file_uploader("Items file (.xlsx)", type=["xlsx"])
stock_file = st.sidebar.file_uploader("Stock file (.xls/.xlsx)", type=["xls", "xlsx"])

st.sidebar.header("Start Date")
start_date = st.sidebar.date_input("Start Date", value=pd.to_datetime(START_DATE_DEFAULT).date())

# ========= Production Plan Settings (Sidebar) =========
st.sidebar.header("Production Plan Settings")

batch_round_to = st.sidebar.number_input(
    "Batch round to (0 = off)",
    min_value=0,
    value=int(PLAN_DEFAULTS["batch_round_to"]),
    step=100,
)

st.sidebar.subheader("Target Months (by demand class)")
tm_defaults = PLAN_DEFAULTS["target_months_map"]
tm = {}
for k in DEMAND_CLASSES:
    tm[k] = st.sidebar.number_input(k, min_value=1, max_value=24, value=int(tm_defaults[k]), step=1)

st.sidebar.subheader("Minimum Batch Months (slow movers)")
mb_defaults = PLAN_DEFAULTS["min_batch_months_map"]
mb = {}
mb["LOW"] = st.sidebar.number_input("LOW (min months)", min_value=0, max_value=24, value=int(mb_defaults["LOW"]), step=1)
mb["VERY_LOW"] = st.sidebar.number_input("VERY_LOW (min months)", min_value=0, max_value=24, value=int(mb_defaults["VERY_LOW"]), step=1)

# (اختياري) إذا بدك تخلي السيفتي قابل للتغيير كمان:
# st.sidebar.subheader("Safety Days (by demand class)")
# sd_defaults = PLAN_DEFAULTS["safety_days_map"]
# sd = {}
# for k in DEMAND_CLASSES:
#     sd[k] = st.sidebar.number_input(f"{k} safety", min_value=0, max_value=60, value=int(sd_defaults[k]), step=1)
# وإلا خليناها ثابتة من PLAN_DEFAULTS:
sd = PLAN_DEFAULTS["safety_days_map"].copy()

if not stock_file:
    st.info("Upload Stock file.")
    st.stop()

# ========= Read =========
stock_df = normalize_cols(read_excel_any(stock_file))
items_df = normalize_cols(read_excel_any(items_file)) if items_file else None

# ========= Columns =========
STOCK_ITEMNO_COL = "Item No."
STOCK_NAME_COL = "Item Name"
MACHINE_COL = "Production Line (Stage1-Stage2-Stage3)"

MPI_QTY_COL = "MPI Stock"
CUST_QTY_COL = "CUST Stock"
TOTAL_QTY_COL = "Total Stock"

MPI_DAYS_COL = "MPI Stock days"
CUST_DAYS_COL = "CUST Stock Days"
MONTHLY_DEMAND_COL = "Min Stock / M.D."

ITEMS_ID_COL = "ID"
ITEMS_DESC_COL = "Description"

required_stock = [
    STOCK_ITEMNO_COL, STOCK_NAME_COL, MACHINE_COL,
    MPI_QTY_COL, CUST_QTY_COL, TOTAL_QTY_COL,
    MPI_DAYS_COL, CUST_DAYS_COL, MONTHLY_DEMAND_COL
]
missing_stock = [c for c in required_stock if c not in stock_df.columns]
if missing_stock:
    st.error("Missing required columns in Stock file:")
    st.write(missing_stock)
    st.stop()

# ========= Merge Description =========
merged = stock_df.copy()
merged[STOCK_ITEMNO_COL] = merged[STOCK_ITEMNO_COL].apply(format_item_no)

if items_df is not None:
    merged = merge_items_stock(
        stock_df=merged,
        items_df=items_df,
        stock_item_col=STOCK_ITEMNO_COL,
        items_item_col=ITEMS_ID_COL,
        keep_items_cols=[ITEMS_DESC_COL],
    )

if "Description" not in merged.columns:
    merged["Description"] = ""

# ========= Compute base metrics =========
try:
    calc = compute_metrics_qty_based(
        merged,
        working_days=WORKING_DAYS,
        monthly_demand_col=MONTHLY_DEMAND_COL,
        mpi_qty_col=MPI_QTY_COL,
        cust_qty_col=CUST_QTY_COL,
        total_qty_col=TOTAL_QTY_COL,
        mpi_days_col=MPI_DAYS_COL,
        cust_days_col=CUST_DAYS_COL,
    )
except Exception as e:
    st.error(f"Compute metrics failed: {e}")
    st.stop()

if calc is None or not hasattr(calc, "columns"):
    st.error("compute_metrics_qty_based returned None (check required columns/types).")
    st.stop()

# ========= Planning columns (Excel formulas etc) =========
master = add_planning_columns(calc, start_date=start_date, working_days=WORKING_DAYS)

# ========= Alerts =========
master = add_alerts(master)

# ========= Tabs =========
tab1, tab_prod, tab_cust, tab_machine, tab_final, tab_export, tab_reports = st.tabs(
    ["Master", "Product Alert", "Customer Alert", "Machine Plan", "Final Plan", "Export (Full)", "Reports"]
)


# ===== Master =====
with tab1:
    st.subheader("Master")

    HIDE_IN_UI = {
        "Plan Qty Input",
        "Months Covered by Plan Qty",
        "Plan Months Input",
        "Qty Needed for Plan Months",
    }

    view_cols = [c for c in master.columns if c not in HIDE_IN_UI]
    st.dataframe(master[view_cols], use_container_width=True, hide_index=True)

# ===== Product Alert =====
with tab_prod:
    st.subheader("Product Alert")

    prod_df = master[master["Product Alert"].isin(["RED", "ORANGE"])].copy()
    if len(prod_df) == 0:
        st.info("No RED/ORANGE items for Product Alert.")
    else:
        alert_order = pd.Categorical(prod_df["Product Alert"], categories=["RED", "ORANGE", "YELLOW", "GREEN"], ordered=True)
        prod_df["__a"] = alert_order
        prod_df = prod_df.sort_values(["__a", "Total Coverage Days"], ascending=[True, True]).drop(columns="__a")

        st.dataframe(prod_df, use_container_width=True, hide_index=True)

        prod_bytes = to_excel_bytes(master_df=prod_df, machine_df=None, itemno_col="Item No.", master_sheet="Product_Alert")
        st.download_button(
            "Download Product Alert Excel",
            data=prod_bytes,
            file_name="Product_Alert.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ===== Customer Alert =====
with tab_cust:
    st.subheader("Customer Alert")

    cust_df = master[master["CUST Alert"].isin(["RED", "ORANGE"])].copy()
    if len(cust_df) == 0:
        st.info("No RED/ORANGE items for Customer Alert.")
    else:
        alert_order = pd.Categorical(cust_df["CUST Alert"], categories=["RED", "ORANGE", "YELLOW", "GREEN"], ordered=True)
        cust_df["__a"] = alert_order
        cust_df = cust_df.sort_values(["__a", "CUST Stock Days"], ascending=[True, True]).drop(columns="__a")

        st.dataframe(cust_df, use_container_width=True, hide_index=True)

        cust_bytes = to_excel_bytes(master_df=cust_df, machine_df=None, itemno_col="Item No.", master_sheet="Customer_Alert")
        st.download_button(
            "Download Customer Alert Excel",
            data=cust_bytes,
            file_name="Customer_Alert.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ===== Machine Plan =====
# ===== Machine Plan =====
with tab_machine:
    st.subheader("Machine Plan (Production Plan)")

    plan = build_production_plan(
        master=master,
        machine_col=MACHINE_COL,
        working_days=WORKING_DAYS,
        demand_col="Monthly Demand",
        daily_col="Daily Demand",
        coverage_days_col="Total Coverage Days",
        alert_col="Product Alert",
        target_months_map=tm,
        safety_days_map=sd,
        min_batch_months_map=mb,
        batch_round_to=int(batch_round_to),
    )

    machines = sorted(plan["Machine"].dropna().unique().tolist())
    chosen = st.selectbox("Machine", machines, index=0 if machines else None)

    view_cols = [
        "Production Rank",
        "Item No.",
        "Item Name",
        "Machine",
        "Product Alert",
        "Monthly Demand",
        "Daily Demand",
        "Total Coverage Days",
        "Demand Class",
        "Target Months (Final)",
        "Safety Days (Final)",
        "Plan Coverage Days",
        "Gap Days",
        "Proposed Production Qty",
    ]
    view_cols = [c for c in view_cols if c in plan.columns]

    view = plan[plan["Machine"] == chosen].copy() if chosen else plan.copy()
    st.dataframe(view[view_cols], use_container_width=True, hide_index=True)

    st.divider()

    # =========================
    # 1) Download FULL (All Machines) - (Sheet لكل ماكينة)
    # =========================
    sheets_all = make_machine_sheets(plan[view_cols].copy())
    excel_bytes_all = to_excel_bytes_multi_sheets(sheets_all)

    st.download_button(
        "Download FULL Production Plan (All Machines)",
        data=excel_bytes_all,
        file_name="Machine_Plan_All_Machines.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_mp_all",
    )

    # =========================
    # 2) Download per machine (ملف لكل ماكينة لحال)
    # =========================
    st.subheader("Download Production Plan per Machine")

    if not machines:
        st.info("No machines found in Production Plan.")
    else:
        for m in machines:
            df_m = plan[plan["Machine"] == m][view_cols].copy()
            if df_m.empty:
                continue

            # ملف Excel فيه شيت واحد (اسم الشيت = الماكينة)
            sheets_m = make_machine_sheets(df_m)
            excel_bytes_m = to_excel_bytes_multi_sheets(sheets_m)

            safe_name = str(m).replace("/", "_").replace("\\", "_").replace(" ", "_")

            st.download_button(
                f"Download {m}",
                data=excel_bytes_m,
                file_name=f"Machine_Plan_{safe_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_mp_{safe_name}",
            )


# ===== Full Export =====
with tab_export:
    st.subheader("Export (Full Master + Production Plan)")

    mp_full = build_production_plan(
        master=master,
        machine_col=MACHINE_COL,
        working_days=WORKING_DAYS,
        demand_col="Monthly Demand",
        daily_col="Daily Demand",
        coverage_days_col="Total Coverage Days",
        alert_col="Product Alert",
        target_months_map=tm,
        safety_days_map=sd,
        min_batch_months_map=mb,
        batch_round_to=int(batch_round_to),
    )

    excel_bytes = to_excel_bytes(
        master_df=master,
        machine_df=mp_full,
        itemno_col="Item No.",
        master_sheet="Master",
    )

    st.download_button(
        "Download Excel (Master + Production_Plan)",
        data=excel_bytes,
        file_name="Stock_Planning_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with tab_final:
    st.subheader("Final Plan (Decide Final Qty + Export to Production Plan.xlsx)")

    # Upload Production Plan template
    tpl = st.file_uploader("Upload Production Plan.xlsx template", type=["xlsx"], key="prod_plan_tpl")
    if tpl is None:
        st.info("Upload Production Plan.xlsx to enable auto-fill.")
        st.stop()

    # Build a compact decision table from the plan
    plan_final = build_production_plan(
        master=master,
        machine_col=MACHINE_COL,
        working_days=WORKING_DAYS,
        demand_col="Monthly Demand",
        daily_col="Daily Demand",
        coverage_days_col="Total Coverage Days",
        alert_col="Product Alert",
        target_months_map=tm,
        safety_days_map=sd,
        min_batch_months_map=mb,
        batch_round_to=int(batch_round_to),
    ).copy()

    # Keep only needed columns
    cols = [
        "Production Rank",
        "Item No.",
        "Item Name",
        "Machine",
        "Product Alert",
        "Monthly Demand",
        "Total Coverage Days",
        "Proposed Production Qty",
    ]
    cols = [c for c in cols if c in plan_final.columns]
    plan_final = plan_final[cols].copy()

    # Add Final Qty column for user decision
    if "Final Qty" not in plan_final.columns:
        plan_final["Final Qty"] = plan_final["Proposed Production Qty"]

    # Let user edit Final Qty
    edited = st.data_editor(
        plan_final,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "Final Qty": st.column_config.NumberColumn("Final Qty", min_value=0, step=100),
        },
    )

    # Export filled template (fills ONLY Qty in Clinet Orders by Item No)
    if st.button("Fill Production Plan.xlsx Automatically"):
        template_bytes = tpl.getvalue()

        filled_bytes, filled_count, not_matched, unmatched = fill_qty_in_client_orders(
            template_bytes=template_bytes,
            plan_df=edited,
            plan_itemno_col="Item No.",
            plan_qty_col="Final Qty",          # or "Proposed Production Qty"
            sheet_name="Clinet Orders",
            header_row=2,
            start_row=3,
            sheet_itemno_header="Item No.",
            sheet_qty_header="Qty",
        )

        st.success(f"Filled Qty for {filled_count} rows. Not matched: {not_matched}")

        if not_matched > 0:
            st.warning("Unmatched Item No. (showing first 50):")
            st.write(unmatched[:50])

        st.download_button(
            "Download Filled Production Plan.xlsx",
            data=filled_bytes,
            file_name="Production Plan - Filled.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

with tab_reports:
    st.subheader("Management Reports (English)")

    master_df = master
    plan_df = plan if "plan" in locals() else None
    final_df = edited if "edited" in locals() else None

    col1, col2 = st.columns(2)

    with col1:
        excel_bytes = export_management_report_excel_bytes(
            master_df=master_df,
            plan_df=plan_df,
            final_df=final_df,
        )
        st.download_button(
            "Download Management Report (Excel)",
            data=excel_bytes,
            file_name="Stock_Planner_Management_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_mgmt_excel",
        )

    with col2:
        pdf_bytes = export_management_report_pdf_bytes(
            master_df=master_df,
            plan_df=plan_df,
            final_df=final_df,
            title="Stock Planner - Management Summary",
        )
        st.download_button(
            "Download Management Summary (PDF)",
            data=pdf_bytes,
            file_name="Stock_Planner_Management_Summary.pdf",
            mime="application/pdf",
            key="dl_mgmt_pdf",
        )

    st.caption(
        "Excel contains full tables (filterable). PDF contains KPIs + Top 10 critical + Top 30 priority + Machine load."
    )
