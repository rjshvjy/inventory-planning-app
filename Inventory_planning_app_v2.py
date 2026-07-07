"""
Inventory_planning_app_v2.py — entry point for the v2 Inventory Planning App.

Phase 2 harness: uploads the three Amazon files, runs the ingestion +
validation layer (ingestion_v2.py), and shows the validation report and
canonical-data previews. Planning logic arrives in later phases; nothing
downstream runs unless validation passes (fail loud, fail specific).

Reference files (configurations.xlsx, inventory_plan_template.xlsx,
fc_registration.pdf) are read from the repo's reference/ folder; optional
uploaders below let you override them for a one-off test (upload beats repo).
"""

import streamlit as st
import pandas as pd
from ingestion_v2 import run_ingestion

st.set_page_config(page_title="Inventory Planning v2", page_icon="📦",
                   layout="wide")
st.title("📦 Inventory Planning v2 — Ingestion & Validation")
st.caption("Phase 2 build: upload the three Amazon files and validate. "
           "Planning phases plug in behind this gate.")

with st.sidebar:
    st.header("Per-run uploads")
    f_sales = st.file_uploader("Amazon sales export (CSV)", type=["csv", "txt"])
    f_general = st.file_uploader("General stock — Manage FBA Inventory (CSV)",
                                 type=["csv", "txt"])
    f_ledger = st.file_uploader("FC-wise stock — Inventory Ledger (CSV)",
                                type=["csv", "txt"])
    window_override = st.number_input(
        "Sales window days (only if the sales file has no date column)",
        min_value=0, max_value=365, value=0,
        help="Leave 0 when the export carries dates — the app infers the "
             "window itself.")
    st.divider()
    with st.expander("Override repo reference files (optional)"):
        o_cfg = st.file_uploader("configurations.xlsx", type=["xlsx"])
        o_master = st.file_uploader("inventory_plan_template.xlsx",
                                    type=["xlsx"])
        o_fcreg = st.file_uploader("fc_registration (pdf/csv)",
                                   type=["pdf", "csv"])
    go = st.button("🔍 Validate", type="primary",
                   disabled=not (f_sales and f_general and f_ledger))

if not go:
    st.info("Upload the three Amazon files in the sidebar and press Validate. "
            "Reference files load from the repo's reference/ folder unless "
            "overridden.")
    st.stop()

can, rep = run_ingestion(
    f_sales, f_general, f_ledger,
    config_path=o_cfg or "reference/configurations.xlsx",
    master_path=o_master or "reference/inventory_plan_template.xlsx",
    fcreg_path=o_fcreg or "reference/fc_registration.pdf",
    window_days_override=window_override or None,
)

c1, c2, c3 = st.columns(3)
c1.metric("Errors (block the run)", len(rep.errors))
c2.metric("Warnings", len(rep.warnings))
c3.metric("Checks noted", len(rep.info))

if rep.errors:
    st.error("**Validation failed — nothing was planned.** Fix the items "
             "below and rerun; the app never guesses past an error.")
    for e in rep.errors:
        st.error(e)
if rep.warnings:
    for w in rep.warnings:
        st.warning(w)
with st.expander("Details / info", expanded=not rep.errors):
    for i in rep.info:
        st.write("• " + i)

if not rep.ok:
    st.stop()

st.success("✅ Validation passed — canonical data assembled. "
           "(Planning engine plugs in here in Phase 3.)")

t1, t2, t3, t4 = st.tabs(["SKU master", "Daily sales (sku × state)",
                          "Available (sku × region)", "National states"])
with t1:
    st.dataframe(can.sku_master.drop(columns=["row"]),
                 use_container_width=True, height=420)
with t2:
    st.caption(f"Window used: {can.sales_window_days:.0f} days "
               f"→ daily = window total ÷ days")
    st.dataframe(can.sales_daily.sort_values("daily", ascending=False),
                 use_container_width=True, height=420)
with t3:
    piv = (can.available.pivot(index="sku_u", columns="region", values="bal")
           .fillna(0).astype(int))
    order = [r for r in can.regions if r in piv.columns] \
        + [c for c in piv.columns if c not in can.regions]
    st.dataframe(piv[order], use_container_width=True, height=420)
with t4:
    st.dataframe(can.national, use_container_width=True, height=420)
