"""
Inventory_planning_app_v2.py — v2 Inventory Planning App (Phase 3: full flow).

Gate order (spec v2.8, handover §10): ingestion ok -> stock workbook ok ->
ASIN suppression judgments -> run_calculation -> build_plan_workbook ->
download. Nothing downstream runs past a failed gate (fail loud, fail
specific). Reference files load from the repo's reference/ folder; optional
uploaders override them for a one-off test (upload beats repo).

Design note: STATELESS reruns — every widget interaction re-runs the whole
pipeline from the uploaded bytes (files are small; this avoids session-state
staleness bugs). Judgment radios and days-of-cover therefore always act on
freshly validated data.
"""

import io

import pandas as pd
import streamlit as st

from ingestion_v2 import run_ingestion
from workbook_builder_v2 import (build_stock_workbook, read_stock_workbook)
from calculation_v2 import run_calculation, detect_asin_ambiguities
from output_v2 import build_plan_workbook

st.set_page_config(page_title="Inventory Planning v2", page_icon="📦",
                   layout="wide")
st.title("📦 Inventory Planning v2")
st.caption("Upload the three Amazon files + your filled daily stock "
           "workbook, set days of cover, resolve any ASIN judgments, plan.")


def _buf(uploaded):
    """Re-wrap an uploader so reruns always read from position 0."""
    if uploaded is None:
        return None
    b = io.BytesIO(uploaded.getvalue())
    b.name = uploaded.name
    return b


# ------------------------------------------------------------------ sidebar
with st.sidebar:
    with st.expander("📖 Help & guides"):
        st.markdown(
            "- [User guide](https://rjshvjy.github.io/inventory-planning-app/"
            "help_user.html) — how to run a plan, step by step\n"
            "- [Admin guide](https://rjshvjy.github.io/inventory-planning-app/"
            "help_admin.html) — architecture, config, maintenance")
    st.header("Amazon files (per run)")
    f_sales = st.file_uploader("Sales export (CSV)", type=["csv", "txt"])
    f_general = st.file_uploader("General stock — Manage FBA Inventory (CSV)",
                                 type=["csv", "txt"])
    f_ledger = st.file_uploader("FC-wise stock — Inventory Ledger (CSV)",
                                type=["csv", "txt"])
    st.header("Daily stock workbook (filled)")
    f_workbook = st.file_uploader("daily_stock_workbook.xlsx", type=["xlsx"])
    st.header("Planning input")
    days_of_cover = st.number_input(
        "Days of cover to plan for", min_value=1, max_value=180, value=30,
        help="THE core planning input (spec §9): stock each region up to "
             "lead time + this many days of sales.")
    window_override = st.number_input(
        "Sales window days (only if the sales file has no date column)",
        min_value=0, max_value=365, value=0)
    st.divider()
    with st.expander("Override repo reference files (optional)"):
        o_cfg = st.file_uploader("configurations.xlsx", type=["xlsx"])
        o_master = st.file_uploader("inventory_plan_template.xlsx",
                                    type=["xlsx"])
        o_fcreg = st.file_uploader("fc_registration (pdf/csv)",
                                   type=["pdf", "csv"])

if not (f_sales and f_general and f_ledger):
    st.info("Upload the three Amazon files in the sidebar to begin. The "
            "daily stock workbook is needed for planning; without it you "
            "can still validate and generate a fresh workbook below.")

# ---------------------------------------------------------- gate 1: ingest
if not (f_sales and f_general and f_ledger):
    st.stop()

def _ingest(fc_resolutions=None, recency_acknowledged=False):
    return run_ingestion(
        _buf(f_sales), _buf(f_general), _buf(f_ledger),
        config_path=_buf(o_cfg) or "reference/configurations.xlsx",
        master_path=_buf(o_master) or "reference/inventory_plan_template.xlsx",
        fcreg_path=_buf(o_fcreg) or "reference/fc_registration.pdf",
        window_days_override=window_override or None,
        fc_resolutions=fc_resolutions,
        recency_acknowledged=recency_acknowledged,
    )

# Pass 1 — detect. v2.9 gates are answered below; pass 2 applies them.
can, rep = _ingest()

# ---------- v2.9 §5c: FC-review gate (interactive, per-run, stateless) ----
fc_resolutions = {}
if rep.ok and rep.needs_fc_review:
    st.subheader("FC review needed — stock in unregistered FC(s)")
    st.write("The ledger holds sellable stock in FC(s) not present in the "
             "registration file. Decide per FC **for this run** (the "
             "permanent fix is committing a refreshed registration export — "
             "this prompt then disappears on its own):")
    aborted = False
    for fc, units in rep.needs_fc_review.items():
        st.markdown(f"**{fc}** — holds **{units} sellable units**, not in "
                    f"the registration file.")
        action = st.selectbox(
            f"How should {fc} be treated this run?",
            options=["map", "ixd", "ignore", "abort"],
            format_func=lambda v: {
                "map": "Map to a region (recommended for a normal regional "
                       "FC)",
                "ixd": "Treat as IXD cross-dock (rare — only if it truly "
                       "redistributes)",
                "ignore": "Ignore — show its stock, never count it",
                "abort": "Abort — I will fix the registration file instead",
            }[v], key=f"fcact_{fc}")
        if action == "abort":
            aborted = True
        elif action == "map":
            region = st.selectbox(
                f"Which region serves {fc}?", options=can.regions,
                key=f"fcreg_{fc}")
            fulfillable = st.selectbox(
                f"Is {fc}'s stock currently fulfillable?",
                options=[True, False],
                format_func=lambda v: ("Fulfillable — count it as supply "
                                       "(default)") if v else
                                      ("Not fulfillable — show it, exclude "
                                       "from planning (e.g. GST-frozen)"),
                key=f"fcful_{fc}")
            fc_resolutions[fc] = {"action": "map", "region": region,
                                  "fulfillable": fulfillable}
        else:
            fc_resolutions[fc] = {"action": action}
    if aborted:
        st.error("Run aborted at your request — commit a refreshed "
                 "registration file and rerun.")
        st.stop()

# ---------- v2.9 §6a: sales-recency acknowledge gate ----------------------
recency_ok = False
if rep.ok and rep.needs_recency_ack:
    ra = rep.needs_recency_ack
    st.subheader("Sales-window recency — please confirm")
    st.write(f"Your sales window ends **{ra['sales_end']}** but stock is as "
             f"of **{ra['stock_anchor']}** — a {ra['gap_days']}-day gap "
             f"(threshold {ra['threshold']:.0f}d). Planning will treat this "
             f"window's velocity as **current**. The Sales-velocity tab in "
             f"the output shows the rates applied, per region.")
    recency_ok = st.checkbox(
        "I understand — use this sales window as current velocity",
        key="recency_ack")

# Pass 2 — apply answers (only if anything needed answering)
if fc_resolutions or recency_ok:
    can, rep = _ingest(fc_resolutions=fc_resolutions or None,
                       recency_acknowledged=recency_ok)

c1, c2, c3 = st.columns(3)
c1.metric("Errors (block the run)", len(rep.errors))
c2.metric("Warnings", len(rep.warnings))
c3.metric("Checks noted", len(rep.info))
for e in rep.errors:
    st.error(e)
for w in rep.warnings:
    st.warning(w)
with st.expander("Validation details", expanded=False):
    for i in rep.info:
        st.write("• " + i)
if not rep.ok:
    st.error("**Validation failed — nothing was planned.** Fix the items "
             "above and rerun; the app never guesses past an error.")
    st.stop()
if rep.needs_fc_review:
    st.info("Answer the FC-review question(s) above to continue — your "
            "choices apply on the next rerun (widgets re-run the pipeline "
            "automatically).")
    st.stop()
if rep.needs_recency_ack:
    st.info("Tick the sales-recency confirmation above to continue.")
    st.stop()
if can.fc_resolutions:
    st.warning("This run uses per-run FC mapping(s): " + ", ".join(
        f"{fc} → {r.get('region', r['action'].upper())}"
        + ("" if r.get("fulfillable", True) else " (NOT fulfillable — "
           "excluded)") for fc, r in can.fc_resolutions.items())
        + ". Affected plan lines are flagged FC_MAPPED in the output.")
st.success("✅ Amazon files validated.")

# fresh-workbook generation is always available once ingestion passes
with st.expander("Need a fresh daily stock workbook? (generate here)"):
    st.write("Builds today's workbook from the validated data — Current "
             "stock prefilled from the ledger for review, In-transit and "
             "At-the-FC tabs ready to fill. Commit the filled file to the "
             "repo per the usual discipline.")
    prefill = st.checkbox("Prefill Current stock from today's ledger", True)
    st.download_button(
        "⬇️ Generate & download fresh workbook",
        data=build_stock_workbook(can, prefill=prefill),
        file_name=f"daily_stock_workbook_{pd.Timestamp.now():%d_%b_%Y}.xlsx")

# ------------------------------------------------------ gate 2: workbook
if not f_workbook:
    st.info("Upload your **filled** daily stock workbook in the sidebar to "
            "continue to planning.")
    st.stop()

stock, rep2 = read_stock_workbook(_buf(f_workbook), can)
for e in rep2.errors:
    st.error(e)
for w in rep2.warnings:
    st.warning(w)
if not rep2.ok:
    st.error("**Stock workbook failed validation — nothing was planned.**")
    st.stop()
st.success(f"✅ Stock workbook read — anchor date "
           f"{stock.stock_as_of.date()} (all depletion math anchors here, "
           f"never on today).")

# --------------------------------------- gate 3: ASIN suppression judgments
asin_judgments = {}
ambiguous = detect_asin_ambiguities(can, stock)
if ambiguous:
    st.subheader("ASIN judgment needed (suppression episode)")
    st.write("These ASINs hold stock under two or more SKUs right now. "
             "Decide per ASIN how to treat the stock **for this run** "
             "(spec §6b — sales are always pooled; this choice is about "
             "stock availability during suppression):")
    master = can.sku_master
    for asin in ambiguous:
        skus = master.loc[master["asin"] == asin, "sku_u"].tolist()
        choice = st.radio(
            f"**{asin}**  ({', '.join(skus)})",
            options=["combine", "only_active"],
            format_func=lambda v: (
                "Combine — all SKUs' stock is available (brief suppression)"
                if v == "combine" else
                "Only active — suppressed SKUs' stock held unavailable "
                "(long suppression)"),
            key=f"judg_{asin}", horizontal=False)
        asin_judgments[asin] = choice

# ------------------------------------------------------------ gate 4: plan
if not st.button("🧮 Calculate plan", type="primary"):
    st.stop()

res = run_calculation(can, stock, float(days_of_cover),
                      asin_judgments=asin_judgments or None)

m = res.meta
k1, k2, k3, k4 = st.columns(4)
k1.metric("Units planned", f"{m['units_planned_total']:,}")
k2.metric("Demand covered (units/window)",
          f"{m['demand_units_resolved']:.0f}")
k3.metric("Demand excluded", f"{m['demand_units_excluded']:.0f}")
k4.metric("Stockout warnings", len(res.stockouts))
for w in res.warnings:
    st.warning(w)
if m["demand_units_excluded"] > 0:
    st.error("Some demand was excluded (see warnings) — add the named "
             "state(s) to the Demand_Map tab and rerun to include it.")

st.subheader("Region priorities")
st.dataframe(pd.DataFrame([res.region_priorities]).T
             .rename(columns={0: "Priority"}), use_container_width=True)

st.subheader("Sales velocity vs plan (eyes-open view)")
st.dataframe(pd.DataFrame(m["velocity_summary"]).rename(columns={
    "region": "Region", "daily_velocity": "Daily velocity (u/day)",
    "planned_units": "Planned units",
    "days_cover_achieved": "Days cover achieved"}),
    use_container_width=True, hide_index=True)

st.subheader("Plan (rounded quantities)")
piv = (res.plan.pivot_table(index="sku_u", columns="region",
                            values="rounded_qty", aggfunc="sum",
                            fill_value=0))
order = [r for r in m["region_order"] if r in piv.columns]
st.dataframe(piv[order + [c for c in piv.columns if c not in order]],
             use_container_width=True, height=460)

with st.expander("Every line, every flag (Calculation details)"):
    st.dataframe(res.plan, use_container_width=True, height=420)
if len(res.stockouts):
    with st.expander(f"Stockout warnings on the rounded plan "
                     f"({len(res.stockouts)})"):
        st.dataframe(res.stockouts, use_container_width=True)
with st.expander("Run meta & flagged assumptions"):
    for k, v in m.items():
        st.write(f"**{k}**: {v}")

# ------------------------------------------------------- output workbook
out = io.BytesIO()
master_src = _buf(o_master) or "reference/inventory_plan_template.xlsx"
writer_warns = build_plan_workbook(res, can, master_src, out, stock=stock)
for w in writer_warns:
    st.warning(w)
st.download_button(
    "⬇️ Download plan workbook (Appointment plan + Calculation_Details)",
    data=out.getvalue(),
    file_name=("inventory_plan_"
               + pd.Timestamp(m["anchor_date"]).strftime("%d_%b_%Y")
               + ".xlsx"),
    type="primary")
