"""
calculation_v2.py — planning engine for the v2 Inventory Planning App.

Pure computation, no Streamlit: consumes the two validated dataclasses
(`Canonical` from ingestion_v2, `StockInputs` from workbook_builder_v2) plus
the user's days-of-cover and resolved ASIN suppression judgments, and returns
a PlanResult. UI wiring (inputs, prompts) lives in the app layer.

Spec anchors (v2.8): §5b demand-state->region resolver (lives HERE, not
ingestion); §7 single-source stock (StockInputs.current is the SOLE planning
stock base); §7a at-FC pending = near-term inbound, flagged, received never
re-counted; §8 IXD model; §9 daily-sales core, two-tier box rounding,
stockout sim on ROUNDED quantities, priorities retained verbatim.

All date arithmetic anchors on StockInputs.stock_as_of — never the run day.

FLAGGED ASSUMPTIONS (see also PlanResult.meta["assumptions"]):
  A1. Tier-2 slow-mover proposals use ceil(raw requirement) units because the
      master template has no inner-pack column yet (spec §4 lists it as
      optional). When the column is added, propose the inner-pack multiple.
  A2. Round-down threshold (near-empty, non-Pri-1) is a module constant
      (ROUNDDOWN_FRAC = 0.30) — spec §9 wants it in config but the Config tab
      does not carry a key for it yet. Promote to Config when convenient.
  A3. At-FC `pending` units are counted as available on day 0 of the horizon
      (they are physically at the FC; the At-the-FC tab carries no reach
      date). Every such line is flagged "AT_FC_PENDING(+n)".
  A4. Requirements are computed at ASIN level (the planning key) and written
      back to the master's canonical SKU rows. For a multi-SKU ASIN the
      requirement lands on the SKU row that currently holds sales in the
      window (the active SKU), falling back to the first master row.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field

import pandas as pd

from ingestion_v2 import (Canonical, Report, INDIAN_STATES, STATE_VARIANTS,
                          STATE_NAME_TO_CODE)
from workbook_builder_v2 import StockInputs, ordered_regions

# ---------------------------------------------------------------- constants

ROUNDDOWN_FRAC = 0.30          # assumption A2 — near-empty round-down fraction
IXD_REGION = "BLR4 IXD"
IGNORE_REGION = "YSXA"

PRI_1, PRI_2, PRI_3, PRI_4 = "Pri-1", "Pri-2", "Pri-3", "Pri-4"


# ------------------------------------------------------------------- result

@dataclass
class PlanResult:
    """Everything the output layer needs; no file writing happens here."""
    plan: pd.DataFrame | None = None
    # columns: sku_u, asin, region, daily_sales, current, in_transit_counted,
    #          at_fc_pending, ixd_offset, raw_requirement, rounded_qty, boxes,
    #          days_cover_achieved, priority, flags
    region_priorities: dict = field(default_factory=dict)   # region -> Pri-N
    stockouts: pd.DataFrame | None = None
    # columns: sku_u, region, day, deficit   (thresholded, on ROUNDED plan)
    warnings: list = field(default_factory=list)
    meta: dict = field(default_factory=dict)


# ------------------------------------------------- step 1+2: demand resolver

def build_code_to_region(can: Canonical) -> dict:
    """State code -> planning region label, from canonical.regions.
    Split states resolve to their DEFAULT region (spec §5a/§5b:
    MAHARASHTRA -> 'Bombay (MH)')."""
    default_label = {s["state"]: s["label"] for s in can.region_splits
                     if s["default"]}
    out = {}
    for lbl in can.regions:
        if lbl in (IXD_REGION, IGNORE_REGION):
            continue
        if "(" not in lbl:
            continue
        code = lbl.rsplit("(", 1)[1].rstrip(")").strip().upper()
        name = lbl.rsplit("(", 1)[0].strip()
        if code in default_label:                 # split state: default only
            if name == default_label[code]:
                out[code] = lbl
        else:
            out.setdefault(code, lbl)
    return out


def resolve_demand_regions(can: Canonical) -> tuple[dict, pd.DataFrame, list]:
    """Spec §5b resolution order. Returns (state->region map for every state
    seen in sales, unresolved sales rows, warning strings). Unresolved demand
    is EXCLUDED but loudly named with its unit count — never silent."""
    code_to_region = build_code_to_region(can)
    fc_state_names = {name for name, code in STATE_NAME_TO_CODE.items()
                      if code in code_to_region}

    state_region, warnings = {}, []
    seen = can.sales_daily["state"].astype(str).str.strip().str.upper()
    for raw in sorted(seen.unique()):
        state = STATE_VARIANTS.get(raw, raw)
        if state in fc_state_names:                      # FC state -> itself
            state_region[raw] = code_to_region[STATE_NAME_TO_CODE[state]]
        elif state in can.demand_map:                    # declared serving
            code = can.demand_map[state]
            if code in code_to_region:
                state_region[raw] = code_to_region[code]
            # (a serving code without FCs is already an ingestion error)

    mask = ~seen.isin(state_region)
    unresolved = can.sales_daily.loc[mask].copy()
    if len(unresolved):
        for st, qty in (unresolved.assign(state=seen[mask])
                        .groupby("state")["qty"].sum().items()):
            warnings.append(f"Demand state '{st}' is neither an FC state nor "
                            f"in Demand_Map — {qty:.0f} units EXCLUDED from "
                            f"planning (add a Demand_Map row to include it).")
    return state_region, unresolved, warnings


def pool_demand(can: Canonical, state_region: dict) -> pd.DataFrame:
    """sales_daily (sku_u x state) -> ASIN x region daily rates. Sales are
    ALWAYS clubbed by ASIN (spec §6b — no prompt)."""
    df = can.sales_daily.copy()
    df["state_u"] = df["state"].astype(str).str.strip().str.upper()
    df = df[df["state_u"].isin(state_region)]
    df["region"] = df["state_u"].map(state_region)
    sku_asin = dict(zip(can.sku_master["sku_u"], can.sku_master["asin"]))
    df["asin"] = df["sku_u"].map(sku_asin)
    g = (df.groupby(["asin", "region"], as_index=False)
           .agg(qty=("qty", "sum"), daily=("daily", "sum")))
    return g


# --------------------------------------------- ASIN pooling helper (spec §6b)

def detect_asin_ambiguities(can: Canonical, stock: StockInputs) -> list[str]:
    """ASINs holding stock under >=2 SKUs this run (suppression episode).
    The APP prompts the human per ASIN; run_calculation receives the resolved
    judgments. Exposed separately so the flow stays testable."""
    sku_asin = dict(zip(can.sku_master["sku_u"], can.sku_master["asin"]))
    cur = stock.current[stock.current["qty"] > 0].copy()
    cur["asin"] = cur["sku_u"].map(sku_asin)
    n_skus = cur.groupby("asin")["sku_u"].nunique()
    return sorted(n_skus[n_skus >= 2].index)


# --------------------------------------------------- step 3+4: supply assembly

def _horizon_days(lead: float, days_of_cover: float) -> float:
    return float(lead) + float(days_of_cover)


def assemble_supply(can: Canonical, stock: StockInputs, days_of_cover: float,
                    asin_judgments: dict, result: PlanResult
                    ) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    """Per-ASIN-per-region supply from the completed workbook (SOLE stock
    input, spec v2.5/v2.8): current (planning regions only) + in-transit
    lines whose committed reach_date lands inside the horizon + at-FC
    pending (day-0, flagged, assumption A3). Returns (supply df with
    component columns, in-transit arrival schedule for the stockout sim,
    ixd stock per ASIN)."""
    anchor = stock.stock_as_of
    sku_asin = dict(zip(can.sku_master["sku_u"], can.sku_master["asin"]))

    # ---- current stock: split planning regions / IXD / excluded-by-judgment
    cur = stock.current.copy()
    cur["asin"] = cur["sku_u"].map(sku_asin)

    excluded_rows = []
    if asin_judgments:
        active_sku = {}          # under "only_active": SKU with sales wins
        sales_by_sku = (can.sales_daily.groupby("sku_u")["qty"].sum())
        for asin, judgment in asin_judgments.items():
            if judgment != "only_active":
                continue
            skus = can.sku_master.loc[can.sku_master["asin"] == asin, "sku_u"]
            ranked = sorted(skus, key=lambda s: -sales_by_sku.get(s, 0))
            active_sku[asin] = ranked[0] if ranked else None
        for asin, act in active_sku.items():
            stranded = cur[(cur["asin"] == asin) & (cur["sku_u"] != act)
                           & (cur["qty"] > 0)]
            for _, r in stranded.iterrows():
                excluded_rows.append(r)
                result.warnings.append(
                    f"ASIN {asin}: {r['qty']:.0f} units under suppressed SKU "
                    f"{r['sku_u']} at {r['region']} held UNAVAILABLE this run "
                    f"(judgment: only_active).")
            cur = cur.drop(stranded.index)

    ixd_stock = (cur[cur["region"] == IXD_REGION]
                 .groupby("asin")["qty"].sum().to_dict())
    cur_plan = cur[cur["region"].isin(can.regions)]
    current = (cur_plan.groupby(["asin", "region"], as_index=False)
               .agg(current=("qty", "sum")))

    # ---- in-transit: count a line iff reach_date <= anchor + horizon(region)
    it_counted, it_schedule = [], []
    it = stock.in_transit if stock.in_transit is not None else pd.DataFrame()
    for _, r in it.iterrows():
        region = r["region"]
        if region == IGNORE_REGION:
            continue
        lead = stock.lead_days.get(region, None)
        if lead is None and region != IXD_REGION:
            continue                      # reader already gated unknown dests
        horizon = _horizon_days(lead if lead is not None else 0, days_of_cover)
        day = (pd.Timestamp(r["reach_date"]).normalize()
               - anchor.normalize()).days
        asin = sku_asin.get(r["sku_u"])
        if region == IXD_REGION:
            # inbound TO the hub: joins the IXD pool on arrival; modelled by
            # adding to ixd stock if it lands within ixd_transfer window
            if day <= can.config["ixd_transfer_days"] + days_of_cover:
                ixd_stock[asin] = ixd_stock.get(asin, 0) + float(r["qty"])
            continue
        if day <= horizon:
            it_counted.append({"asin": asin, "region": region,
                               "qty": float(r["qty"])})
            it_schedule.append({"asin": asin, "region": region,
                                "day": max(0, day), "qty": float(r["qty"])})
    in_transit = (pd.DataFrame(it_counted)
                  .groupby(["asin", "region"], as_index=False)["qty"].sum()
                  .rename(columns={"qty": "in_transit_counted"})
                  if it_counted else
                  pd.DataFrame(columns=["asin", "region",
                                        "in_transit_counted"]))

    # ---- at-FC pending: near-term inbound, day 0, flagged (§7a, A3)
    af_rows = []
    af = stock.at_fc if stock.at_fc is not None else pd.DataFrame()
    for _, r in af.iterrows():
        pend = float(r.get("pending") or 0)
        if pend <= 0 or r["region"] in (IGNORE_REGION,):
            continue
        asin = sku_asin.get(r["sku_u"])
        if r["region"] == IXD_REGION:
            ixd_stock[asin] = ixd_stock.get(asin, 0) + pend
            continue
        af_rows.append({"asin": asin, "region": r["region"], "qty": pend})
    at_fc = (pd.DataFrame(af_rows)
             .groupby(["asin", "region"], as_index=False)["qty"].sum()
             .rename(columns={"qty": "at_fc_pending"})
             if af_rows else
             pd.DataFrame(columns=["asin", "region", "at_fc_pending"]))

    supply = current.merge(in_transit, on=["asin", "region"], how="outer") \
                    .merge(at_fc, on=["asin", "region"], how="outer")
    for c in ("current", "in_transit_counted", "at_fc_pending"):
        if c not in supply:
            supply[c] = 0.0
    supply = supply.fillna({"current": 0, "in_transit_counted": 0,
                            "at_fc_pending": 0})
    schedule = (pd.DataFrame(it_schedule) if it_schedule else
                pd.DataFrame(columns=["asin", "region", "day", "qty"]))
    return supply, schedule, ixd_stock


# ------------------------------------------------------------ step 5: IXD (§8)

def apply_ixd(demand: pd.DataFrame, ixd_stock: dict, cfg: dict
              ) -> pd.DataFrame:
    """Proportional-to-daily-sales split of each ASIN's IXD stock across its
    demand regions, scaled by ixd_confidence. Returns asin x region
    ixd_offset. Arrival day for the sim = ixd_transfer_days."""
    rows = []
    conf = float(cfg["ixd_confidence"])
    for asin, qty in ixd_stock.items():
        if qty <= 0 or conf <= 0 or asin is None:
            continue
        d = demand[demand["asin"] == asin]
        tot = d["daily"].sum()
        if tot <= 0:
            continue                       # no demand anywhere: keep at hub
        for _, r in d.iterrows():
            rows.append({"asin": asin, "region": r["region"],
                         "ixd_offset": qty * conf * r["daily"] / tot})
    return (pd.DataFrame(rows) if rows else
            pd.DataFrame(columns=["asin", "region", "ixd_offset"]))


# ---------------------------------------------- step 6+7: requirement + boxes

def raw_requirements(demand: pd.DataFrame, supply: pd.DataFrame,
                     ixd: pd.DataFrame, lead_days: dict,
                     days_of_cover: float) -> pd.DataFrame:
    df = demand.merge(supply, on=["asin", "region"], how="outer") \
               .merge(ixd, on=["asin", "region"], how="left")
    for c in ("daily", "current", "in_transit_counted", "at_fc_pending",
              "ixd_offset"):
        if c not in df:
            df[c] = 0.0
        df[c] = df[c].fillna(0.0)
    df["lead"] = df["region"].map(lead_days)
    df = df[df["lead"].notna()]            # only regions with a lead time
    depletion = df["daily"] * (df["lead"] + days_of_cover)
    avail = (df["current"] + df["in_transit_counted"] + df["at_fc_pending"]
             + df["ixd_offset"])
    df["raw_requirement"] = (depletion - avail).clip(lower=0.0)
    return df


def round_boxes(df: pd.DataFrame, master: pd.DataFrame, cfg: dict,
                days_of_cover: float) -> pd.DataFrame:
    """Two-tier rounding (§9). Flags: FULL_CARTON / PARTIAL_BOX (Tier-2,
    assumption A1) / ROUNDED_DOWN (near-empty, A2) / AT_FC_PENDING /
    IXD_OFFSET annotations added here so every adjustment is visible."""
    upb = dict(zip(master["asin"], master["units_per_box"]))
    out = df.copy()
    rounded, boxes, flags = [], [], []
    ceiling = float(cfg["days_cover_ceiling"])
    for _, r in out.iterrows():
        f = []
        req, daily = r["raw_requirement"], r["daily"]
        box = float(upb.get(r["asin"]) or 1)
        if r["at_fc_pending"] > 0:
            f.append(f"AT_FC_PENDING(+{r['at_fc_pending']:.0f})")
        if r["ixd_offset"] > 0:
            f.append(f"IXD_OFFSET(-{r['ixd_offset']:.1f})")
        if req <= 0:
            rounded.append(0); boxes.append(0.0); flags.append(";".join(f))
            continue
        carton_days = (box / daily) if daily > 0 else math.inf
        if daily > 0 and carton_days > ceiling:
            q = math.ceil(req)                      # Tier 2 (A1: no inner-pack
            f.append(f"PARTIAL_BOX({q} of {box:.0f}; full carton = "
                     f"{carton_days:.0f}d cover)")  # column in master yet)
        else:
            q = math.ceil(req / box) * box          # Tier 1: full cartons
            overshoot = q - req
            if (overshoot > 0 and req < ROUNDDOWN_FRAC * box
                    and r.get("priority_pre") != PRI_1):
                q = 0
                f.append(f"ROUNDED_DOWN(need {req:.1f} < "
                         f"{ROUNDDOWN_FRAC:.0%} of a {box:.0f}-box)")
            else:
                f.append("FULL_CARTON")
        rounded.append(int(q))
        boxes.append(round(q / box, 2) if box else 0.0)
        flags.append(";".join(f))
    out["rounded_qty"], out["boxes"], out["flags"] = rounded, boxes, flags
    out["days_cover_achieved"] = out.apply(
        lambda r: ((r["current"] + r["in_transit_counted"]
                    + r["at_fc_pending"] + r["ixd_offset"] + r["rounded_qty"])
                   / r["daily"] - r["lead"]) if r["daily"] > 0 else float("inf"),
        axis=1)
    return out


# --------------------------------------------------- step 8: stockout sim (§9)

def simulate_stockouts(plan: pd.DataFrame, schedule: pd.DataFrame,
                       cfg: dict, days_of_cover: float) -> pd.DataFrame:
    """Day-walk on ROUNDED quantities. Arrivals: in-transit on its committed
    day, at-FC pending on day 0, IXD offset on ixd_transfer_days, the planned
    shipment on the region's lead day. Deficits below
    min_stockout_deficit_units are suppressed."""
    thresh = float(cfg["min_stockout_deficit_units"])
    ixd_day = int(cfg["ixd_transfer_days"])
    hits = []
    sched = {(r["asin"], r["region"]): [] for _, r in schedule.iterrows()}
    for _, r in schedule.iterrows():
        sched[(r["asin"], r["region"])].append((int(r["day"]), r["qty"]))
    for _, r in plan.iterrows():
        daily = r["daily"]
        if daily <= 0:
            continue
        lead = int(r["lead"])
        horizon = int(math.ceil(lead + days_of_cover))
        stock = r["current"] + r["at_fc_pending"]        # day-0 components
        arrivals = dict(sched.get((r["asin"], r["region"]), []))
        for day in range(horizon + 1):
            if day in arrivals:
                stock += arrivals[day]
            if day == ixd_day and r["ixd_offset"] > 0:
                stock += r["ixd_offset"]
            if day == lead and r["rounded_qty"] > 0:
                stock += r["rounded_qty"]
            stock -= daily
            if stock < -thresh:
                hits.append({"asin": r["asin"], "region": r["region"],
                             "day": day, "deficit": round(-stock, 1)})
                break
    return (pd.DataFrame(hits) if hits else
            pd.DataFrame(columns=["asin", "region", "day", "deficit"]))


# ------------------------------------------------------ step 9: priorities

def compute_priorities(plan: pd.DataFrame) -> tuple[pd.Series, dict]:
    """Retained verbatim from the old app (spec §9): risk = lead −
    days-of-stock; Pri-1 if risk >= 2 or stock == 0; Pri-2 if >= 0;
    Pri-3 if >= −3; else Pri-4. Region priority from mean risk with the
    >=1 / >=−1 / >=−3 thresholds."""
    pri, risk_by_region = [], {}
    for _, r in plan.iterrows():
        if r["daily"] <= 0:
            pri.append(PRI_4)
            continue
        days_of_stock = r["current"] / r["daily"]
        risk = r["lead"] - days_of_stock
        if risk >= 2 or r["current"] == 0:
            pri.append(PRI_1)
        elif risk >= 0:
            pri.append(PRI_2)
        elif risk >= -3:
            pri.append(PRI_3)
        else:
            pri.append(PRI_4)
        risk_by_region.setdefault(r["region"], []).append(risk)
    region_pri = {}
    for region, risks in risk_by_region.items():
        avg = sum(risks) / len(risks)
        region_pri[region] = (PRI_1 if avg >= 1 else PRI_2 if avg >= -1
                              else PRI_3 if avg >= -3 else PRI_4)
    return pd.Series(pri, index=plan.index), region_pri


# ------------------------------------------------------------------ pipeline

def run_calculation(can: Canonical, stock: StockInputs, days_of_cover: float,
                    asin_judgments: dict | None = None,
                    today: pd.Timestamp | None = None) -> PlanResult:
    """Top-level entry (contract per handover §9 / spec v2.8). The app must
    only call this after ingestion AND workbook-read both pass (rep.ok)."""
    result = PlanResult()
    asin_judgments = asin_judgments or {}

    # unresolved ASIN ambiguities are loud, defaulted to combine, and flagged
    for asin in detect_asin_ambiguities(can, stock):
        if asin not in asin_judgments:
            result.warnings.append(
                f"ASIN {asin} holds stock under 2+ SKUs but no suppression "
                f"judgment was provided — defaulting to COMBINE this run "
                f"(the app should prompt; spec §6b).")
            asin_judgments[asin] = "combine"

    # 1-2: demand
    state_region, unresolved, warns = resolve_demand_regions(can)
    result.warnings.extend(warns)
    demand = pool_demand(can, state_region)

    # 3-4: supply
    supply, schedule, ixd_stock = assemble_supply(
        can, stock, days_of_cover, asin_judgments, result)

    # 5: IXD
    ixd = apply_ixd(demand, ixd_stock, can.config)

    # 6: raw requirements
    df = raw_requirements(demand, supply, ixd, stock.lead_days, days_of_cover)

    # 9a: pre-rounding priorities (round-down guard needs Pri-1 knowledge)
    df["priority_pre"], _ = compute_priorities(df)

    # 7: rounding
    df = round_boxes(df, can.sku_master, can.config, days_of_cover)

    # 8: stockout sim on rounded quantities
    result.stockouts = simulate_stockouts(df, schedule, can.config,
                                          days_of_cover)

    # 9b: final priorities (on the same pre-shipment risk, per old app)
    df["priority"], result.region_priorities = compute_priorities(df)

    # A4: attribute each ASIN row back to its active canonical SKU
    sales_by_sku = can.sales_daily.groupby("sku_u")["qty"].sum()
    asin_rows = can.sku_master.groupby("asin")["sku_u"].apply(list).to_dict()

    def active_sku(asin):
        skus = asin_rows.get(asin, [])
        ranked = sorted(skus, key=lambda s: -sales_by_sku.get(s, 0))
        return ranked[0] if ranked else None
    df["sku_u"] = df["asin"].map(active_sku)

    # split states: ensure the non-default (manual) region appears with zeros
    manual_regions = [f"{s['label']} ({s['state']})" for s in can.region_splits
                      if not s["default"]]
    ordered = ordered_regions(can)
    plan_cols = ["sku_u", "asin", "region", "daily", "current",
                 "in_transit_counted", "at_fc_pending", "ixd_offset",
                 "raw_requirement", "rounded_qty", "boxes",
                 "days_cover_achieved", "priority", "flags"]
    df = df.rename(columns={"daily": "daily"})
    result.plan = (df[plan_cols]
                   .sort_values(["sku_u", "region"])
                   .reset_index(drop=True))

    # v2.9 §5c: flag every plan line in a region that received user-mapped
    # FC stock this run (per-run mapping, surfaced in output, never silent)
    fcmap_regions = {}
    for fc, r_ in can.fc_resolutions.items():
        if r_.get("action") == "map" and r_.get("fulfillable", True):
            fcmap_regions.setdefault(r_["region"], []).append(fc)
    if fcmap_regions:
        result.plan["flags"] = result.plan.apply(
            lambda x: ";".join(filter(None, [x["flags"]] + [
                f"FC_MAPPED({fc}->{x['region']})"
                for fc in fcmap_regions.get(x["region"], [])])), axis=1)
        for region, fcs in fcmap_regions.items():
            result.warnings.append(
                f"Region {region} includes stock from user-mapped FC(s) "
                f"{fcs} this run (not in registration) — lines flagged "
                f"FC_MAPPED.")
    for fc, r_ in can.fc_resolutions.items():
        if r_.get("action") == "map" and not r_.get("fulfillable", True):
            result.warnings.append(
                f"FC {fc} stock mapped to {r_['region']} but marked NOT "
                f"fulfillable — shown, excluded from planning supply.")

    # v2.9 change 4: per-region + overall velocity summary (drives the
    # Sales-velocity output tab; presentation of existing math, no new math)
    vel = (result.plan.groupby("region")
           .agg(daily_velocity=("daily", "sum"),
                planned_units=("rounded_qty", "sum")).reset_index())
    vel["days_cover_achieved"] = (
        vel["planned_units"] / vel["daily_velocity"]).where(
        vel["daily_velocity"] > 0).round(1)
    overall = {"region": "OVERALL",
               "daily_velocity": round(float(vel["daily_velocity"].sum()), 2),
               "planned_units": int(vel["planned_units"].sum()),
               "days_cover_achieved": None}
    velocity_summary = vel.round({"daily_velocity": 2}).to_dict("records")         + [overall]

    result.meta = {
        "velocity_summary": velocity_summary,
        "fc_resolutions": dict(can.fc_resolutions),
        "anchor_date": str(stock.stock_as_of.date()) if stock.stock_as_of
                       else None,
        "days_of_cover": days_of_cover,
        "sales_window_days": can.sales_window_days,
        "config": dict(can.config),
        "lead_days": dict(stock.lead_days),
        "asin_judgments": dict(asin_judgments),
        "manual_split_regions": manual_regions,   # start empty in the output
        "region_order": ordered,
        "units_planned_total": int(result.plan["rounded_qty"].sum()),
        "demand_units_resolved": float(demand["qty"].sum()),
        "demand_units_excluded": float(unresolved["qty"].sum())
                                 if len(unresolved) else 0.0,
        "assumptions": ["A1 Tier-2 uses ceil(raw) — no inner-pack column yet",
                        "A2 round-down fraction is module constant 0.30",
                        "A3 at-FC pending counted at day 0 (no reach date)",
                        "A4 ASIN requirement written to the active SKU row"],
    }
    return result
