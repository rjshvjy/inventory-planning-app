"""
output_v2.py — writes the plan Excel for the v2 Inventory Planning App.

Consumes PlanResult (calculation_v2) + Canonical (ingestion_v2) and produces
a workbook by CLONING the master template's 'Appointment plan' sheet shape
(spec §4/§13): per-SKU rows x region columns, priorities row, date stamp —
plus a Calculation_Details sheet carrying every per-line component, flag,
warning, stockout and run-meta item so no adjustment is ever invisible.

Rules honoured:
- Quantities land on each SKU's own master row (master['row']); the writer
  NEVER re-orders or adds SKU rows — the master is the shape contract.
- Manual split regions (e.g. 'Pune (MH)') are left EMPTY (spec §5a: human
  fills by judgment). BLR4 IXD and SellerFlex (YSXA) columns stay empty —
  the plan never ships to the hub or own warehouse.
- Totals (Units Total / Bottles / Litres / Boxes) are computed from the
  written quantities and the master's per-unit columns.
- ASIN-pooled lines (assumption A4) are written on the ACTIVE SKU's row —
  same attribution the engine used.
"""

from __future__ import annotations

import warnings
from datetime import date

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

from ingestion_v2 import Canonical
from calculation_v2 import PlanResult

warnings.filterwarnings("ignore")

HEADER_ROW = 2
PRIORITY_ROW = 3
DETAILS_SHEET = "Calculation_Details"


def _header_map(ws) -> dict:
    return {str(ws.cell(HEADER_ROW, c).value).strip(): c
            for c in range(1, ws.max_column + 1)
            if ws.cell(HEADER_ROW, c).value}


def build_plan_workbook(res: PlanResult, can: Canonical,
                        template_path, out_path) -> list[str]:
    """Clone the master template and fill the plan. Returns warnings raised
    during writing (missing columns, unplaceable rows) — caller shows them."""
    warns: list[str] = []
    wb = load_workbook(template_path)
    ws = wb["Appointment plan"]
    hdr = _header_map(ws)

    # ---- date stamp: value cell right of the 'From:' label on row 1
    for c in range(1, ws.max_column + 1):
        if str(ws.cell(1, c).value or "").strip().rstrip(":") == "From":
            ws.cell(1, c + 1, res.meta.get("anchor_date") or str(date.today()))
            break

    # ---- region columns present in the template
    region_col = {r: hdr[r] for r in res.meta.get("region_order", [])
                  if r in hdr}
    for r in res.meta.get("region_order", []):
        if r not in hdr and r not in ("BLR4 IXD", "YSXA"):
            warns.append(f"Template has no column for region '{r}' — its "
                         f"quantities were NOT written. Add the column to "
                         f"the master (header row {HEADER_ROW}).")
    manual = set(res.meta.get("manual_split_regions", []))

    # ---- region priorities on row 3
    for region, pri in res.region_priorities.items():
        if region in region_col:
            cell = ws.cell(PRIORITY_ROW, region_col[region], pri)
            cell.font = Font(bold=True)

    # ---- quantities on each SKU's own master row
    sku_row = dict(zip(can.sku_master["sku_u"], can.sku_master["row"]))
    per_unit = can.sku_master.set_index("sku_u")[
        ["units_per_box", "bottles_per_unit", "litres_per_unit"]]
    plan = res.plan[res.plan["rounded_qty"] > 0]
    written: dict[str, float] = {}
    for _, line in plan.iterrows():
        sku, region = line["sku_u"], line["region"]
        if region in manual:
            continue                      # human fills split targets (§5a)
        r = sku_row.get(sku)
        c = region_col.get(region)
        if r is None or c is None:
            warns.append(f"Could not place {sku} x {region} "
                         f"({line['rounded_qty']:.0f}u) on the template.")
            continue
        ws.cell(r, c, int(line["rounded_qty"]))
        written[sku] = written.get(sku, 0) + int(line["rounded_qty"])

    # ---- totals from written quantities + master per-unit data
    for sku, total in written.items():
        r = sku_row[sku]
        pu = per_unit.loc[sku]
        if "Units Total" in hdr:
            ws.cell(r, hdr["Units Total"], int(total))
        if "Bottles" in hdr and pd.notna(pu["bottles_per_unit"]):
            ws.cell(r, hdr["Bottles"], round(total * pu["bottles_per_unit"], 1))
        if "Litres" in hdr and pd.notna(pu["litres_per_unit"]):
            ws.cell(r, hdr["Litres"], round(total * pu["litres_per_unit"], 2))
        if "Boxes" in hdr and pd.notna(pu["units_per_box"]) \
                and pu["units_per_box"]:
            ws.cell(r, hdr["Boxes"], round(total / pu["units_per_box"], 2))

    # ---- Calculation_Details sheet (full transparency)
    if DETAILS_SHEET in wb.sheetnames:
        del wb[DETAILS_SHEET]
    det = wb.create_sheet(DETAILS_SHEET)
    det.cell(1, 1, "CALCULATION DETAILS — every component, flag and warning "
                   "behind the plan. Auto-written; do not edit.")
    det.cell(1, 1).font = Font(bold=True)
    row = 3

    def block(title, df, cols=None):
        nonlocal row
        det.cell(row, 1, title).font = Font(bold=True)
        row += 1
        if df is None or len(df) == 0:
            det.cell(row, 1, "(none)")
            row += 2
            return
        use = cols or list(df.columns)
        for j, cname in enumerate(use, 1):
            det.cell(row, j, cname).font = Font(bold=True)
        row += 1
        for _, r_ in df.iterrows():
            for j, cname in enumerate(use, 1):
                v = r_[cname]
                det.cell(row, j, round(v, 3) if isinstance(v, float) else v)
            row += 1
        row += 1

    block("PLAN LINES", res.plan)
    block("STOCKOUT WARNINGS (on the rounded plan)", res.stockouts)
    block("RUN WARNINGS",
          pd.DataFrame({"warning": res.warnings})
          if res.warnings else pd.DataFrame())
    meta_df = pd.DataFrame([{"key": k, "value": str(v)}
                            for k, v in res.meta.items()])
    block("RUN META (incl. flagged assumptions)", meta_df)

    wb.save(out_path)
    return warns
