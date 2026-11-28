"""Capstone volume Section Examples workbook (modules P.1 - P.9): lifecycle, source register, the three capstone builds,
QA, defence, portfolio and interviews - all worked on the bootcamp datasets."""
from __future__ import annotations

import hashlib
import os
import sqlite3
import sys

import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(__file__))
from common import CORE, SEC, STAGE, W2D, W3D  # noqa: E402
from wk_common import DATASETS_TEXT, L, lc_adm, sm_sales  # noqa: E402
from xlhelp import Book  # noqa: E402

OUT = SEC / "ETEH_Capstone_Section_Examples.xlsx"
XL0 = pd.Timestamp("1899-12-30")

sm = sm_sales()
sm["Year"] = sm.Order_Date.dt.year
sm["Quarter"] = "Q" + sm.Order_Date.dt.quarter.astype(str)
a24, b25 = sm[sm.Year == 2024], sm[sm.Year == 2025]
stores = L("sm_stores")
cv = L("cv_perf")
lc = lc_adm()
census = L("lc_census").copy()
census["Census_Date"] = pd.to_datetime(census.Census_Date)
census["Quarter"] = census.Census_Date.dt.year.astype(str) + "-Q" + census.Census_Date.dt.quarter.astype(str)
dirty = L("dirty_sm")
dlog = L("dirty_sm_log")
ctrl = L("dirty_sm_ctrl").set_index("Control total (after cleaning)").Value

b = Book(str(OUT), "ETEH Capstone - Section Examples", "Volume 5 · Capstone, QA, Defence and Career")
F = b.F

g25 = b25.Sales.sum() / a24.Sales.sum() - 1
m24, m25 = a24.Profit.sum() / a24.Sales.sum(), b25.Profit.sum() / b25.Sales.sum()

# ======================================================================================
# P.1 Lifecycle
# ======================================================================================
s = b.sheet("P.1 Project Lifecycle", "P.1  The Seven-Stage Project Lifecycle - applied to Capstone 1 (SuperMart)",
            "Brief, provenance, a reconciled cleaning log, quality gates, a time box and the post-project review", "P.1.1 - P.1.12")
s.band("P.1.2", "Stage 1 - Scoping: the one-page brief")
s.grid([["Decision it informs", "Which pricing and discount actions SuperMart takes before Q1 2026"],
        ["Question", f"Why did FY2025 sales grow {g25:.1%} while gross margin fell from {m24:.1%} to {m25:.1%}?"],
        ["Measures", "Sales, Profit, Margin %, Discount depth, Orders, Basket value, Target attainment"],
        ["Dimensions", "Store, Zone, Category, Sub-category, Quarter, Customer segment"],
        ["Period", "FY2024 vs FY2025 (calendar years), lines dated 2024-01-01 to 2025-12-31"],
        ["Audience / deadline", "SuperMart board; draft in 2 weeks, final pack in 3 weeks"],
        ["Sign-off", "Commercial Director (requester) + Head of Finance (control totals)"]], header=["Brief item", "Capstone 1 entry"], merge_last=True)
s.band("P.1.3", "Stage 2 - Acquisition: provenance recorded before touching the data")
s.grid([["SuperMart_Stores_Sales_2022_2025.csv", "POS ledger export", f"{len(sm):,}", "2026-01-05", "Finance - Revenue Assurance", "01_RAW/ (read-only)"],
        ["SuperMart_Stores_Directory.csv", "Store master", f"{len(stores)}", "2026-01-05", "Operations", "01_RAW/"],
        ["SuperMart_Sales_2025_DIRTY.csv", "Manual store returns (Week 2)", f"{len(dirty):,}", "2026-01-06", "Store admin team", "01_RAW/"]],
       header=["File", "Source system", "Rows", "Extracted", "Owner", "Stored in"])
s.band("P.1.4", "Stage 3 - Cleaning: the row-count reconciliation for the dirty 2025 extract")
blank = int(dirty.isna().all(axis=1).sum()) if dirty.isna().all(axis=1).any() else int(dirty.Order_ID.isna().sum())
dups = len(dirty) - blank - int(ctrl["Clean row count"])
cl = pd.DataFrame({"Step": ["Rows received", "Remove blank spacer rows", "Remove exact duplicates", "Rows after cleaning"],
                   "Rows_removed": [None, blank, dups, None]})
top = s.r + 1
fx = {(0, "Rows"): ("=" + str(len(dirty)), len(dirty)),
      (1, "Rows"): (f"=D{top + 1}-C{top + 2}", len(dirty) - blank),
      (2, "Rows"): (f"=D{top + 2}-C{top + 3}", len(dirty) - blank - dups),
      (3, "Rows"): (f"=D{top + 3}", int(ctrl["Clean row count"]))}
cl["Rows"] = 0
s.frame(cl, fmts={"Rows_removed": "int", "Rows": "int"}, formulas=fx)
s.calc([("Control total from the POS ledger (FY2025 lines)", int(ctrl["Clean row count"]), int(ctrl["Clean row count"]), "int", "Independent figure from Finance"),
        ("Difference", f"=D{top + 4}-{int(ctrl['Clean row count'])}", 0, "int", "Must be 0 - Gate 3 passes"),
        ("Sales after cleaning = ledger?", f"={int(ctrl['Total Sales (₦)'])}-SUMIFS(tSM[Sales],tSM[Year],2025)", 0, "naira", "₦0 difference after fixing text numbers and ×100 outliers")])
s.band("P.1.5-9", "Stages 4-7 and the quality gates")
s.grid([["Gate 1 - Scope", "Brief signed by requester", "Signed brief (above)"],
        ["Gate 2 - Provenance", "Every file has source, date, rows, owner", "Provenance table (above)"],
        ["Gate 3 - Cleaning", "Rows after cleaning = control total", f"{int(ctrl['Clean row count']):,} = {int(ctrl['Clean row count']):,} ✔"],
        ["Gate 4 - Model", "Every relationship many-to-one, date table marked", "Model diagram screenshot"],
        ["Gate 5 - Visuals", "Every visual answers a question in the brief", "Visual-to-question map"],
        ["Gate 6 - Verification", f"Dashboard total = ₦{b25.Sales.sum():,.0f}", "Reconciliation cell on dashboard"],
        ["Gate 7 - Handover", "README, cleaning log, measures, refresh steps", "00_README sheet"]], header=["Gate", "Pass criterion", "Evidence"])
s.band("P.1.10", "Time-boxing a 40-hour capstone")
hrs = s.inputs([("Total hours available", 40, "int", "Change it - the plan re-allocates")])["Total hours available"]
tb = pd.DataFrame({"Stage": ["Scoping", "Acquisition", "Cleaning", "Modelling", "Visualisation", "Verification", "Handover"],
                   "Share": [.10, .10, .30, .20, .15, .10, .05]})
s.table(tb, "tTimeBox", fmts={"Share": "pct0"}, calc={"Hours": (f"=[@Share]*{hrs}", "dec1", (tb.Share * 40).tolist())}, total={"Stage": "Total", "Share": "sum", "Hours": "sum"})
s.band("P.1.11-12", "Stakeholders and the post-project review")
s.grid([["Week 1", "One-page scope to Commercial Director", "Sign-off on question and period"],
        ["Week 2", "Draft dashboard preview to Finance", "Agree control totals before polish"],
        ["Week 3", "Final pack + QA note to board", "Decision on Q4 discount policy"]], header=["When", "What", "Purpose"])
s.grid([["Did it change a decision?", "Board capped Furniture discounts at 20% from Q1 2026"], ["What took longer?", "Cleaning the 331 mixed-format dates in the store returns"],
        ["Automate next time", "Monthly folder import (Week 2, 2.3) instead of manual copy"], ["Raise with data owner", "Stores typing '₦' into the Sales column of the returns template"]],
       header=["Review question", "Answer recorded"], merge_last=True)

# ======================================================================================
# P.2 Source register (rewritten - internal sources only)
# ======================================================================================
s = b.sheet("P.2 Source Register", "P.2  Source Register, Before/After Evidence and Reproducibility",
            "Every bootcamp dataset registered with owner, period, rows and a file fingerprint; a before/after sample from the dirty extract", "P.2.1 - P.2.5")
s.band("P.2.1", "Beginner: source register essentials")
cat = pd.read_csv(STAGE / "catalogue.csv")
files = {"SuperMart_Stores_Sales_2022_2025.csv": "Finance - Revenue Assurance", "SuperMart_Stores_Directory.csv": "Operations",
         "SuperMart_Online_transactions.csv": "SuperMart Online - Data Engineering", "Mercyland_Inpatient_Episodes_2019_2024.csv": "Mercyland - Health Records",
         "Mercyland_Daily_Bed_Census.csv": "Mercyland - Bed Management", "Mercyland_Diabetes_Screening.csv": "Mercyland - Screening Clinic",
         "Kingsway_Student_Results.csv": "Kingsway - Registry", "Kingsway_Student_Welfare_Survey_SENSITIVE.csv": "Kingsway - Student Affairs (restricted)",
         "Kingsway_Digital_Habits_Survey.csv": "Kingsway - Student Affairs", "ETEH_Talent_Salary_Survey.csv": "ETEH Talent - Research",
         "AtlanticPearl_Reservations_2023_2024.csv": "Atlantic Pearl - Revenue Management"}
reg = []
for fn, owner in files.items():
    p = CORE / fn
    raw = p.read_bytes()
    nrows = raw.count(b"\n") - 1
    reg.append([fn, owner, nrows, round(len(raw) / 1024), hashlib.sha256(raw).hexdigest()[:16], "2026-11-27", "Internal training copy - not for public redistribution"])
regd = pd.DataFrame(reg, columns=["File", "Owner", "Rows", "Size_KB", "SHA256_first16", "Registered", "Licence / use"])
s.table(regd, "tRegister", fmts={"Rows": "int", "Size_KB": "int"},
        widths={"File": 40, "Owner": 34, "SHA256_first16": 20, "Licence / use": 40})
s.para("The SHA-256 fingerprint proves the file you analyse is the file that was registered: recompute it (PowerShell: Get-FileHash file.csv) and compare. "
       "If it differs, someone changed the RAW file.", "note")
s.h2("Where each dataset came from (lineage)")
s.table(cat[["Dataset", "Organisation", "Rows", "Period", "Built_From"]], "tLineage", fmts={"Rows": "int"},
        widths={"Dataset": 40, "Organisation": 30, "Period": 24, "Built_From": 70})
s.band("P.2.2", "Beginner: preserve before and after states - 8 lines from the dirty 2025 extract")
smi = sm.set_index("Row_ID")
picks = []
dd = dirty.dropna(subset=["Row_ID"]).copy()
for col, test in [("Sales", lambda v: isinstance(v, str)), ("Discount", lambda v: isinstance(v, str)), ("Category", lambda v: isinstance(v, str) and v not in ("Furniture", "Office Supplies", "Technology")),
                  ("Order_Date", lambda v: isinstance(v, str) and not str(v)[:4].isdigit()), ("Store_Name", lambda v: isinstance(v, str) and (v.isupper() or v.islower())),
                  ("Zone", lambda v: isinstance(v, str) and v != v.strip()), ("Customer_Name", lambda v: v == "N/A"), ("Quantity", lambda v: isinstance(v, (int, float)) and not pd.isna(v) and v < 0)]:
    hit = dd[dd[col].map(test)].head(1)
    if len(hit):
        rid = int(hit.Row_ID.iloc[0])
        after = smi.loc[rid, col] if col != "Order_Date" else smi.loc[rid, "Order_Date"].strftime("%Y-%m-%d")
        if col == "Customer_Name":
            after = "(blank - flagged for store to confirm)"
        picks.append([rid, col, repr(hit[col].iloc[0]) if col == "Zone" else str(hit[col].iloc[0]), str(after)])
ba = pd.DataFrame(picks, columns=["Row_ID", "Field", "BEFORE (as received)", "AFTER (cleaned)"])
ba = ba[["Row_ID", "Field", "BEFORE (as received)", "AFTER (cleaned)"]]
s.table(ba, "tBeforeAfter", fmts={"Row_ID": "gen"}, widths={"Field": 14, "BEFORE (as received)": 30, "AFTER (cleaned)": 34})
s.para("Store the BEFORE file untouched in 01_RAW; the AFTER table is produced only by the named Power Query steps, so anyone can regenerate it.", "note")
s.band("P.2.3", "Intermediate: citation, licence and privacy")
s.grid([["Cite internal sources precisely", "'SuperMart POS ledger export, SuperMart_Stores_Sales_2022_2025.csv, extracted 05-Jan-2026, Finance'"],
        ["Licence / permitted use", "Training copies are for the bootcamp only; ask the data owner before reusing any file in a public portfolio"],
        ["Privacy", "Customer, patient and respondent names are fictional, but treat them as real: never publish row-level data"],
        ["Sensitive files", "Kingsway_Student_Welfare_Survey_SENSITIVE.csv: aggregate tables only, minimum cell size 5"]], header=["Rule", "Applied to the bootcamp data"], merge_last=True)
s.band("P.2.4-5", "Reproducibility, review and the evidence-to-decision handoff")
s.grid([["Named steps", "Every Power Query step renamed (e.g. 'Removed spacer rows', 'Fixed ₦ text')"],
        ["Assumptions stored", "Assumptions sheet: FY = calendar year; net sales after discount excl. VAT"],
        ["Refresh test", "Replace the RAW file with the next month's export → Refresh All → control totals still reconcile"],
        ["Second-person test", "A colleague follows the README on another PC and gets ₦733,215,255"],
        ["Handoff pack", "Source register + before/after sample + workbook/SQL + QA checks + findings + uncertainty + owner + follow-up date"],
        ["Facts vs recommendations", "Facts: 'Q4 margin fell to 9.8%'. Recommendation: 'Cap Furniture discounts at 20%' - kept in separate sections"]],
       header=["Practice", "Example"], merge_last=True)

# ======================================================================================
# P.3 Capstone 1 SuperMart
# ======================================================================================
s = b.sheet("P.3 Capstone1 SuperMart", "P.3  Capstone 1 - SuperMart Retail: growth with a shrinking margin",
            "Brief, inventory, model measures, basket analysis, margin erosion, slow movers, memo and verification", "P.3.1 - P.3.10")
s.band("P.3.1", "Brief and success criteria - the headline numbers (live formulas over tSM)")
s.calc([("FY2024 sales", "=SUMIFS(tSM[Sales],tSM[Year],2024)", a24.Sales.sum(), "naira", ""),
        ("FY2025 sales", "=SUMIFS(tSM[Sales],tSM[Year],2025)", b25.Sales.sum(), "naira", ""),
        ("Sales growth", "=SUMIFS(tSM[Sales],tSM[Year],2025)/SUMIFS(tSM[Sales],tSM[Year],2024)-1", g25, "pct", "The good news"),
        ("Margin FY2024", "=SUMIFS(tSM[Profit],tSM[Year],2024)/SUMIFS(tSM[Sales],tSM[Year],2024)", m24, "pct", ""),
        ("Margin FY2025", "=SUMIFS(tSM[Profit],tSM[Year],2025)/SUMIFS(tSM[Sales],tSM[Year],2025)", m25, "pct", ""),
        ("Margin change (points)", "=(SUMIFS(tSM[Profit],tSM[Year],2025)/SUMIFS(tSM[Sales],tSM[Year],2025)-SUMIFS(tSM[Profit],tSM[Year],2024)/SUMIFS(tSM[Sales],tSM[Year],2024))*100",
         (m25 - m24) * 100, "dec2", "The board's question: why is margin falling while sales grow?")])
s.band("P.3.2-3", "Data inventory and cleaning")
s.grid([["POS sales lines", "SuperMart_Stores_Sales_2022_2025.csv", f"{len(sm):,} lines, {sm.Order_ID.nunique():,} orders, 2022-2025"],
        ["Products (derived)", "Distinct Product_ID in the ledger", f"{sm.Product_ID.nunique():,} SKUs in 3 categories, 17 sub-categories"],
        ["Stores", "SuperMart_Stores_Directory.csv", f"{len(stores)} stores, 4 zones, FY2025 targets"],
        ["Calendar", "Date table built in Power Query", "2022-01-01 to 2025-12-31"],
        ["Store returns (dirty)", "SuperMart_Sales_2025_DIRTY.csv", f"{len(dirty):,} rows → {int(ctrl['Clean row count']):,} after cleaning (see P.1, P.2)"]],
       header=["Source", "File", "Scale"])
s.band("P.3.4", "Modelling: the measures and their FY2025 values")
o25 = b25.groupby("Order_ID").agg(lines=("Row_ID", "count"), val=("Sales", "sum"), cats=("Category", "nunique"))
meas = [["Revenue", "SUM ( FactSales[Sales] )", b25.Sales.sum(), "naira"], ["Profit", "SUM ( FactSales[Profit] )", b25.Profit.sum(), "naira"],
        ["Gross Margin %", "DIVIDE ( [Profit], [Revenue] )", m25, "pct"], ["Orders", "DISTINCTCOUNT ( FactSales[Order_ID] )", len(o25), "int"],
        ["Basket Value", "DIVIDE ( [Revenue], [Orders] )", b25.Sales.sum() / len(o25), "naira"],
        ["Discount Depth", "DIVIDE ( SUMX ( FactSales, FactSales[Discount] * FactSales[Sales] ), [Revenue] )", (b25.Discount * b25.Sales).sum() / b25.Sales.sum(), "pct"],
        ["Revenue LY", "CALCULATE ( [Revenue], SAMEPERIODLASTYEAR ( DimDate[Date] ) )", a24.Sales.sum(), "naira"],
        ["Growth %", "DIVIDE ( [Revenue] - [Revenue LY], [Revenue LY] )", g25, "pct"],
        ["Target Attainment", "DIVIDE ( [Revenue], SUM ( DimStore[Annual_Target_2025] ) )", b25.Sales.sum() / stores.Annual_Target_2025.sum(), "pct"],
        ["Share of Total", "DIVIDE ( [Revenue], CALCULATE ( [Revenue], ALL ( DimStore ) ) )", None, None]]
md = pd.DataFrame([[a, b_, "" if c is None else c] for a, b_, c, d in meas], columns=["Measure", "DAX", "FY2025 value"])
tm = s.table(md, "tCap1Measures", widths={"Measure": 18, "DAX": 70, "FY2025 value": 18})
for i, r in enumerate(meas):
    if r[2] is not None:
        s.ws.write(tm.r0 + 1 + i, tm.ci("FY2025 value"), float(r[2]), b.nf(r[3], style="plain"))
s.band("P.3.5", "Basket analysis (FY2025 orders)")
s.calc([("Orders", "=ROWS(tBasket)", len(o25), "int", ""),
        ("Lines per order - mean", "=AVERAGE(tBasket[Lines])", o25.lines.mean(), "dec2", ""),
        ("Lines per order - most common", "=MODE.SNGL(tBasket[Lines])", int(o25.lines.mode()[0]), "int", "Most shoppers buy one product line"),
        ("Basket value - mean", "=AVERAGE(tBasket[Basket_Value])", o25.val.mean(), "naira", ""),
        ("Basket value - median", "=MEDIAN(tBasket[Basket_Value])", o25.val.median(), "naira", "Mean ≈ 3x median: a few large furniture/tech orders"),
        ("Orders spanning 2+ categories", '=COUNTIFS(tBasket[Categories],">1")/ROWS(tBasket)', (o25.cats > 1).mean(), "pct", "Cross-category attach rate")])
att = []
for c1 in ["Furniture", "Office Supplies", "Technology"]:
    has1 = set(b25[b25.Category == c1].Order_ID)
    row = [c1, len(has1)]
    for c2 in ["Furniture", "Office Supplies", "Technology"]:
        has2 = set(b25[b25.Category == c2].Order_ID)
        row.append(len(has1 & has2) / len(has1) if c1 != c2 else None)
    att.append(row)
atd = pd.DataFrame(att, columns=["If basket contains", "Orders", "…also Furniture", "…also Office Supplies", "…also Technology"])
s.table(atd.astype(object).fillna(""), "tAttach", fmts={"Orders": "int", "…also Furniture": "pct", "…also Office Supplies": "pct", "…also Technology": "pct"}, widths={"If basket contains": 20})
s.band("P.3.6", "Margin erosion: quarter by quarter, and price vs mix by category")
qq = []
for q in ["Q1", "Q2", "Q3", "Q4"]:
    r = [q]
    for d in (a24, b25):
        x = d[d.Quarter == q]
        r += [x.Profit.sum() / x.Sales.sum(), (x.Discount * x.Sales).sum() / x.Sales.sum()]
    qq.append(r)
qd = pd.DataFrame({"Quarter": ["Q1", "Q2", "Q3", "Q4"]})
f_m = '=SUMIFS(tSM[Profit],tSM[Year],{y},tSM[Quarter],[@Quarter])/SUMIFS(tSM[Sales],tSM[Year],{y},tSM[Quarter],[@Quarter])'
f_d = '=SUMPRODUCT((tSM[Year]={y})*(tSM[Quarter]=[@Quarter])*tSM[Discount]*tSM[Sales])/SUMIFS(tSM[Sales],tSM[Year],{y},tSM[Quarter],[@Quarter])'
tq = s.table(qd, "tMarginQ", calc={"Margin_2024": (f_m.format(y=2024), "pct", [r[1] for r in qq]), "Discount_2024": (f_d.format(y=2024), "pct", [r[2] for r in qq]),
                                   "Margin_2025": (f_m.format(y=2025), "pct", [r[3] for r in qq]), "Discount_2025": (f_d.format(y=2025), "pct", [r[4] for r in qq])})
s.chart("column", [{"name": "Margin 2024", "categories": tq.ref("Quarter"), "values": tq.ref("Margin_2024"), "color": "#9DB4C0"},
                   {"name": "Margin 2025", "categories": tq.ref("Quarter"), "values": tq.ref("Margin_2025")}],
        f"Q4 2025 margin fell to {qq[3][3]:.1%} (from {qq[3][1]:.1%}) as discount depth rose to {qq[3][4]:.1%}", y_fmt="0%", rows=16, size=(640, 280))
cats = ["Furniture", "Office Supplies", "Technology"]
mix = []
for c in cats:
    s4, s5 = a24[a24.Category == c], b25[b25.Category == c]
    sh4, sh5 = s4.Sales.sum() / a24.Sales.sum(), s5.Sales.sum() / b25.Sales.sum()
    mg4, mg5 = s4.Profit.sum() / s4.Sales.sum(), s5.Profit.sum() / s5.Sales.sum()
    mix.append([c, sh4, sh5, mg4, mg5, sh5 * (mg5 - mg4), (sh5 - sh4) * mg4])
mxd = pd.DataFrame([[r[0]] for r in mix], columns=["Category"])
sh_f = "=SUMIFS(tSM[Sales],tSM[Year],{y},tSM[Category],[@Category])/SUMIFS(tSM[Sales],tSM[Year],{y})"
mg_f = "=SUMIFS(tSM[Profit],tSM[Year],{y},tSM[Category],[@Category])/SUMIFS(tSM[Sales],tSM[Year],{y},tSM[Category],[@Category])"
s.table(mxd, "tMixRate", widths={"Category": 18},
        calc={"Share_2024": (sh_f.format(y=2024), "pct", [r[1] for r in mix]), "Share_2025": (sh_f.format(y=2025), "pct", [r[2] for r in mix]),
              "Margin_2024": (mg_f.format(y=2024), "pct", [r[3] for r in mix]), "Margin_2025": (mg_f.format(y=2025), "pct", [r[4] for r in mix]),
              "Rate_effect_pts": ("=[@Share_2025]*([@Margin_2025]-[@Margin_2024])*100", "dec2", [r[5] * 100 for r in mix]),
              "Mix_effect_pts": ("=([@Share_2025]-[@Share_2024])*[@Margin_2024]*100", "dec2", [r[6] * 100 for r in mix])},
        total={"Category": "Total", "Rate_effect_pts": "sum", "Mix_effect_pts": "sum"})
s.insight("Rate + mix effects add up to the total margin change. The rate effect dominates: margins fell inside Furniture and Office Supplies, "
          "mostly in Q4 when discounts deepened. Mix actually helped slightly (less Furniture, more Office Supplies).")
pm = b25.groupby("Product_ID").agg(S=("Sales", "sum"), P=("Profit", "sum"))
s.calc([("SKUs sold in FY2025", '=COUNTIFS(tSKU[Sales_2025],">0")', int((pm.S > 0).sum()), "int", ""),
        ("…with margin below 15%", '=COUNTIFS(tSKU[Sales_2025],">0",tSKU[Margin_2025],"<0.15")', int((pm.P / pm.S < .15).sum()), "int", "Repricing review list"),
        ("…loss-making (negative profit)", '=COUNTIFS(tSKU[Sales_2025],">0",tSKU[Profit_2025],"<0")', int((pm.P < 0).sum()), "int", "")])
s.band("P.3.7", "Stock and slow movers - anti join: SKUs sold in FY2024 but not in FY2025")
ps = pm.S.sort_values()
k30 = int(len(ps) * .3)
s.calc([("SKUs in the ledger (all years)", "=ROWS(tSKU)", sm.Product_ID.nunique(), "int", ""),
        ("Sold in FY2024 but not FY2025 (left anti join)", '=COUNTIFS(tSKU[Sales_2024],">0",tSKU[Sales_2025],0)', len(set(a24.Product_ID) - set(b25.Product_ID)), "int", "Delisting / dead-stock candidates"),
        ("Not sold at all in FY2025", '=COUNTIFS(tSKU[Sales_2025],0)', sm.Product_ID.nunique() - b25.Product_ID.nunique(), "int", ""),
        ("Revenue share of the bottom 30% of FY2025 SKUs", f'=SUMIFS(tSKU[Sales_2025],tSKU[Sales_2025],">0",tSKU[Sales_2025],"<="&AGGREGATE(15,6,tSKU[Sales_2025]/(tSKU[Sales_2025]>0),{k30}))/SUM(tSKU[Sales_2025])',
         ps.iloc[:k30].sum() / ps.sum(), "pct", "Long tail: ties up shelf space for almost no revenue")])
s.band("P.3.8-10", "Dashboard specification, insight memo and verification")
s.grid([["KPI cards", "Revenue, Margin %, Basket Value, Growth %", "One number each, vs last year"],
        ["Category × month columns", "Where did growth come from?", "Clustered columns, 2025 vs 2024"],
        ["Store ranking", "Which stores missed target?", "Sorted bar with 100% line"],
        ["Margin by quarter", "When did margin fall?", "Column chart (above)"],
        ["Slicers", "Zone, Category, Segment", "Connected to every visual (Report Connections)"]], header=["Visual", "Question it answers", "Design"])
worst_q = qq[3]
s.calc([("Memo headline", f'="Margin fell "&TEXT(-(SUMIFS(tSM[Profit],tSM[Year],2025)/SUMIFS(tSM[Sales],tSM[Year],2025)-SUMIFS(tSM[Profit],tSM[Year],2024)/SUMIFS(tSM[Sales],tSM[Year],2024))*100,"0.0")&" points despite "&TEXT(SUMIFS(tSM[Sales],tSM[Year],2025)/SUMIFS(tSM[Sales],tSM[Year],2024)-1,"0%")&" sales growth - driven by deeper Q4 discounts"',
         f"Margin fell {-(m25 - m24) * 100:.1f} points despite {g25:.0%} sales growth - driven by deeper Q4 discounts", None, ""),
        ("Recommendation", "Cap Furniture discounts at 20% from Q1 2026 (owner: Commercial Director); review the loss-making SKUs by March.",
         "Cap Furniture discounts at 20% from Q1 2026 (owner: Commercial Director); review the loss-making SKUs by March.", None, "")])
s.grid([["Revenue reconciles to POS ledger", f"₦{b25.Sales.sum():,.0f}", "✔"], ["Profit reconciles to ledger", f"₦{b25.Profit.sum():,.0f}", "✔"],
        ["Row counts documented", f"{len(b25):,} FY2025 lines; {len(sm):,} total", "✔"], ["Slicers tested on every visual", "3 slicers × 6 visuals", "✔"],
        ["Second-machine render", "Opened on a colleague's laptop; charts and ₦ symbol intact", "✔"]], header=["Verification check", "Evidence", "Status"])

# ======================================================================================
# P.4 Capstone 2 Kingsway
# ======================================================================================
s = b.sheet("P.4 Capstone2 Kingsway", "P.4  Capstone 2 - Kingsway University: progression and the at-risk list",
            "Fixed cohort, documented progression definition, attrition by faculty, an actionable at-risk list and a quantified intervention", "P.4.1 - P.4.10")
s.band("P.4.1 / P.4.4", "Cohort and the progression definition - fixed BEFORE calculating")
s.grid([["Cohort", "All students on the 2024/25 first-semester registry extract (Kingsway_Student_Results.csv)"],
        ["Progressed", "Final_Score ≥ 62 AND Lecture_Attendance_Pct ≥ 65 (the Senate's 'good standing' rule)"],
        ["At risk", "Final_Score < 62 OR Lecture_Attendance_Pct < 65 - the complement of 'progressed'"],
        ["Target", "90% of the cohort in good standing"]], header=["Definition", "Rule"], merge_last=True)
prog = ((cv.Final_Score >= 62) & (cv.Lecture_Attendance_Pct >= 65))
tin = s.inputs([("Score threshold", 62, "int", "Senate rule"), ("Attendance threshold (%)", 65, "int", ""), ("Target progression", 0.90, "pct0", "")])
ST, AT, TG = tin["Score threshold"], tin["Attendance threshold (%)"], tin["Target progression"]
s.calc([("Cohort students", "=ROWS(tCV)", len(cv), "int", "Registry control figure: 5,840"),
        ("Progressed", f'=COUNTIFS(tCV[Final_Score],">="&{ST},tCV[Lecture_Attendance_Pct],">="&{AT})', int(prog.sum()), "int", ""),
        ("Progression %", f'=COUNTIFS(tCV[Final_Score],">="&{ST},tCV[Lecture_Attendance_Pct],">="&{AT})/ROWS(tCV)', prog.mean(), "pct", ""),
        ("Gap to target (points)", f'=(COUNTIFS(tCV[Final_Score],">="&{ST},tCV[Lecture_Attendance_Pct],">="&{AT})/ROWS(tCV)-{TG})*100', (prog.mean() - .9) * 100, "dec1", "Negative = below target")])
s.band("P.4.2-3", "Data inventory and cleaning (registry extract)")
s.grid([["Matric_No", "Text with prefix (CVU-…) - keep as text so leading characters survive"], ["Final_Score", "1 score of 101 capped at 100 (documented)"],
        ["Lecturer_Rating, Guardian_Education, Commute_Band", "Blanks labelled 'Not Recorded' rather than deleted"],
        ["Faculty / Level", "Mapped from department; 4 faculties, levels 100-400"]], header=["Field", "Cleaning applied"], merge_last=True)
s.band("P.4.5", "Attrition risk by faculty and level")
fac = cv.groupby("Faculty").size().reset_index(name="Students").sort_values("Students", ascending=False)
fac_r = [(cv[cv.Faculty == f_].pipe(lambda d: ~((d.Final_Score >= 62) & (d.Lecture_Attendance_Pct >= 65)))).mean() for f_ in fac.Faculty]
s.table(fac, "tFacultyRisk", fmts={"Students": "int"}, widths={"Faculty": 20},
        calc={"At_risk": (f'=[@Students]-COUNTIFS(tCV[Faculty],[@Faculty],tCV[Final_Score],">="&{ST},tCV[Lecture_Attendance_Pct],">="&{AT})', "int",
                          [int(r * n) for r, n in zip(fac_r, fac.Students)]),
              "At_risk_rate": ("=[@At_risk]/[@Students]", "pct", fac_r)},
        total={"Faculty": "Total", "Students": "sum", "At_risk": "sum"})
lev = cv.groupby("Level").size().reset_index(name="Students")
lev_r = [(cv[cv.Level == l_].pipe(lambda d: ~((d.Final_Score >= 62) & (d.Lecture_Attendance_Pct >= 65)))).mean() for l_ in lev.Level]
s.table(lev, "tLevelRisk", fmts={"Students": "int", "Level": "gen"},
        calc={"At_risk_rate": (f'=1-COUNTIFS(tCV[Level],[@Level],tCV[Final_Score],">="&{ST},tCV[Lecture_Attendance_Pct],">="&{AT})/[@Students]', "pct", lev_r)})
s.band("P.4.6", "The at-risk list - actionable: who, why, and what support")
risk = cv[~prog].copy()
risk["Trigger"] = np.where((risk.Final_Score < 62) & (risk.Lecture_Attendance_Pct < 65), "Score and attendance",
                           np.where(risk.Final_Score < 62, "Low score", "Low attendance"))
s.calc([("At-risk students", f"=ROWS(tCV)-COUNTIFS(tCV[Final_Score],\">=\"&{ST},tCV[Lecture_Attendance_Pct],\">=\"&{AT})", int((~prog).sum()), "int", ""),
        ("…triggered by attendance only", f'=COUNTIFS(tCV[Final_Score],">="&{ST},tCV[Lecture_Attendance_Pct],"<"&{AT})', int((risk.Trigger == "Low attendance").sum()), "int", "Pastoral follow-up"),
        ("…by score only", f'=COUNTIFS(tCV[Final_Score],"<"&{ST},tCV[Lecture_Attendance_Pct],">="&{AT})', int((risk.Trigger == "Low score").sum()), "int", "Academic tutoring"),
        ("…by both", f'=COUNTIFS(tCV[Final_Score],"<"&{ST},tCV[Lecture_Attendance_Pct],"<"&{AT})', int((risk.Trigger == "Score and attendance").sum()), "int", "Highest priority: adviser meeting in Week 5")])
rl = risk.sort_values(["Final_Score", "Lecture_Attendance_Pct"]).head(20)[["Matric_No", "Faculty", "Department", "Level", "Final_Score", "Lecture_Attendance_Pct", "Trigger"]]
rl["Recommended_support"] = rl.Trigger.map({"Score and attendance": "Adviser meeting + tutorial programme", "Low score": "Week-5 tutorial programme", "Low attendance": "Pastoral call + attendance contract"})
s.table(rl, "tAtRiskTop20", fmts={"Level": "gen", "Final_Score": "int", "Lecture_Attendance_Pct": "int"}, widths={"Department": 22, "Trigger": 20, "Recommended_support": 34})
s.para("Top 20 by lowest score shown; in the real build use =SORTBY(FILTER(tCV, (tCV[Final_Score]<62)+(tCV[Lecture_Attendance_Pct]<65)), …) so the list refreshes. "
       "Restricted distribution: named lists go to advisers only.", "note")
s.band("P.4.7", "Quantifying the intervention - state the assumption and a range")
slope = np.polyfit(cv.Lecture_Attendance_Pct, cv.Final_Score, 1)[0]
low_att = cv[(cv.Lecture_Attendance_Pct < 65)]
lift = s.inputs([("Attendance lift for low-attendance students (points)", 10, "int", "Tutorial programme assumption: +10 points attendance")])["Attendance lift for low-attendance students (points)"]
newp = ((cv.Final_Score + np.where(cv.Lecture_Attendance_Pct < 65, slope * 10, 0) >= 62) & (np.where(cv.Lecture_Attendance_Pct < 65, cv.Lecture_Attendance_Pct + 10, cv.Lecture_Attendance_Pct) >= 65)).mean()
s.calc([("Exam points per attendance point (slope)", "=SLOPE(tCV[Final_Score],tCV[Lecture_Attendance_Pct])", slope, "dec3", "Observed association, not proven cause"),
        ("Expected score gain for a lifted student", f"=SLOPE(tCV[Final_Score],tCV[Lecture_Attendance_Pct])*{lift}", slope * 10, "dec1", ""),
        ("Low-attendance students reachable", f'=COUNTIFS(tCV[Lecture_Attendance_Pct],"<"&{AT})', len(low_att), "int", ""),
        ("Progression % if the lift is achieved (estimate)", newp, newp, "pct", "Computed by re-applying the rule to the lifted cohort; report as a range, e.g. ± half the effect")])
s.band("P.4.8-10", "Dashboard, senate paper and verification")
s.grid([["KPI cards", "Cohort, Progression %, At-risk count, Average attendance"], ["Faculty bar (sorted)", "At-risk rate by faculty (above)"],
        ["Level trend", "At-risk rate by level 100-400"], ["At-risk list", "Restricted page, adviser view only"],
        ["Senate paper", f"Context (target 90%) → finding ({prog.mean():.1%}, {(prog.mean() - .9) * 100:+.1f} points) → evidence (attendance link) → recommendation (Week-5 tutorials, owner: Deans) → risks"],
        ["Verification", f"Cohort {len(cv):,} = registry; progressed {int(prog.sum()):,} = examinations office figure; every slicer drives every visual"]],
       header=["Item", "Content"], merge_last=True)

# ======================================================================================
# P.5 Capstone 3 Mercyland
# ======================================================================================
s = b.sheet("P.5 Capstone3 Mercyland", "P.5  Capstone 3 - Mercyland Hospital: occupancy, length of stay and readmission safety",
            "Bed-day occupancy from the census, LOS distribution, readmission by ward, escalation rule - no patient identifiers", "P.5.1 - P.5.10")
qocc = census.groupby("Quarter").Occupancy_Pct.mean()
pk_q = qocc.idxmax()
s.band("P.5.1", "Brief - the board's question")
s.calc([("Peak quarterly average occupancy", "=MAX(tOccQ[Occupancy])", float(qocc.max()), "pct", f"Quarter: {pk_q}"),
        ("Days above 95% occupancy (all census days)", '=COUNTIFS(tCensus[Occupancy_Pct],">0.95")', int((census.Occupancy_Pct > .95).sum()), "int", "Safe-staffing escalation days"),
        ("Days above 100% (patients on trolleys)", '=COUNTIFS(tCensus[Occupancy_Pct],">1")', int((census.Occupancy_Pct > 1).sum()), "int", "")])
s.band("P.5.2-3", "Data inventory and cleaning disclosures")
dl = L("dirty_lc_log")
s.grid([["Admissions", "Mercyland_Inpatient_Episodes_2019_2024.csv / mercyland_hospital.db", f"{len(lc):,} admissions, {lc.Admit_Date.min():%d-%b-%Y} to {lc.Admit_Date.max():%d-%b-%Y}"],
        ["Wards", "Mercyland_Wards.csv", "4 wards, 200 staffed beds (60 + 50 + 50 + 40)"], ["Bed census", "Mercyland_Daily_Bed_Census.csv", f"{len(census):,} days"],
        ["Doctors", "Mercyland_Consultants.csv", "28 doctors across departments"]], header=["Source", "File", "Scale"])
s.table(dl.head(10), "tLcCleanLog", widths={dl.columns[0]: 24, dl.columns[1]: 44, dl.columns[-1]: 40})
s.band("P.5.4", "Measuring occupancy properly: occupied bed-days ÷ available bed-days, by quarter")
oq = census.groupby("Quarter").agg(Days=("Census_Date", "count")).reset_index()
s.para("Part quarters at the start (2019-Q2) and end (2024-Q2) of the census are included but flagged by their day count.", "note")
toq = s.table(oq, "tOccQ", fmts={"Days": "int"},
              calc={"Occupied_bed_days": ("=SUMIFS(tCensus[Inpatients],tCensus[Quarter],[@Quarter])", "int", [int(census[census.Quarter == q].Inpatients.sum()) for q in oq.Quarter]),
                    "Available_bed_days": ("=SUMIFS(tCensus[Open_Beds],tCensus[Quarter],[@Quarter])", "int", [int(census[census.Quarter == q].Open_Beds.sum()) for q in oq.Quarter]),
                    "Occupancy": ("=[@Occupied_bed_days]/[@Available_bed_days]", "pct", [census[census.Quarter == q].Inpatients.sum() / census[census.Quarter == q].Open_Beds.sum() for q in oq.Quarter])})
s.chart("line", [{"name": "Occupancy", "categories": toq.ref("Quarter"), "values": toq.ref("Occupancy")}], f"Quarterly bed occupancy - peak {qocc.max():.1%} in {pk_q}",
        y_fmt="0%", legend=False, rows=16, size=(760, 280))
s.band("P.5.5", "Length of stay: mean, median, 90th percentile by ward (2023 admissions)")
l23 = lc[lc.Admit_Date.dt.year == 2023]
wd = l23.groupby("Ward").size().reset_index(name="Admissions")
s.table(wd, "tLOSWard", fmts={"Admissions": "int"}, widths={"Ward": 16},
        calc={"Mean_LOS": ("=AVERAGEIFS(tAdm23[LOS_Days],tAdm23[Ward],[@Ward])", "dec1", [l23[l23.Ward == w].LOS_Days.mean() for w in wd.Ward]),
              "Median_LOS": ("=AGGREGATE(17,6,tAdm23[LOS_Days]/(tAdm23[Ward]=[@Ward]),2)", "dec1", [l23[l23.Ward == w].LOS_Days.median() for w in wd.Ward]),
              "P90_LOS": ("=AGGREGATE(16,6,tAdm23[LOS_Days]/(tAdm23[Ward]=[@Ward]),0.9)", "dec1", [l23[l23.Ward == w].LOS_Days.quantile(.9) for w in wd.Ward])})
s.para("Mercyland's stays are spread evenly from 1 to 30 days, so mean ≈ median here. Report both anyway: the P90 (28 days) is what discharge planning needs.", "note")
s.band("P.5.6", "Readmission safety by department (2023)")
dp = l23.groupby("Department").size().reset_index(name="Discharges")
s.table(dp, "tReadmitDept", fmts={"Discharges": "int"}, widths={"Department": 24},
        calc={"Readmitted": ('=COUNTIFS(tAdm23[Department],[@Department],tAdm23[Readmitted_30d],"Yes")', "int", [int((l23[l23.Department == d].Readmitted_30d == "Yes").sum()) for d in dp.Department]),
              "Rate": ("=[@Readmitted]/[@Discharges]", "pct", [(l23[l23.Department == d].Readmitted_30d == "Yes").mean() for d in dp.Department])})
s.warn("Readmission is a safety measure: any department-level finding goes to the clinical lead for validation before it appears in a board pack.")
s.band("P.5.7", "Escalation rule as a conditional column (Q4 2023 admissions)")
q4 = lc[(lc.Admit_Date >= "2023-10-01") & (lc.Admit_Date <= "2023-12-31")][["Episode_ID", "Admit_Date", "Ward", "Department", "Admission_Route", "Age", "LOS_Days", "Readmitted_30d"]].reset_index(drop=True)
reason = np.where(q4.Readmitted_30d == "Yes", "Readmitted within 30 days", np.where(q4.LOS_Days >= 28, "Stay of 28+ days", np.where(q4.Age >= 75, "Age 75+", "")))
s.calc([("Q4 2023 admissions", "=ROWS(tEscQ4)", len(q4), "int", ""),
        ("Escalations", '=SUMPRODUCT(--(tEscQ4[Escalation_reason]<>""))', int((reason != "").sum()), "int", "Every report applies the same rule"),
        ("…readmitted", '=COUNTIFS(tEscQ4[Escalation_reason],"Readmitted*")', int((reason == "Readmitted within 30 days").sum()), "int", "Rule order matters: first matching reason is recorded"),
        ("…long stay", '=COUNTIFS(tEscQ4[Escalation_reason],"Stay*")', int((reason == "Stay of 28+ days").sum()), "int", ""),
        ("…age 75+", '=COUNTIFS(tEscQ4[Escalation_reason],"Age*")', int((reason == "Age 75+").sum()), "int", "")])
s.code('// Power Query conditional column (same rule)\n= Table.AddColumn(Source, "Escalation_reason", each\n    if [Readmitted_30d] = "Yes" then "Readmitted within 30 days"\n    else if [LOS_Days] >= 28 then "Stay of 28+ days"\n    else if [Age] >= 75 then "Age 75+"\n    else "", type text)')
s.band("P.5.8-10", "Dashboard, board pack and verification")
s.grid([["KPI cards", "Admissions, Occupancy %, Mean LOS, Readmission %"], ["Occupancy trend", "Quarterly line with 85% target band (above)"],
        ["Ward bars", "Admissions and P90 LOS by ward"], ["Board pack", "Summary, occupancy, LOS, readmission safety note, staffing recommendation, data-quality appendix, confidentiality statement"],
        ["Verification", f"Admissions {len(lc):,} = Health Records; billing ₦{lc.Bill_Amount_NGN.sum():,.0f} = finance ledger; occupancy = bed census; no names or Patient_IDs in any output"]],
       header=["Item", "Content"], merge_last=True)
s.h2("Escalation list - Q4 2023 (Episode_ID only, no patient identifiers)")
s.table(q4, "tEscQ4", fmts={"Admit_Date": "date", "Age": "int", "LOS_Days": "int"}, widths={"Department": 22},
        calc={"Escalation_reason": ('=IF([@Readmitted_30d]="Yes","Readmitted within 30 days",IF([@LOS_Days]>=28,"Stay of 28+ days",IF([@Age]>=75,"Age 75+","")))', None, reason.tolist())})

# ======================================================================================
# P.6 QA
# ======================================================================================
s = b.sheet("P.6 Workbook QA", "P.6  Workbook QA - README, reconciliation, stress, edge cases and sign-off",
            "Live reconciliation tests over this workbook's data, an edge-case audit and a release sign-off", "P.6.1 - P.6.8")
s.band("P.6.2", "The 00_README standard (Capstone 1 example)")
s.grid([["Purpose", "Explain FY2025 margin decline; support Q1 2026 pricing decision"], ["Audience", "SuperMart board, Commercial Director"],
        ["Sources & owners", "POS ledger (Finance), Store directory (Operations) - see P.2 register"],
        ["Refresh", "Drop new CSV in 01_RAW → Data > Refresh All → check reconciliation cells = 0"],
        ["Measures", "Defined on the Measures sheet with DAX and plain-English meaning"], ["Exclusions", "None for FY totals; dirty returns cleaned per log"],
        ["Limitations", "Calendar-year view; targets not adjusted for store size or opening date"], ["Version / date / contact", "v1.0, 2026-01-20, analyst@supermart.ng"]],
       header=["README section", "Content"], merge_last=True)
s.band("P.6.3", "Reconciliation testing - every test is a live formula; difference must be 0")
tests = [("SuperMart lines (all years)", "=ROWS(tSM)", len(sm), 9994), ("SuperMart FY2025 sales", "=SUMIFS(tSM[Sales],tSM[Year],2025)", b25.Sales.sum(), int(ctrl["Total Sales (₦)"])),
         ("SuperMart FY2025 profit", "=SUMIFS(tSM[Profit],tSM[Year],2025)", b25.Profit.sum(), int(ctrl["Total Profit (₦)"])),
         ("SuperMart FY2025 quantity", "=SUMIFS(tSM[Quantity],tSM[Year],2025)", int(b25.Quantity.sum()), int(ctrl["Total Quantity"])),
         ("Kingsway cohort", "=ROWS(tCV)", len(cv), 5840), ("Mercyland 2023 admissions", "=ROWS(tAdm23)", len(l23), int((lc.Admit_Date.dt.year == 2023).sum())),
         ("Bed census days", "=ROWS(tCensus)", len(census), 1797)]
td = pd.DataFrame([[t[0], t[3], float(t[2])] for t in tests], columns=["Test", "Control_total", "Workbook_value"])
tt = s.table(td, "tRecon", fmts={"Control_total": "#,##0"}, widths={"Test": 32},
             calc={"Difference": ("=[@Workbook_value]-[@Control_total]", "#,##0", [float(t[2]) - t[3] for t in tests]),
                   "Result": ('=IF(ABS([@Difference])<0.5,"PASS","FAIL")', None, ["PASS" if abs(t[2] - t[3]) < .5 else "FAIL" for t in tests])})
for i, t in enumerate(tests):
    s.ws.write_formula(tt.r0 + 1 + i, tt.ci("Workbook_value"), t[1], b.nf("#,##0", style="result"), float(t[2]))
s.band("P.6.4-7", "Stress, refresh, edge cases and performance")
s.calc([("Blank Sales cells in tSM", "=COUNTBLANK(tSM[Sales])", 0, "int", "Required field - must be 0"),
        ("Error values in tSM[Sales]", "=SUMPRODUCT(--ISERROR(tSM[Sales]))", 0, "int", "Ctrl+G > Special > Formulas > Errors finds them on a sheet"),
        ("Smallest store by FY2025 lines", '=MIN(COUNTIFS(tSM[Year],2025,tSM[Store_ID],{"ST01","ST02","ST03","ST04","ST05","ST06","ST07","ST08","ST09","ST10","ST11","ST12"}))',
         int(b25.Store_ID.value_counts().min()), "int", "Edge case: does every visual still render for the smallest group?"),
        ("Zero-sales months possible?", '=IF(MIN(COUNTIFS(tSM[Year],2025,tSM[Quarter],{"Q1","Q2","Q3","Q4"}))=0,"Yes - test DIVIDE","No")', "No", None, "Test a month with no records anyway")])
s.grid([["10x rows", "Append the sales table to itself 10 times (≈100k rows)", "Refresh < 60 s, file < 30 MB"],
        ["Extra column in the middle", "Insert a 'Store_Location' column in the RAW CSV", "Queries reference columns by name, not position"],
        ["Empty month", "Remove all December rows", "No #DIV/0!; KPI cards show blank/0 cleanly"],
        ["Refresh from clean state", "Close, reopen, Refresh All", "Row counts and reconciliation still PASS"]], header=["Stress test", "How", "Pass criterion"])
s.band("P.6.8", "Release sign-off")
s.grid([["Tested by", "Peer analyst (not the builder)"], ["What was tested", "Reconciliation table (all PASS), stress tests, slicer coverage, second machine"],
        ["Control totals used", f"₦{b25.Sales.sum():,.0f} sales; {len(sm):,} lines"], ["Version released", "v1.0 - frozen, archived read-only"], ["Date", "2026-01-20"]],
       header=["Sign-off item", "Record"], merge_last=True)

# ======================================================================================
# P.7 Defence
# ======================================================================================
s = b.sheet("P.7 Executive Defence", "P.7  Executive Defence - ten minutes and ten questions",
            "A timed plan and prepared answers built from the capstone numbers", "P.7.1 - P.7.8")
s.band("P.7.2-3", "Structuring the ten minutes")
dp_ = pd.DataFrame({"Segment": ["Question and headline", "Evidence + live demo", "Recommendation", "Limitations and next steps"],
                    "Minutes": [2, 4, 2, 2], "Content": [f"Margin fell {-(m25 - m24) * 100:.1f} pts while sales grew {g25:.0%}", "Slicer → all visuals; exception list; reconciliation cell",
                                                           "Cap Furniture discounts at 20%; owner, date, expected effect", "Calendar-year view; targets not adjusted; next: SKU-level price test"]})
s.table(dp_, "tDefencePlan", widths={"Segment": 26, "Content": 60}, total={"Segment": "Total", "Minutes": "sum"})
s.band("P.7.4-6", "Anticipating panel questions - answers ready")
qa = [["What is the source and date?", "POS ledger export, 05-Jan-2026, Finance; registered with a SHA-256 fingerprint (P.2)"],
      ["How is margin defined?", "Profit ÷ net sales after discount, excluding VAT"],
      ["What did you exclude?", "Nothing from FY totals; 15 blank and 60 duplicate rows removed from the store returns"],
      ["Does it reconcile?", f"Yes - ₦{b25.Sales.sum():,.0f} = ledger, difference ₦0 (P.6)"],
      ["Is the sample big enough?", f"{len(b25):,} FY2025 lines, {b25.Order_ID.nunique():,} orders - all transactions, not a sample"],
      ["What is the assumption behind the effect?", "If Q4 discount depth returns to the FY2024 level, Q4 margin recovers towards it - an estimate, shown as a range"],
      ["Biggest limitation?", "One year of comparison; no competitor pricing data"],
      ["What would you do differently?", "Automate the monthly folder import; add a price-test design"],
      ["How does it refresh?", "Refresh All; reconciliation cells must stay at 0"], ["Who owns it?", "Commercial Analytics; README lists the contact"]]
s.table(pd.DataFrame(qa, columns=["Question", "Prepared answer"]), "tPanelQA", widths={"Question": 36, "Prepared answer": 90})
s.band("P.7.7-8", "Closing for a decision and reflecting")
s.grid([["The ask", "Approve a 20% discount cap on Furniture from 1 Jan 2026 - Commercial Director - decision today"],
        ["Expected effect", "Recover part of the Q4 margin gap; tracked monthly on the dashboard"],
        ["Reflection (within 24 h)", "Panel probed the definition of discount depth → add it to the README"]], header=["Item", "Content"], merge_last=True)

# ======================================================================================
# P.8 Portfolio
# ======================================================================================
s = b.sheet("P.8 Portfolio & CV", "P.8  Portfolio and CV - write-ups and bullets with real numbers",
            "Six-heading write-up, anonymisation, packaging and CV bullets drawn from the three capstones", "P.8.1 - P.8.9")
s.band("P.8.2", "Project write-up - six headings (Capstone 1)")
s.grid([["Question", f"Why did SuperMart's margin fall from {m24:.1%} to {m25:.1%} while sales grew {g25:.0%}?"],
        ["Data", f"{len(sm):,} POS lines (2022-2025), 12 stores, {sm.Product_ID.nunique():,} SKUs; dirty store returns cleaned and reconciled"],
        ["Method", "Power Query cleaning, star-schema Data Model, DAX measures, rate/mix decomposition, Pareto"],
        ["Finding", f"Q4 2025 margin fell to {qq[3][3]:.1%} as discount depth rose to {qq[3][4]:.1%}; Furniture margin almost wiped out"],
        ["Recommendation", "Cap Furniture discounts at 20%; review loss-making SKUs"], ["Outcome", "Board adopted the Furniture cap for Q1 2026 (training scenario)"]],
       header=["Heading", "Content"], merge_last=True)
s.band("P.8.3-4", "Anonymising and packaging")
s.grid([["Names", "Customer names → 'Customer A, B…'; never publish Patient or respondent rows"], ["Aggregation", "Publish store/zone/faculty totals; suppress cells under 5"],
        ["Screenshots", "Crop or blur any name column"], ["Statement", "'All data is anonymised training data from the ETEH bootcamp.'"],
        ["Folder", "01_Writeup.pdf, 02_Dashboard.pdf, 03_Workbook_sample.xlsx, 00_README.txt - each named _v1.0_2026-01-20"]], header=["Step", "Practice"], merge_last=True)
s.band("P.8.5-6", "CV bullets - action verb, method, scale, result")
s.grid([[f"Built a star-schema Excel model over {len(sm):,} retail transactions for 12 stores, reconciling FY2025 sales of ₦{b25.Sales.sum() / 1e6:,.1f}m to the ledger with zero variance."],
        [f"Diagnosed a {-(m25 - m24) * 100:.1f}-point margin decline through rate/mix decomposition, identifying {int((pm.P < 0).sum())} loss-making SKUs for repricing."],
        [f"Defined and automated a progression measure for {len(cv):,} students, producing an adviser-ready at-risk list of {int((~prog).sum()):,} students by trigger."],
        [f"Measured bed occupancy from {len(census):,} daily census records and built a rule-based escalation flag applied consistently to every report."],
        ["Wrote and verified SQL joins across a 5-table marketplace database (11,000 orders), catching a join error that inflated revenue about 5.5x (₦7.19bn vs ₦1.32bn)."]],
       header=["CV bullet (every number is defensible from the workbook)"])
s.grid([["Excel", "Power Query, Data Model, DAX, PivotTables, dynamic arrays, what-if & Solver"], ["SQL", "SELECT to multi-table JOIN, aggregates, subqueries, window functions (SQLite)"],
        ["Statistics", "Descriptive statistics, correlation, regression, FORECAST.ETS"], ["AI", "Co-pilot use with verification against control totals"]],
       header=["Tool", "Describe honestly as"], merge_last=True)
s.band("P.8.7-9", "Presence, portfolio questions, keeping it current")
s.grid([["LinkedIn headline", "Data Analyst | Excel, Power Query, DAX, SQL | Retail, Health and Education analytics"],
        ["'Walk me through this dashboard'", "Question → KPI cards → the Q4 margin chart → the recommendation, in 60 seconds"],
        ["'How did you verify it?'", "Reconciliation table: 7 tests, all PASS, ₦0 variance"], ["Cadence", "One new project per quarter; retire work you would not defend"]],
       header=["Item", "Example"], merge_last=True)

# ======================================================================================
# P.9 Interviews
# ======================================================================================
s = b.sheet("P.9 Interviews & Roadmap", "P.9  Interview Preparation and the Roadmap Ahead",
            "A live Excel test with model answers, three SQL screen questions (incl. a window function) run on supermart_online.db, STAR", "P.9.1 - P.9.8")
s.band("P.9.2", "Live Excel test - 3 questions on tSM, model answers as formulas")
st25 = b25.groupby("Store_Name").agg(S=("Sales", "sum"), P=("Profit", "sum"))
st25["m"] = st25.P / st25.S
s.calc([("Q1: Total FY2025 sales for the South-South zone", '=SUMIFS(tSM[Sales],tSM[Year],2025,tSM[Zone],"South-South")', b25[b25.Zone == "South-South"].Sales.sum(), "naira", "Say aloud: 'filtering to 2025 and the zone'"),
        ("Q2: Store with the highest FY2025 margin", "Typed answer after a PivotTable: " + st25.m.idxmax(), "Typed answer after a PivotTable: " + st25.m.idxmax(), None, f"{st25.m.max():.1%} - explain why margin, not sales"),
        ("Q3: Share of FY2025 lines with a discount of 40% or more", '=COUNTIFS(tSM[Year],2025,tSM[Discount],">=0.4")/COUNTIFS(tSM[Year],2025)', (b25.Discount >= .4).mean(), "pct", "Then explain one number: these lines are mostly loss-making")])
s.band("P.9.3", "SQL screen - three questions of increasing difficulty (results from supermart_online.db)")
db = W3D / "supermart_online.db"
sqls = [("Easy - filter and aggregate: delivered orders and revenue by payment method",
         "SELECT payment_method, COUNT(*) AS orders, SUM(total_price) AS revenue\nFROM transactions\nWHERE order_status = 'Delivered'\nGROUP BY payment_method\nORDER BY revenue DESC"),
        ("Medium - join with stated cardinality (many transactions → one product): top 5 brands by net revenue",
         "SELECT p.brand_name, COUNT(*) AS orders, SUM(t.total_price) AS net_revenue\nFROM transactions t\nJOIN products p ON p.product_id = t.product_id   -- many-to-one\nWHERE t.order_status IN ('Delivered','Completed')\nGROUP BY p.brand_name\nORDER BY net_revenue DESC\nLIMIT 5"),
        ("Hard - window function: best-selling seller in each zone",
         "WITH s AS (\n  SELECT sl.store_zone, sl.store_name, SUM(t.total_price) AS revenue,\n         RANK() OVER (PARTITION BY sl.store_zone ORDER BY SUM(t.total_price) DESC) AS rnk\n  FROM transactions t JOIN sellers sl ON sl.seller_id = t.seller_id\n  WHERE t.order_status IN ('Delivered','Completed')\n  GROUP BY sl.store_zone, sl.store_name)\nSELECT store_zone, store_name, revenue FROM s WHERE rnk = 1 ORDER BY revenue DESC")]
with sqlite3.connect(db) as con:
    for i, (title, q) in enumerate(sqls):
        s.h2(title)
        s.code(q)
        res = pd.read_sql_query(q, con)
        s.table(res, f"tSqlInt{i + 1}", fmts={c: ("naira" if "revenue" in c else "int") for c in res.columns if res[c].dtype != object})
s.band("P.9.4-5", "Explaining your process and a STAR answer")
s.grid([["Situation", "SuperMart Online SQL project: the AI-written zone query returned ₦7.19bn revenue (56,725 rows)"],
        ["Task", "Verify before the number went into the deck"], ["Action", "Checked row count (should be 10,430 lines with a delivery record) and their control total ₦1.32bn; found the join on customer_id; rewrote on delivery_id"],
        ["Result", "Correct figure published; added a control-total check to every SQL deliverable"]], header=["STAR", "Answer"], merge_last=True)
s.band("P.9.6-8", "Roadmap: Power BI, SQL, Python")
s.grid([["Power BI", "Rebuild Capstone 1 in Power BI Desktop: same star schema and DAX; learn the visual layer", "2-4 weeks"],
        ["SQL", "Window functions, CTEs, indexes on supermart_online.db and mercyland_hospital.db; then PostgreSQL", "4-6 weeks"],
        ["Python", "pandas to re-create the cleaning log for SuperMart_Sales_2025_DIRTY.csv; matplotlib for the margin chart", "6-8 weeks"]],
       header=["Next tool", "First project", "Time"])

# ---------------------------------------------------------------------- data sheets
b.data_sheet("D_SuperMart_All", sm[["Row_ID", "Order_ID", "Order_Date", "Year", "Quarter", "Store_ID", "Store_Name", "Zone", "Segment", "Customer_ID", "Product_ID",
                                    "Category", "Sub_Category", "Sales", "Quantity", "Discount", "Profit"]], "tSM",
             fmts={"Order_Date": "date", "Sales": "naira", "Profit": "naira", "Discount": "pct0", "Year": "gen"}, title="SuperMart POS ledger 2022-2025 (9,994 lines)")
o25t = o25.reset_index().rename(columns={"lines": "Lines", "val": "Basket_Value", "cats": "Categories"})
b.data_sheet("D_Basket_2025", o25t, "tBasket", fmts={"Basket_Value": "naira"}, title="FY2025 orders: lines, value and categories per basket")
sku = sm.groupby("Product_ID").agg(Category=("Category", "first"), Sub_Category=("Sub_Category", "first")).reset_index()
for y in (2024, 2025):
    g_ = sm[sm.Year == y].groupby("Product_ID").agg(S=("Sales", "sum"), P=("Profit", "sum"))
    sku[f"Sales_{y}"] = sku.Product_ID.map(g_.S).fillna(0)
    sku[f"Profit_{y}"] = sku.Product_ID.map(g_.P).fillna(0)
sku["Margin_2025"] = np.where(sku.Sales_2025 > 0, sku.Profit_2025 / sku.Sales_2025.replace(0, np.nan), np.nan)
b.data_sheet("D_SKU", sku, "tSKU", fmts={"Sales_2024": "naira", "Profit_2024": "naira", "Sales_2025": "naira", "Profit_2025": "naira", "Margin_2025": "pct"},
             title="SKU summary (1,862 products): FY2024 and FY2025 sales and profit (values aggregated from tSM)")
b.data_sheet("D_Students", cv[["Matric_No", "Faculty", "Department", "Level", "Gender", "Study_Hours_Wk", "Lecture_Attendance_Pct", "Last_Sem_Score", "Final_Score", "Grade"]], "tCV",
             fmts={"Level": "gen"}, title="Kingsway cohort 2024/25 (5,840)")
b.data_sheet("D_Adm_2023", l23[["Episode_ID", "Admit_Date", "Ward", "Department", "Admission_Route", "Age", "LOS_Days", "Readmitted_30d"]].reset_index(drop=True), "tAdm23",
             fmts={"Admit_Date": "date"}, title="Mercyland 2023 admissions - minimised (no names, no Patient_No)")
b.data_sheet("D_Census", census[["Census_Date", "Quarter", "Inpatients", "Open_Beds", "Occupancy_Pct"]], "tCensus",
             fmts={"Census_Date": "date", "Occupancy_Pct": "pct"}, title="Mercyland daily bed census")

b.cover("Capstone Section Examples - Lifecycle, Source Register, Three Capstones, QA, Defence and Career",
        ["One worksheet per module (P.1 - P.9). P.3, P.4 and P.5 are worked versions of the three capstones using the real bootcamp data, "
         "so the numbers students quote in their defence can be traced to a formula.",
         "P.2 is an internal source register: every bootcamp file with its owner, rows and SHA-256 fingerprint, plus before/after evidence from the dirty extract.",
         "P.6 holds a live reconciliation table - all tests must show PASS.",
         "Health and student data appear only as minimised extracts (no names or patient IDs)."],
        [(k, DATASETS_TEXT[k]) for k in ["SuperMart Stores", "SuperMart Online", "Kingsway", "Mercyland"]])
b.close()
print("Capstone ->", OUT)
