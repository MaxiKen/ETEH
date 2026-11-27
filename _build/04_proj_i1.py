"""Instructor Capstone 1 - SuperMart Retail: growth with a shrinking margin.
Full solution workbook (live formulas, KPI cards, reconciliation, diagnosis, what-if, SKU actions, memo, defence) + instructor brief PDF."""
from __future__ import annotations

import numpy as np
import pandas as pd

from common import IPR
from proj_common import brief_pdf, kpi_cards
from wk_common import L, sm_sales
from xlhelp import Book

OUT = IPR / "Capstone1_SuperMart_Stores"
OUT.mkdir(parents=True, exist_ok=True)
XLSX = OUT / "Capstone1_SuperMart_Stores_SOLUTION.xlsx"
PDF = OUT / "Capstone1_Instructor_Brief.pdf"

# ------------------------------------------------------------------ data & reference answers (pandas)
sm = sm_sales()
sm["Year"] = sm.Order_Date.dt.year
sm["Quarter"] = "Q" + sm.Order_Date.dt.quarter.astype(str)
led = sm[sm.Year.isin([2024, 2025])].reset_index(drop=True)
band = np.select([led.Discount == 0, led.Discount <= .2, led.Discount <= .4], ["0%", "1-20%", "21-40%"], ">40%")
a24, b25 = led[led.Year == 2024], led[led.Year == 2025]
S24, S25, P24, P25 = a24.Sales.sum(), b25.Sales.sum(), a24.Profit.sum(), b25.Profit.sum()
g25, m24, m25 = S25 / S24 - 1, P24 / S24, P25 / S25
ORD25 = b25.Order_ID.nunique()
stores = L("sm_stores")
truth = L("dirty_sm_truth").copy()
dirty = L("dirty_sm")
dlog = L("dirty_sm_log")
ctrl = L("dirty_sm_ctrl").set_index("Control total (after cleaning)").Value
assert len(truth) == len(b25) and truth.Sales.sum() == S25 and truth.Profit.sum() == P25
TARGET = stores.Annual_Target_2025.sum()

# store league
lg = stores[["Store_ID", "Store_Name", "Zone"]].copy()
lg["S24"] = lg.Store_ID.map(a24.groupby("Store_ID").Sales.sum())
lg["S25"] = lg.Store_ID.map(b25.groupby("Store_ID").Sales.sum())
lg["P25"] = lg.Store_ID.map(b25.groupby("Store_ID").Profit.sum())
lg["T25"] = lg.Store_ID.map(stores.set_index("Store_ID").Annual_Target_2025)
lg = lg.sort_values("S25", ascending=False).reset_index(drop=True)
loss_stores = lg[lg.P25 < 0]

# discount bands
BANDS = ["0%", "1-20%", "21-40%", ">40%"]
bd = {}
for y, d in ((2024, a24), (2025, b25)):
    bb = band[led.Year == y]
    for k in BANDS:
        x = d[bb == k]
        bd[(y, k)] = (x.Sales.sum(), x.Profit.sum(), len(x))
deep25_loss = bd[(2025, "21-40%")][1] + bd[(2025, ">40%")][1]

# category rate / mix
CATS = ["Furniture", "Office Supplies", "Technology"]
mix = []
for c in CATS:
    s4, s5 = a24[a24.Category == c], b25[b25.Category == c]
    sh4, sh5 = s4.Sales.sum() / S24, s5.Sales.sum() / S25
    mg4, mg5 = s4.Profit.sum() / s4.Sales.sum(), s5.Profit.sum() / s5.Sales.sum()
    mix.append((c, sh4, sh5, mg4, mg5, sh5 * (mg5 - mg4) * 100, (sh5 - sh4) * mg4 * 100))
qm = {(y, q): d[d.Quarter == q].Profit.sum() / d[d.Quarter == q].Sales.sum() for y, d in ((2024, a24), (2025, b25)) for q in ["Q1", "Q2", "Q3", "Q4"]}
subl = b25.groupby("Sub_Category").agg(Sales=("Sales", "sum"), Profit=("Profit", "sum")).sort_values("Profit").reset_index()


# what-if
def whatif(cap, loss):
    x = b25
    hit = x.Discount > cap
    newsale = x.Sales / (1 - x.Discount) * (1 - cap)
    s = S25 + (hit * (newsale * (1 - loss) - x.Sales)).sum()
    p = P25 + (hit * ((x.Profit + newsale - x.Sales) * (1 - loss) - x.Profit)).sum()
    return s, p, int(hit.sum())


CAP, LOSS = 0.20, 0.25
ws_s, ws_p, ws_n = whatif(CAP, LOSS)

# SKUs (whole 2022-2025 range)
allsku = sm.groupby("Product_ID").agg(Category=("Category", "first"), Sub_Category=("Sub_Category", "first"), Product_Name=("Product_Name", "first")).reset_index()
for y in (2024, 2025):
    d = sm[sm.Year == y].groupby("Product_ID").agg(S=("Sales", "sum"), P=("Profit", "sum"))
    allsku[f"Sales_{y}"] = allsku.Product_ID.map(d.S).fillna(0).astype(int)
    allsku[f"Profit_{y}"] = allsku.Product_ID.map(d.P).fillna(0).astype(int)
n_unsold25 = int((allsku.Sales_2025 == 0).sum())
n_dropped = int(((allsku.Sales_2024 > 0) & (allsku.Sales_2025 == 0)).sum())
n_loss = int(((allsku.Sales_2025 > 0) & (allsku.Profit_2025 < 0)).sum())
loss_val = int(allsku[(allsku.Sales_2025 > 0) & (allsku.Profit_2025 < 0)].Profit_2025.sum())
top_loss = allsku[allsku.Profit_2025 < 0].sort_values("Profit_2025").head(15)

FACTS = dict(S25=S25, g=g25, m24=m24, m25=m25, dm=(m25 - m24) * 100, deep=deep25_loss, n_loss=n_loss, loss_val=loss_val, unsold=n_unsold25,
             cap_p=ws_p - P25, cap_m=ws_p / ws_s)


def build():
    b = Book(str(XLSX), "Capstone 1 - SuperMart Retail (SOLUTION)", "Instructor Capstone 1 · SuperMart Retail · FY2025 board review")
    F = b.F

    # =========================== 1 Brief & Plan
    s = b.sheet("1 Brief & Plan", "Capstone 1 - Brief, success criteria and delivery plan", "Instructor demonstration project · 7-stage lifecycle (Vol 5, P.1)", "P.1, P.3.1")
    s.band("1.1", "The board's question")
    s.scenario(f"SuperMart Retail grew FY2025 sales {g25:.1%} to ₦{S25 / 1e6:,.1f}m across 12 stores, yet gross margin fell from {m24:.1%} to {m25:.1%}. "
               "The board asks: WHY is margin falling while sales grow, WHICH stores and products are responsible, and WHAT should we do in FY2026?")
    s.grid([["Success criterion", "Evidence in this workbook"],
            ["Numbers reconcile to the ledger", "2 Clean & Reconcile - every control total shows OK"],
            ["Root cause quantified, not asserted", "4 Margin Diagnosis - discount bands, rate vs mix, sub-categories"],
            ["Recommendation is costed", "5 What-If - discount cap with volume-loss assumption and a range"],
            ["Actionable lists", "6 SKU Actions - unsold and loss-making SKUs"],
            ["One-page memo + defence", "7 Memo & Defence"]], merge_last=True)
    s.band("1.2", "Delivery plan - how to run this capstone live (about 6 contact hours)")
    s.grid([["Stage", "What you demonstrate", "Time", "Workbook sheet"],
            ["1 Frame", "Restate the question, agree KPIs and definitions (margin = profit ÷ sales)", "20 min", "1 Brief & Plan"],
            ["2 Acquire", "Load the dirty store returns + ledger with Power Query", "30 min", "2 Clean & Reconcile"],
            ["3 Clean", "Fix the 16 defect types; prove control totals", "60 min", "2 Clean & Reconcile"],
            ["4 Model", "Ledger table, store dimension, measures", "45 min", "3 Dashboard"],
            ["5 Analyse", "Discount bands, rate/mix, quarters, SKUs", "75 min", "4, 6"],
            ["6 Recommend", "Discount-cap what-if with sensitivity", "45 min", "5 What-If"],
            ["7 Communicate", "Dashboard, memo, 10-question defence", "45 min", "3, 7"]], merge_last=True)
    s.band("1.3", "Files")
    s.grid([["01_Section_Examples/00_Core_Datasets/SuperMart_Stores_Sales_2022_2025.csv", "Clean POS ledger (D_Ledger here holds FY2024-FY2025)"],
            ["01_Section_Examples/00_Core_Datasets/SuperMart_Stores_Directory.csv", "12 stores with FY2025 targets"],
            ["01_Section_Examples/Week2_Data_Files/SuperMart_Sales_2025_DIRTY.csv", f"Store returns as submitted - {len(dirty):,} rows with 16 defect types"]],
           header=["File", "Use"], merge_last=True)

    # =========================== data sheets (built first so formulas can reference them)
    dled = led[["Row_ID", "Order_ID", "Order_Date", "Year", "Quarter", "Store_ID", "Zone", "Category", "Sub_Category", "Product_ID", "Sales", "Quantity", "Discount", "Profit"]]
    # placeholder ordering: analysis sheets are created next; data sheets appended at the end by Book order of creation
    s2 = b.sheet("2 Clean & Reconcile", "Stage 2-3 - Cleaning the store returns and proving the numbers", "The cleaned extract must equal the ledger, to the naira", "P.1, P.2, P.6")
    s3 = b.sheet("3 Dashboard", "Board dashboard - FY2025 at a glance", "All tiles and tables are live formulas over tLedger and tStores", "P.3.4, P.3.8")
    s4 = b.sheet("4 Margin Diagnosis", "Why is margin falling? Discount depth, rate vs mix, quarters, sub-categories", "", "P.3.6")
    s5 = b.sheet("5 What-If", "Costing the recommendation - cap discounts at a maximum rate", "Change the yellow inputs; everything recalculates", "P.3.9")
    s6 = b.sheet("6 SKU Actions", "Product actions - unsold and loss-making SKUs", "", "P.3.7")
    s7 = b.sheet("7 Memo & Defence", "Board memo, defence questions and the marking guide", "", "P.3.10, P.7")
    tL = b.data_sheet("D_Ledger", dled, "tLedger",
                      fmts={"Order_Date": "date", "Year": "gen", "Sales": "naira", "Profit": "naira", "Discount": "pct0", "Row_ID": "gen"},
                      title=f"POS ledger FY2024-FY2025 ({len(dled):,} lines)", widths={"Order_ID": 16, "Sub_Category": 14, "Product_ID": 16})
    tS = b.data_sheet("D_Stores", stores, "tStores", fmts={"Annual_Target_2025": "naira", "Opened": "date"}, title="Store directory", widths={"Store_Name": 26, "Store_Manager": 18})
    tC = b.data_sheet("D_Clean2025", truth.drop(columns=["Customer_Name"]), "tClean25",
                      fmts={"Order_Date": "date", "Ship_Date": "date", "Sales": "naira", "Profit": "naira", "Discount": "pct0", "Row_ID": "gen"},
                      title=f"Store returns AFTER cleaning ({len(truth):,} rows; customer names removed)", widths={"Order_ID": 16, "Product_Name": 40})
    tK = b.data_sheet("D_SKU", allsku, "tSKU", fmts={c: "naira" for c in ["Sales_2024", "Profit_2024", "Sales_2025", "Profit_2025"]},
                      title="SKU summary - every product sold 2022-2025, with FY2024/FY2025 sales and profit (from the ledger)", widths={"Product_Name": 50, "Product_ID": 16})

    # =========================== 2 Clean & Reconcile
    s = s2
    s.band("2.1", "Defect log - what was wrong with the store returns, and the fix")
    s.table(dlog, "tDefects", fmts={"Rows affected": "int"}, widths={dlog.columns[1]: 40, dlog.columns[-1]: 40})
    s.para(f"Raw file {len(dirty):,} rows → remove 60 exact duplicates and 15 spacer rows → {len(truth):,} rows. Order matters: standardise text BEFORE removing duplicates, "
           "and correct ×100 keying outliers from receipts before any totals are taken.", "note")
    s.band("2.2", "Control totals - the cleaned extract against the head-office control sheet")
    rows = [("Clean row count", "=ROWS(tClean25)", len(truth), "int"), ("Total Sales (₦)", "=SUM(tClean25[Sales])", truth.Sales.sum(), "naira"),
            ("Total Profit (₦)", "=SUM(tClean25[Profit])", truth.Profit.sum(), "naira"), ("Total Quantity", "=SUM(tClean25[Quantity])", truth.Quantity.sum(), "int"),
            ("Distinct orders", "=COUNTA(UNIQUE(tClean25[Order_ID]))", truth.Order_ID.nunique(), "int"),
            ("Distinct customers", "=COUNTA(UNIQUE(tClean25[Customer_ID]))", truth.Customer_ID.nunique(), "int")]
    s.grid([], header=["Control", "Expected (control sheet)", "Computed (live)", "Difference", "Status"])
    for lab, fx, val, k in rows:
        r = s.r
        s.ws.write(r, 1, lab, F["lab"])
        s.ws.write_number(r, 2, float(ctrl[lab]), b.nf(k))
        s.ws.write_formula(r, 3, fx, b.nf(k, style="result"), float(val))
        s.ws.write_formula(r, 4, f"=D{r + 1}-C{r + 1}", b.nf(k), float(val - ctrl[lab]))
        s.ws.write_formula(r, 5, f'=IF(D{r + 1}=C{r + 1},"OK","CHECK")', F["cell"], "OK" if val == ctrl[lab] else "CHECK")
        s.r += 1
    s.gap()
    s.band("2.3", "Second proof - the cleaned extract equals the FY2025 ledger")
    s.calc([("Ledger FY2025 sales", "=SUMIFS(tLedger[Sales],tLedger[Year],2025)", S25, "naira", "Independent source: POS ledger"),
            ("Ledger FY2025 lines", "=COUNTIFS(tLedger[Year],2025)", len(b25), "int", ""),
            ("Sales difference vs cleaned extract", "=SUMIFS(tLedger[Sales],tLedger[Year],2025)-SUM(tClean25[Sales])", 0, "naira", "Must be exactly 0"),
            ("Profit difference vs cleaned extract", "=SUMIFS(tLedger[Profit],tLedger[Year],2025)-SUM(tClean25[Profit])", 0, "naira", "Must be exactly 0")])
    s.insight("Two independent proofs agree to the naira: the cleaning is complete and nothing was lost or double-counted. Only now may analysis begin.")

    # =========================== 3 Dashboard
    s = s3
    s.band("3.1", "KPI cards")
    kp = kpi_cards(s, [("FY2025 SALES", "=SUMIFS(tLedger[Sales],tLedger[Year],2025)", S25, "nairam", "12 stores"),
                       ("GROWTH vs FY2024", "=SUMIFS(tLedger[Sales],tLedger[Year],2025)/SUMIFS(tLedger[Sales],tLedger[Year],2024)-1", g25, "pct", "the good news"),
                       ("GROSS MARGIN", "=SUMIFS(tLedger[Profit],tLedger[Year],2025)/SUMIFS(tLedger[Sales],tLedger[Year],2025)", m25, "pct", f"FY2024 {m24:.1%}"),
                       ("TARGET ATTAINMENT", "=SUMIFS(tLedger[Sales],tLedger[Year],2025)/SUM(tStores[Annual_Target_2025])", S25 / TARGET, "pct", "vs ₦" + f"{TARGET / 1e6:,.0f}m"),
                       ("LOSS-MAKING STORES", '=COUNTIFS(tStoreLeague[Profit_2025],"<0")', len(loss_stores), "int", "negative FY2025 profit")])
    s.band("3.2", "Store league table (sorted by FY2025 sales)")
    lgd = lg[["Store_ID", "Store_Name", "Zone"]]
    tl = s.table(lgd, "tStoreLeague", widths={"Store_Name": 26, "Zone": 14},
                 calc={"Sales_2024": ("=SUMIFS(tLedger[Sales],tLedger[Store_ID],[@Store_ID],tLedger[Year],2024)", "nairam", lg.S24.tolist()),
                       "Sales_2025": ("=SUMIFS(tLedger[Sales],tLedger[Store_ID],[@Store_ID],tLedger[Year],2025)", "nairam", lg.S25.tolist()),
                       "Growth": ("=[@Sales_2025]/[@Sales_2024]-1", "pct", (lg.S25 / lg.S24 - 1).tolist()),
                       "Profit_2025": ("=SUMIFS(tLedger[Profit],tLedger[Store_ID],[@Store_ID],tLedger[Year],2025)", "nairam", lg.P25.tolist()),
                       "Margin_2025": ("=[@Profit_2025]/[@Sales_2025]", "pct", (lg.P25 / lg.S25).tolist()),
                       "Target_Att": ("=[@Sales_2025]/SUMIFS(tStores[Annual_Target_2025],tStores[Store_ID],[@Store_ID])", "pct", (lg.S25 / lg.T25).tolist())},
                 total={"Store_ID": "Total", "Sales_2024": "sum", "Sales_2025": "sum", "Profit_2025": "sum"})
    s.warn(f"{len(loss_stores)} stores lost money in FY2025 ({', '.join(loss_stores.Store_Name)}) - several of them still beat their sales target. "
           "Sales targets alone reward discounting.")
    s.chart("bar", [{"name": "Margin FY2025", "categories": tl.ref("Store_Name"), "values": tl.ref("Margin_2025"), "color": "#13807A",
                     "points": [{"fill": {"color": "#B23A48" if v < 0 else "#13807A"}} for v in (lg.P25 / lg.S25)]}],
            f"{len(loss_stores)} of 12 stores made a loss in FY2025", y_fmt="0%", legend=False, rows=18, size=(760, 330))

    # =========================== 4 Margin Diagnosis
    s = s4
    s.band("4.1", "Margin by discount band - where profit is destroyed")
    bdf = pd.DataFrame({"Discount_Band": BANDS, "Above": [-1, 0, 0.2, 0.4], "Up_to": [0, 0.2, 0.4, 1]})

    def bf(y, what):
        return f'SUMIFS(tLedger[{what}],tLedger[Year],{y},tLedger[Discount],">"&[@Above],tLedger[Discount],"<="&[@Up_to])'
    tb = s.table(bdf, "tBands", fmts={"Above": "pct0", "Up_to": "pct0"}, widths={"Discount_Band": 14},
                 calc={"Sales_2024": ("=" + bf(2024, "Sales"), "nairam", [bd[(2024, k)][0] for k in BANDS]),
                       "Margin_2024": ("=" + bf(2024, "Profit") + "/[@Sales_2024]", "pct", [bd[(2024, k)][1] / bd[(2024, k)][0] for k in BANDS]),
                       "Sales_2025": ("=" + bf(2025, "Sales"), "nairam", [bd[(2025, k)][0] for k in BANDS]),
                       "Profit_2025": ("=" + bf(2025, "Profit"), "nairam", [bd[(2025, k)][1] for k in BANDS]),
                       "Margin_2025": ("=[@Profit_2025]/[@Sales_2025]", "pct", [bd[(2025, k)][1] / bd[(2025, k)][0] for k in BANDS]),
                       "Share_2025": ("=[@Sales_2025]/SUM([Sales_2025])", "pct", [bd[(2025, k)][0] / S25 for k in BANDS])})
    s.para("Band rule: discount greater than 'Above' and up to 'Up_to'. The 0% band uses Above = −100% so that only zero discounts fall in it.", "note")
    s.insight(f"Lines discounted above 20% lost ₦{-deep25_loss / 1e6:,.1f}m in FY2025 on only {(bd[(2025, '21-40%')][0] + bd[(2025, '>40%')][0]) / S25:.1%} of sales. "
              f"Lines discounted above 40% run at a {bd[(2025, '>40%')][1] / bd[(2025, '>40%')][0]:.0%} margin - every naira sold there loses money.")
    s.chart("column", [{"name": "Margin FY2025", "categories": tb.ref("Discount_Band"), "values": tb.ref("Margin_2025"),
                        "points": [{"fill": {"color": c}} for c in ["#13807A", "#13807A", "#B23A48", "#B23A48"]]}],
            "Margin collapses once discount passes 20%", y_fmt="0%", legend=False, rows=15, size=(560, 260))
    s.band("4.2", "Rate vs mix: did margins fall INSIDE categories, or did the mix shift?")
    mxd = pd.DataFrame({"Category": CATS})
    sh_f = "=SUMIFS(tLedger[Sales],tLedger[Year],{y},tLedger[Category],[@Category])/SUMIFS(tLedger[Sales],tLedger[Year],{y})"
    mg_f = "=SUMIFS(tLedger[Profit],tLedger[Year],{y},tLedger[Category],[@Category])/SUMIFS(tLedger[Sales],tLedger[Year],{y},tLedger[Category],[@Category])"
    s.table(mxd, "tRateMix", widths={"Category": 18},
            calc={"Share_2024": (sh_f.format(y=2024), "pct", [r[1] for r in mix]), "Share_2025": (sh_f.format(y=2025), "pct", [r[2] for r in mix]),
                  "Margin_2024": (mg_f.format(y=2024), "pct", [r[3] for r in mix]), "Margin_2025": (mg_f.format(y=2025), "pct", [r[4] for r in mix]),
                  "Rate_pts": ("=[@Share_2025]*([@Margin_2025]-[@Margin_2024])*100", "dec2", [r[5] for r in mix]),
                  "Mix_pts": ("=([@Share_2025]-[@Share_2024])*[@Margin_2024]*100", "dec2", [r[6] for r in mix])},
            total={"Category": "Total", "Rate_pts": "sum", "Mix_pts": "sum"})
    s.para(f"Rate + mix = {sum(r[5] + r[6] for r in mix):+.2f} points = the total margin change ({(m25 - m24) * 100:+.2f}). "
           f"Furniture margin fell from {mix[0][3]:.1%} to {mix[0][4]:.1%}: the rate effect inside Furniture is the single biggest driver.", "insight")
    s.band("4.3", "Quarter by quarter")
    qd = pd.DataFrame({"Quarter": ["Q1", "Q2", "Q3", "Q4"]})
    qf = "=SUMIFS(tLedger[Profit],tLedger[Year],{y},tLedger[Quarter],[@Quarter])/SUMIFS(tLedger[Sales],tLedger[Year],{y},tLedger[Quarter],[@Quarter])"
    tq = s.table(qd, "tQuarterMargin", calc={"Margin_2024": (qf.format(y=2024), "pct", [qm[(2024, q)] for q in qd.Quarter]),
                                             "Margin_2025": (qf.format(y=2025), "pct", [qm[(2025, q)] for q in qd.Quarter])})
    s.chart("line", [{"name": "FY2024", "categories": tq.ref("Quarter"), "values": tq.ref("Margin_2024"), "color": "#9DB4C0"},
                     {"name": "FY2025", "categories": tq.ref("Quarter"), "values": tq.ref("Margin_2025")}],
            f"Q4 2025 margin {qm[(2025, 'Q4')]:.1%} - the lowest quarter in two years", y_fmt="0%", rows=15, size=(560, 260))
    s.band("4.4", "Sub-categories ranked by FY2025 profit (worst first)")
    sd = subl[["Sub_Category"]].copy()
    s.table(sd, "tSubCat", widths={"Sub_Category": 16},
            calc={"Sales_2025": ("=SUMIFS(tLedger[Sales],tLedger[Year],2025,tLedger[Sub_Category],[@Sub_Category])", "naira", subl.Sales.tolist()),
                  "Profit_2025": ("=SUMIFS(tLedger[Profit],tLedger[Year],2025,tLedger[Sub_Category],[@Sub_Category])", "naira", subl.Profit.tolist()),
                  "Margin_2025": ("=[@Profit_2025]/[@Sales_2025]", "pct", (subl.Profit / subl.Sales).tolist())})

    # =========================== 5 What-If
    s = s5
    s.band("5.1", "Assumptions (yellow) - a discount cap re-prices every line above the cap")
    inp = s.inputs([("Maximum discount allowed", CAP, "pct0", "Lines discounted above this are re-priced at the cap"),
                    ("Volume lost on re-priced lines", LOSS, "pct0", "Some customers will not buy at the higher price - the key uncertainty")])
    C, LS = inp["Maximum discount allowed"], inp["Volume lost on re-priced lines"]
    yr = "(tLedger[Year]=2025)"
    hit = f"(tLedger[Discount]>{C})"
    newp = f"tLedger[Sales]/(1-tLedger[Discount])*(1-{C})"
    s.calc([("FY2025 lines affected", f'=COUNTIFS(tLedger[Year],2025,tLedger[Discount],">"&{C})', ws_n, "int", ""),
            ("Sales after the cap", f"=SUMIFS(tLedger[Sales],tLedger[Year],2025)+SUMPRODUCT({yr}*{hit}*({newp}*(1-{LS})-tLedger[Sales]))", ws_s, "naira", ""),
            ("Profit after the cap", f"=SUMIFS(tLedger[Profit],tLedger[Year],2025)+SUMPRODUCT({yr}*{hit}*((tLedger[Profit]+{newp}-tLedger[Sales])*(1-{LS})-tLedger[Profit]))",
             ws_p, "naira", "Cost of goods is unchanged; only price and volume move"),
            ("Profit gain vs actual FY2025", f"=E{s.r + 4}-SUMIFS(tLedger[Profit],tLedger[Year],2025)", ws_p - P25, "naira", ""),
            ("Margin after the cap", f"=E{s.r + 4}/E{s.r + 3}", ws_p / ws_s, "pct", f"Actual FY2025: {m25:.1%}")])
    s.band("5.2", "Sensitivity - profit gain (₦m) for different caps and volume losses")
    caps, losses = [0.10, 0.20, 0.30], [0.0, 0.25, 0.50]
    r0 = s.r
    s.ws.write(r0, 1, "Cap ↓ / Volume lost →", F["hdr"])
    for j, lv in enumerate(losses):
        s.ws.write_number(r0, 2 + j, lv, b.nf("pct0", style="input"))
    for i, cv in enumerate(caps):
        rr = r0 + 1 + i
        s.ws.write_number(rr, 1, cv, b.nf("pct0", style="input"))
        for j, lv in enumerate(losses):
            cc, ll = f"$B{rr + 1}", f"{chr(67 + j)}${r0 + 1}"
            fx = (f"=SUMPRODUCT({yr}*(tLedger[Discount]>{cc})*((tLedger[Profit]+tLedger[Sales]/(1-tLedger[Discount])*(1-{cc})-tLedger[Sales])*(1-{ll})"
                  f"-tLedger[Profit]))/1000000")
            s.ws.write_formula(rr, 2 + j, fx, b.nf("dec1", style="result"), round((whatif(cv, lv)[1] - P25) / 1e6, 6))
    s.r = r0 + 5
    lo, hi = whatif(0.20, 0.50)[1] - P25, whatif(0.20, 0.0)[1] - P25
    s.insight(f"A 20% cap adds ₦{lo / 1e6:,.1f}m-₦{hi / 1e6:,.1f}m profit even if half the re-priced volume is lost. Present it as a range, not a single number.")
    s.warn("This is a static model: it ignores competitor reaction and basket effects (a discounted table may bring in accessory sales). Pilot in 3 stores first.")

    # =========================== 6 SKU Actions
    s = s6
    s.band("6.1", "Counts that drive the product actions")
    s.calc([("SKUs in the range (sold any year 2022-2025)", "=ROWS(tSKU)", len(allsku), "int", ""),
            ("SKUs with no FY2025 sale", "=COUNTIFS(tSKU[Sales_2025],0)", n_unsold25, "int", "Anti join: range minus FY2025 sales"),
            ("…of which sold in FY2024 but not FY2025", '=COUNTIFS(tSKU[Sales_2024],">0",tSKU[Sales_2025],0)', n_dropped, "int", "Recently dead stock - clearance candidates"),
            ("Loss-making SKUs in FY2025", '=COUNTIFS(tSKU[Sales_2025],">0",tSKU[Profit_2025],"<0")', n_loss, "int", "Repricing review list"),
            ("Profit lost on those SKUs", '=SUMIFS(tSKU[Profit_2025],tSKU[Profit_2025],"<0")', loss_val, "naira", "")])
    s.band("6.2", "15 worst SKUs by FY2025 profit")
    s.table(top_loss[["Product_ID", "Sub_Category", "Product_Name", "Sales_2025", "Profit_2025"]].reset_index(drop=True), "tWorstSKU",
            fmts={"Sales_2025": "naira", "Profit_2025": "naira"}, widths={"Product_Name": 52, "Product_ID": 17})

    # =========================== 7 Memo & Defence
    s = s7
    s.band("7.1", "Model board memo (one page)")
    memo = [f"TO: Board of SuperMart Retail   FROM: Analytics   RE: FY2025 margin review",
            f"FINDING: Sales grew {g25:.1%} to ₦{S25 / 1e6:,.1f}m but margin fell {abs(m25 - m24) * 100:.1f} points to {m25:.1%}.",
            f"EVIDENCE: Lines discounted above 20% lost ₦{-deep25_loss / 1e6:,.1f}m. The fall happened inside categories (rate effect), led by Furniture "
            f"({mix[0][3]:.1%} → {mix[0][4]:.1%}); Q4 was the weakest quarter ({qm[(2025, 'Q4')]:.1%}). {len(loss_stores)} stores made a loss while several beat their sales target.",
            f"RECOMMENDATION: Cap discounts at 20% from Q1 FY2026 (owner: Commercial Director), piloted in 3 stores for 8 weeks. Expected profit gain "
            f"₦{lo / 1e6:,.0f}m-₦{hi / 1e6:,.0f}m a year. Add a margin target alongside sales targets. Review the {n_loss} loss-making SKUs and clear the {n_dropped} dropped SKUs.",
            "RISKS: volume loss larger than assumed; competitor response. Mitigation: pilot, weekly margin tracking, stop-loss if volume falls more than 50%."]
    for m in memo:
        s.para(m, "text")
    s.band("7.2", "Ten defence questions with model answers")
    qa = [("Is the margin fall real or a data artefact?", "Real: the cleaned extract reconciles to the ledger to the naira (sheet 2)."),
          ("Is it price or mix?", f"Price (rate). Rate effect {sum(r[5] for r in mix):+.2f} pts vs mix {sum(r[6] for r in mix):+.2f} pts."),
          ("Which stores?", ", ".join(loss_stores.Store_Name) + " made losses."),
          ("Why not just stop discounting?", "Zero-discount lines are only part of demand; a cap keeps promotions but removes the loss-making depth."),
          ("How confident is the ₦ gain?", "It is a range under explicit volume-loss assumptions; the pilot will measure the real elasticity."),
          ("Did Q4 seasonality cause it?", f"Q4 2024 margin was {qm[(2024, 'Q4')]:.1%} vs {qm[(2025, 'Q4')]:.1%} in 2025 - same season, deeper discount."),
          ("What about the unsold SKUs?", f"{n_unsold25} SKUs had no FY2025 sale; {n_dropped} sold in FY2024 - clearance and delisting candidates."),
          ("Can stores game the cap?", "Monitor with the Discount_Band measure weekly per store; exceptions need approval."),
          ("Why medians nowhere?", "Margin is a ratio of totals (profit ÷ sales), never an average of line margins."),
          ("What would change your recommendation?", "If the pilot shows volume loss above about 50% on re-priced lines, revisit the cap level.")]
    s.grid([[q, a] for q, a in qa], header=["Question", "Model answer"], merge_last=True)
    s.band("7.3", "Marking guide (if run as an assessed capstone)")
    s.grid([["Data & reconciliation", "20", "All control totals OK; defects logged with counts"],
            ["Analysis", "30", "Discount bands, rate/mix, quarters and SKUs correct to the naira"],
            ["Recommendation", "20", "Costed with explicit assumptions and a range"],
            ["Dashboard", "15", "KPI cards, sorted visuals, finding-titles, no clutter"],
            ["Memo & defence", "15", "One page, numbers quoted, risks; answers 8 of 10 questions"]],
           header=["Criterion", "Marks", "Full marks when…"], merge_last=True)

    b.cover("Capstone 1 - SuperMart Retail: growth with a shrinking margin (SOLUTION)",
            ["Instructor solution workbook. Every number is a live formula over the ledger (D_Ledger), the store directory and the cleaned store returns.",
             "Run order: 1 Brief → 2 Clean & Reconcile → 3 Dashboard → 4 Margin Diagnosis → 5 What-If → 6 SKU Actions → 7 Memo & Defence."],
            [("SuperMart Stores", "SuperMart_Stores_Sales_2022_2025.csv + SuperMart_Stores_Directory.csv (built from 'Sample - Superstore.csv')"),
             ("Store returns (dirty)", "SuperMart_Sales_2025_DIRTY.csv - FY2025 returns with 16 injected defect types")])
    b.close()
    return XLSX


def brief():
    lo, hi = whatif(0.20, 0.50)[1] - P25, whatif(0.20, 0.0)[1] - P25
    brief_pdf(PDF, "Instructor Capstone 1 - SuperMart Retail: growth with a shrinking margin",
              "Full-lifecycle demonstration project · Power Query cleaning · reconciliation · DAX-style measures · what-if · board memo", "Instructor Capstone 1", [
                  ("h", "1. Purpose"),
                  ("p", "A complete worked project you run live in the capstone week (Vol 5, P.1-P.3 and P.6-P.7). It models the seven-stage lifecycle end to end on data the students "
                        "already know, so they can copy the method for their own capstone."),
                  ("h", "2. The question"),
                  ("p", f"Sales grew <b>{g25:.1%}</b> to <b>₦{S25 / 1e6:,.1f}m</b> in FY2025, but gross margin fell from <b>{m24:.1%}</b> to <b>{m25:.1%}</b>. Why, where, and what should the board do?"),
                  ("h", "3. Data"),
                  ("t", [["File", "Rows", "Role"], ["SuperMart_Stores_Sales_2022_2025.csv", f"{len(sm):,}", "Clean POS ledger"],
                         ["SuperMart_Stores_Directory.csv", "12", "Stores and FY2025 targets"],
                         ["SuperMart_Sales_2025_DIRTY.csv", f"{len(dirty):,}", "Store returns with 16 defect types → clean to 3,312"]], [80, 25, 75]),
                  ("h", "4. Answer key - the numbers students should reach"),
                  ("t", [["Measure", "Value"], ["Clean rows / sales / profit", f"{len(truth):,} / ₦{S25:,.0f} / ₦{P25:,.0f}"],
                         ["Growth / margin FY2024 → FY2025", f"{g25:.1%} / {m24:.1%} → {m25:.1%}"],
                         ["Loss on lines discounted > 20%", f"₦{-deep25_loss:,.0f}"], ["Loss-making stores", f"{len(loss_stores)}: " + ", ".join(loss_stores.Store_Name)],
                         ["Furniture margin", f"{mix[0][3]:.1%} → {mix[0][4]:.1%}"], ["Q4 2025 margin", f"{qm[(2025, 'Q4')]:.1%}"],
                         ["SKUs with no FY2025 sale / loss-making", f"{n_unsold25} / {n_loss} (₦{loss_val:,.0f})"],
                         ["20% cap, 0-50% volume loss", f"+₦{lo / 1e6:,.1f}m to +₦{hi / 1e6:,.1f}m profit"]], [80, 100]),
                  ("h", "5. Session plan"),
                  ("b", ["Frame (20 min) - success criteria; margin = Σprofit ÷ Σsales.", "Acquire + clean (90 min) - Power Query on the dirty file; prove the six control totals.",
                         "Model (45 min) - tables, relationships, measures.", "Analyse (75 min) - discount bands, rate/mix, quarters, SKUs.",
                         "Recommend (45 min) - discount-cap what-if and sensitivity grid.", "Communicate (45 min) - dashboard, memo, defence."]),
                  ("h", "6. Common student mistakes"),
                  ("b", ["Averaging line margins instead of Σprofit ÷ Σsales.", "Removing duplicates before standardising text (misses the near-duplicates).",
                         "Summing the ×100 keying outliers - sales look 30% too high.", "Presenting the what-if as one certain number."]),
                  ("note", "The solution workbook is Capstone1_SuperMart_Stores_SOLUTION.xlsx in this folder. Sheet 7 holds the memo, ten defence questions and a marking guide."),
              ])


if __name__ == "__main__":
    print(build())
    brief()
    print("facts", {k: (round(v, 4) if isinstance(v, float) else v) for k, v in FACTS.items()})
