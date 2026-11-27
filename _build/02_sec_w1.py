"""Week 1 Section Examples workbook (modules 1.1 - 1.9)."""
from __future__ import annotations

import os
import sys

import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(__file__))
from common import CORE, SEC, STAGE, W2D, W3D  # noqa: E402
from wk_common import DATASETS_TEXT, L, hotel, lc_adm, sm_sales, smo  # noqa: E402
from xlhelp import Book  # noqa: E402

OUT = SEC / "ETEH_Week1_Section_Examples.xlsx"

sm = sm_sales()
sm["Year"] = sm.Order_Date.dt.year
s25 = sm[sm.Year == 2025].copy()
stores = L("sm_stores")
lc = lc_adm()
cv = L("cv_perf")
cvs = L("cv_social")
sal = L("sal_clean")
ht = hotel()
o = smo()

b = Book(str(OUT), "ETEH Week 1 - Section Examples", "Week 1 · Foundations of Data Analytics and Excel")


def store_summary():
    g = s25.groupby("Store_ID").agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"),
                                    Orders=("Order_ID", "nunique")).reset_index()
    g = stores[["Store_ID", "Store_Name", "Zone", "Annual_Target_2025"]].merge(g, on="Store_ID")
    return g.rename(columns={"Annual_Target_2025": "Target"})


SS = store_summary()

# ======================================================================================
# 1.1 Introduction to Data Analytics
# ======================================================================================
s = b.sheet("1.1 Intro to Analytics", "1.1  Introduction to Data Analytics",
            "Data vs information, DIKW, sharp questions, and deciding with data rather than assumptions",
            "1.1.1 - 1.1.14")
s.band("1.1.1-3", "Data versus information: twelve raw till lines from Lagos - Ikeja")
s.scenario("Folake Adeyemi manages SuperMart Lagos - Ikeja. Her till system produces lines like the ones below - "
           "each line is DATA: a fact with no context. On its own, one line tells her nothing about how the store is doing. "
           "The formulas underneath turn the same twelve lines into INFORMATION she can act on.")
raw = s25[s25.Store_ID == "ST01"].sort_values("Order_Date").head(12)[
    ["Order_ID", "Order_Date", "Customer_Name", "Category", "Sub_Category", "Sales", "Quantity", "Profit"]].reset_index(drop=True)
t11 = s.table(raw, "tIkejaLines", fmts={"Sales": "naira", "Profit": "naira", "Order_Date": "date", "Quantity": "int"},
              widths={"Customer_Name": 18, "Sub_Category": 14, "Order_ID": 16})
cats = ["Furniture", "Office Supplies", "Technology"]
rows = [("Total sales on these 12 lines", "=SUM(tIkejaLines[Sales])", raw.Sales.sum(), "naira", "The headline number: what came through the tills."),
        ("Total profit", "=SUM(tIkejaLines[Profit])", raw.Profit.sum(), "naira", "Sales minus cost - the money the store actually keeps."),
        ("Profit margin", "=SUM(tIkejaLines[Profit])/SUM(tIkejaLines[Sales])", raw.Profit.sum() / raw.Sales.sum(), "pct", "Kobo of profit per ₦1 of sales."),
        ("Distinct customers served", "=COUNTA(UNIQUE(tIkejaLines[Customer_Name]))", raw.Customer_Name.nunique(), "int", "UNIQUE is covered in 2.8.62 - here it simply removes repeats.")]
for c in cats:
    rows.append((f"Sales - {c}", f'=SUMIFS(tIkejaLines[Sales],tIkejaLines[Category],"{c}")', raw[raw.Category == c].Sales.sum(), "naira",
                 "Same data, now grouped - this is information."))
s.calc(rows)
s.insight("Data = the 12 rows. Information = 'the first 12 lines of 2025 brought in "
          f"₦{raw.Sales.sum():,.0f} at a {raw.Profit.sum() / raw.Sales.sum():.1%} margin'. The data did not change - only the context did.")

s.band("1.1.4-7", "The DIKW ladder on FY2025 store results")
s.scenario("Head office has FY2025 results for all 12 stores. Climb the ladder: DATA (the table) → INFORMATION (totals, margins) → "
           "KNOWLEDGE (patterns: which zone converts sales into profit best) → WISDOM (a decision you can defend).")
t = s.table(SS[["Store_ID", "Store_Name", "Zone", "Sales", "Profit", "Orders"]], "tStores25",
            fmts={"Sales": "naira", "Profit": "naira", "Orders": "int"},
            calc={"Margin": ("=[@Profit]/[@Sales]", "pct", (SS.Profit / SS.Sales).tolist()),
                  "Avg_Order_Value": ("=[@Sales]/[@Orders]", "naira", (SS.Sales / SS.Orders).tolist())},
            total={"Store_ID": "Total", "Sales": "sum", "Profit": "sum", "Orders": "sum"},
            widths={"Store_Name": 26})
zones = ["South-West", "South-South", "North-Central", "South-East"]
zrows = []
for z in zones:
    zz = SS[SS.Zone == z]
    zrows.append((f"Margin - {z}", f'=SUMIFS(tStores25[Profit],tStores25[Zone],"{z}")/SUMIFS(tStores25[Sales],tStores25[Zone],"{z}")',
                  zz.Profit.sum() / zz.Sales.sum(), "pct", "KNOWLEDGE: compare zones on the same yardstick."))
s.calc([("DATA: store-year records", "=ROWS(tStores25)", 12, "int", "Twelve facts, one per store."),
        ("INFORMATION: chain sales", "=SUM(tStores25[Sales])", SS.Sales.sum(), "naira", "Total FY2025 revenue."),
        ("INFORMATION: chain margin", "=SUM(tStores25[Profit])/SUM(tStores25[Sales])", SS.Profit.sum() / SS.Sales.sum(), "pct", "Weighted margin (not an average of margins).")] + zrows)
best_z = max(zones, key=lambda z: SS[SS.Zone == z].Profit.sum() / SS[SS.Zone == z].Sales.sum())
worst_z = min(zones, key=lambda z: SS[SS.Zone == z].Profit.sum() / SS[SS.Zone == z].Sales.sum())
s.insight(f"WISDOM: {best_z} turns sales into profit best and {worst_z} worst. Before we copy {best_z} pricing everywhere, "
          "Module 1.2 asks WHY (diagnostic) - the discount analysis there explains most of the gap.")
s.chart("bar", [{"name": "Margin", "categories": t.ref("Store_Name"), "values": t.ref("Margin"), "labels": True, "label_fmt": "0%"}],
        "FY2025 profit margin by store", y_fmt="0%", legend=False, rows=17, size=(720, 320))

s.band("1.1.8-9", "Who does what: analyst, scientist, business analyst - on our own projects")
s.grid([["Data Analyst", "Builds the FY2025 store scorecard; explains why Owerri missed target", "Excel, Power Query, PivotTables, SQL", "SuperMart, Lifecare"],
        ["Data Scientist", "Builds a model predicting which Lifecare patients will be readmitted within 30 days", "Python/R, statistics, machine learning", "Lifecare"],
        ["Business Analyst", "Maps how Crestview registers students and specifies a new at-risk alert process", "Process maps, requirements, stakeholder workshops", "Crestview"],
        ["Data Engineer", "Loads SuperMart Online orders nightly into the SQLite/SQL database with keys", "SQL, pipelines, scheduling", "SuperMart Online"]],
       header=["Role", "Example task in this bootcamp", "Typical tools", "Dataset"], merge_last=False)
s.gap()

s.band("1.1.11-12", "From a business problem to an analytical question")
s.scenario("Stakeholders speak in problems; analysts answer questions that a dataset can settle. Each row converts one real "
           "complaint into a measurable question, the metric, and the exact fields you would use.")
s.grid([["SuperMart MD: 'Some stores are bleeding money.'", "Which stores had a profit margin below 8% in FY2025, and which discount levels drive it?", "Profit ÷ Sales by store and discount band", "Store_ID, Sales, Profit, Discount"],
        ["Crestview Registrar: 'Too many students are failing.'", "What share of students scored below 50, and how does it differ by attendance band?", "% of students with Exam_Score < 50", "Exam_Score, Attendance_Band, Faculty"],
        ["Lifecare CMD: 'The wards are always full.'", "What was daily bed occupancy in 2023 and on how many days did it exceed 95%?", "Patients in house ÷ 520 staffed beds", "Census_Date, Patients_In_House"],
        ["Eko Suites GM: 'Online guests cancel a lot.'", "What is the cancellation rate by booking channel and lead time?", "Cancelled ÷ all bookings", "Market_Segment, Lead_Time_Days, Booking_Status"]],
       header=["Business problem (as said)", "Analytical question (testable)", "Metric", "Fields needed"])
s.gap()

s.band("1.1.14", "Deciding with data, not assumptions: 'Ikeja is our best store'")
s.scenario("Everyone at head office 'knows' Lagos - Ikeja is the best store because it sells the most. Test the assumption "
           "against the target head office set in December 2024.")
tt = SS[["Store_ID", "Store_Name", "Sales", "Target"]].copy()
t14 = s.table(tt, "tTarget25", fmts={"Sales": "naira", "Target": "naira"},
              calc={"Attainment": ("=[@Sales]/[@Target]", "pct", (tt.Sales / tt.Target).tolist()),
                    "Status": ('=IF([@Attainment]>=1,"Met target","Below target")', None,
                               np.where(tt.Sales / tt.Target >= 1, "Met target", "Below target").tolist())},
              widths={"Store_Name": 26})
att = tt.Sales / tt.Target
s.calc([("Store with highest sales", "=INDEX(tTarget25[Store_Name],MATCH(MAX(tTarget25[Sales]),tTarget25[Sales],0))",
         tt.Store_Name[tt.Sales.idxmax()], None, "The assumption - biggest till."),
        ("Store with highest attainment", "=INDEX(tTarget25[Store_Name],MATCH(MAX(tTarget25[Attainment]),tTarget25[Attainment],0))",
         tt.Store_Name[att.idxmax()], None, "The evidence - best against its own target."),
        ("Stores that met target", '=COUNTIF(tTarget25[Status],"Met target")', int((att >= 1).sum()), "int", ""),
        ("Lowest attainment", "=MIN(tTarget25[Attainment])", att.min(), "pct", tt.Store_Name[att.idxmin()])])
s.insight(f"Ikeja is the biggest store, but {tt.Store_Name[att.idxmax()]} beat its target by the widest margin ({att.max():.0%}). "
          f"'Best' depends on the yardstick - agree the metric before you crown a winner.")
s.tryit("Change Lekki's Target to 75,000,000 in the table. Watch Attainment, Status and the 'highest attainment' answer update.")

# ======================================================================================
# 1.2 Types of Data Analytics
# ======================================================================================
s = b.sheet("1.2 Types of Analytics", "1.2  Types of Data Analytics",
            "Descriptive → Diagnostic → Predictive → Prescriptive on SuperMart, Crestview and Lifecare", "1.2.1 - 1.2.6")
s.band("1.2.1", "Descriptive - what happened? SuperMart sales 2022-2025")
yr = sm.groupby("Year").agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"), Orders=("Order_ID", "nunique")).reset_index()
growth = [None] + (yr.Sales.values[1:] / yr.Sales.values[:-1] - 1).tolist()
ty = s.table(yr, "tYears", fmts={"Year": "gen", "Sales": "naira", "Profit": "naira", "Orders": "int"},
             calc={"Margin": ("=[@Profit]/[@Sales]", "pct", (yr.Profit / yr.Sales).tolist()),
                   "Sales_Growth": ('=IFERROR([@Sales]/OFFSET([@Sales],-1,0)-1,"")', "pct", [g if g is not None else "" for g in growth])})
s.insight(f"Sales grew from ₦{yr.Sales.iloc[0] / 1e6:,.1f}m (2022) to ₦{yr.Sales.iloc[-1] / 1e6:,.1f}m (2025); "
          f"margin improved from {yr.Profit.iloc[0] / yr.Sales.iloc[0]:.1%} to {yr.Profit.iloc[-1] / yr.Sales.iloc[-1]:.1%}. 2023 was the only dip.")
s.chart("column", [{"name": "Sales", "categories": ty.ref("Year"), "values": ty.ref("Sales")}], "Sales by year (₦)",
        y_fmt='₦#,##0,,"m"', secondary={"type": "line", "name": "Margin", "categories": ty.ref("Year"), "values": ty.ref("Margin"),
                                       "y2_title": "Margin", "y2_fmt": "0%"}, rows=16)

s.band("1.2.2", "Diagnostic - why did it happen? Discounts and profit (all years)")
sm["Disc_Band"] = pd.cut(sm.Discount, [-0.01, 0.0, 0.2, 0.4, 1.0], labels=["0% (no discount)", "1-20%", "21-40%", "Above 40%"])
db = sm.groupby("Disc_Band", observed=True).agg(Lines=("Row_ID", "count"), Sales=("Sales", "sum"), Profit=("Profit", "sum")).reset_index()
db["Disc_Band"] = db.Disc_Band.astype(str)
tdb = s.table(db, "tDiscBand", fmts={"Lines": "int", "Sales": "naira", "Profit": "naira"},
              calc={"Margin": ("=[@Profit]/[@Sales]", "pct", (db.Profit / db.Sales).tolist())}, widths={"Disc_Band": 18})
s.insight(f"Lines discounted above 20% lose money: 21-40% earned ₦{db.Profit.iloc[2] / 1e6:,.1f}m and 'Above 40%' "
          f"₦{db.Profit.iloc[3] / 1e6:,.1f}m. Discounting - not location - is the main reason some stores have weak margins.")

s.band("1.2.3", "Predictive - what might happen? A first forecast for 2026")
fc = float(np.polyval(np.polyfit(yr.Year, yr.Sales, 1), 2026))
s.calc([("Linear-trend forecast, 2026 sales", "=FORECAST.LINEAR(2026,tYears[Sales],tYears[Year])", fc, "naira",
         "Fits a straight line through 4 years. Week 4 (4.5) adds seasonality and confidence bands."),
        ("Implied growth vs 2025", "=FORECAST.LINEAR(2026,tYears[Sales],tYears[Year])/INDEX(tYears[Sales],4)-1", fc / yr.Sales.iloc[-1] - 1, "pct",
         "Is this plausible? The last two years grew faster than the trend line.")])
s.warn("Four data points are a thin basis for prediction. Treat this as a starting estimate, never as a promise.")

s.band("1.2.4", "Prescriptive - what should we do? Cap the discount")
s.scenario("Proposed policy: no line may be discounted above a cap. The table shows profit at each discount level; "
           "the formula adds up the profit on lines ABOVE the cap - the losses the policy would stop (assuming those sales still happen at the cap).")
dl = sm.groupby("Discount").agg(Lines=("Row_ID", "count"), Sales=("Sales", "sum"), Profit=("Profit", "sum")).reset_index()
tdl = s.table(dl, "tDiscLevel", fmts={"Discount": "pct0", "Lines": "int", "Sales": "naira", "Profit": "naira"})
inp = s.inputs([("Discount cap (type a %)", 0.2, "pct0", "Try 30% or 10%")])
cap = inp["Discount cap (type a %)"]
s.calc([("Lines above the cap", f'=SUMIFS(tDiscLevel[Lines],tDiscLevel[Discount],">"&{cap})', int(dl[dl.Discount > 0.2].Lines.sum()), "int", ""),
        ("Profit on lines above the cap", f'=SUMIFS(tDiscLevel[Profit],tDiscLevel[Discount],">"&{cap})', dl[dl.Discount > 0.2].Profit.sum(), "naira",
         "Negative = money lost on deep discounts."),
        ("Share of all sales affected", f'=SUMIFS(tDiscLevel[Sales],tDiscLevel[Discount],">"&{cap})/SUM(tDiscLevel[Sales])',
         dl[dl.Discount > 0.2].Sales.sum() / dl.Sales.sum(), "pct", "How much revenue the policy touches.")])
s.insight("Prescriptive analytics recommends an action and sizes it. Here: cap discounts at 20% and review the affected lines with store managers.")

s.band("1.2.5-6", "The four types applied to the university and the hospital")
cvb = cv.groupby("Attendance_Band").agg(Students=("Student_ID", "count"), At_Risk=("At_Risk", lambda x: (x == "Yes").sum())).reset_index()
tcv = s.table(cvb, "tCvAttend", fmts={"Students": "int", "At_Risk": "int"},
              calc={"At_Risk_Rate": ("=[@At_Risk]/[@Students]", "pct", (cvb.At_Risk / cvb.Students).tolist())})
lct = lc.groupby("Admission_Type").agg(Admissions=("Admission_ID", "count"), Readmitted=("Readmit_30d", lambda x: (x == "Yes").sum()),
                                       Avg_LOS=("Length_of_Stay", "mean")).reset_index()
tlc = s.table(lct, "tLcType", fmts={"Admissions": "int", "Readmitted": "int", "Avg_LOS": "dec1"},
              calc={"Readmit_Rate": ("=[@Readmitted]/[@Admissions]", "pct", (lct.Readmitted / lct.Admissions).tolist())})
s.grid([["Descriptive", "What share of students is at risk? (see table above)", "How many admissions per type, average stay?"],
        ["Diagnostic", "At-risk rate falls sharply as attendance rises", "Emergency admissions are readmitted more often"],
        ["Predictive", "Flag likely failures from attendance + hours studied (4.4)", "Score each discharge for 30-day readmission risk"],
        ["Prescriptive", "Mentor every student below 70% attendance by week 6", "Phone follow-up within 7 days for emergency discharges"]],
       header=["Type", "Crestview University", "Lifecare General Hospital"])

# ======================================================================================
# 1.3 The Data Analytics Process
# ======================================================================================
s = b.sheet("1.3 Analytics Process", "1.3  The Data Analytics Process",
            "Objective → collect → clean → transform → analyse → interpret → present, with control totals", "1.3.1 - 1.3.12")
s.band("1.3.1", "Define the objective - SMART, on the SuperMart FY2025 review")
s.grid([["Specific", "Explain why 5 of 12 stores missed their FY2025 sales target"],
        ["Measurable", "Attainment = Sales ÷ Target per store; margin = Profit ÷ Sales"],
        ["Achievable", "Uses the POS extract (3,312 FY2025 lines) and the store directory only"],
        ["Relevant", "Feeds the January 2026 budget meeting"],
        ["Time-bound", "Draft by Friday of Week 1; final after QA in Week 2"]], header=["Element", "Our objective"], merge_last=True)
s.gap()
s.band("1.3.2-4", "Data sources used in this bootcamp")
cat = pd.read_csv(STAGE / "catalogue.csv")
s.grid(cat[["Dataset", "Organisation", "Rows", "Columns", "Grain", "Period"]].values.tolist(),
       header=["File", "Organisation", "Rows", "Columns", "Grain (one row =)", "Period"], fmts=[None, None, "int", "int", None, None])
s.para("Every file lives in 01_Section_Examples/00_Core_Datasets (CSV), Week2_Data_Files (messy formats) and Week3_Data_Files (SQLite databases + SQL scripts). "
       "Full field definitions: 00_Data_Dictionary_and_Catalogue.xlsx.", "note")
s.gap()
s.band("1.3.5", "Clean and prepare - eight messy POS lines, fixed with formulas")
s.scenario("Below are real defects from SuperMart_Sales_2025_DIRTY.csv (Week2_Data_Files). Each yellow-headed column on the right "
           "repairs one problem with a formula, so the fix is visible and repeatable. Week 2 automates the same steps in Power Query.")
dirty = L("dirty_sm")
pick = []
pick += dirty[dirty.Store_Name.str.isupper().fillna(False)].head(2).index.tolist()
pick += dirty[dirty.Sales.astype(str).str.contains("₦")].head(2).index.tolist()
pick += dirty[dirty.Discount.astype(str).str.contains("%")].head(2).index.tolist()
pick += dirty[dirty.Zone.astype(str).str.startswith(" ") | dirty.Zone.astype(str).str.endswith(" ")].head(2).index.tolist()
dd = dirty.loc[pick, ["Order_ID", "Store_Name", "Zone", "Sales", "Discount"]].astype(str).reset_index(drop=True)
dd.columns = ["Order_ID", "Store_Name_raw", "Zone_raw", "Sales_raw", "Discount_raw"]


def _clean_sales(v):
    return float(str(v).replace("₦", "").replace(",", ""))


def _clean_disc(v):
    v = str(v)
    return float(v.replace("%", "")) / 100 if "%" in v else float(v)


tcl = s.table(dd, "tDirty8",
              calc={"Store_Name": ("=PROPER(TRIM([@Store_Name_raw]))", None, [x.strip().title() for x in dd.Store_Name_raw]),
                    "Zone": ("=TRIM([@Zone_raw])", None, [x.strip() for x in dd.Zone_raw]),
                    "Sales": ('=VALUE(SUBSTITUTE(SUBSTITUTE([@Sales_raw],"₦",""),",",""))', "naira", [_clean_sales(x) for x in dd.Sales_raw]),
                    "Discount": ('=IF(ISNUMBER(SEARCH("%",[@Discount_raw])),VALUE(SUBSTITUTE([@Discount_raw],"%",""))/100,VALUE([@Discount_raw]))', "pct0",
                                 [_clean_disc(x) for x in dd.Discount_raw])},
              widths={"Store_Name_raw": 24, "Store_Name": 24, "Zone_raw": 15, "Order_ID": 16, "Sales_raw": 14})
s.warn("PROPER also changes 'GRA' to 'Gra' in 'Port Harcourt - GRA'. Always eyeball the output - a mapping table (Week 2) is safer for names.")

s.band("1.3.7-8", "Analyse and interpret - where does FY2025 profit come from?")
cg = s25.groupby("Category").agg(Sales=("Sales", "sum"), Profit=("Profit", "sum")).reset_index()
tcg = s.table(cg, "tCat25", fmts={"Sales": "naira", "Profit": "naira"},
              calc={"Margin": ("=[@Profit]/[@Sales]", "pct", (cg.Profit / cg.Sales).tolist()),
                    "Share_of_Profit": ("=[@Profit]/SUM([Profit])", "pct", (cg.Profit / cg.Profit.sum()).tolist())},
              total={"Category": "Total", "Sales": "sum", "Profit": "sum"})
top = cg.sort_values("Profit").iloc[-1]
low = cg.sort_values("Margin" if "Margin" in cg else "Profit").iloc[0]
s.insight(f"ANALYSIS: {top.Category} delivers {top.Profit / cg.Profit.sum():.0%} of FY2025 profit. INTERPRETATION: Furniture sells "
          f"₦{cg.set_index('Category').Sales['Furniture'] / 1e6:,.0f}m but keeps only {cg.set_index('Category').Profit['Furniture'] / cg.set_index('Category').Sales['Furniture']:.1%} - "
          "the next question is whether Furniture discounts are the cause (they are: see 1.2.2).")

s.band("1.3.11", "Accuracy and data quality - control totals")
s.scenario("A control total is a figure you know independently (from the source system) that your cleaned data must reproduce. "
           "Below, the POS system's own totals for FY2025 are checked against the clean store table on sheet 1.1.")
ctrl = L("dirty_sm_ctrl").set_index("Control total (after cleaning)").Value
s.calc([("POS system total sales FY2025", int(ctrl["Total Sales (₦)"]), int(ctrl["Total Sales (₦)"]), "naira", "Typed from the POS end-of-year report."),
        ("Our cleaned total (sheet 1.1)", "=SUM(tStores25[Sales])", SS.Sales.sum(), "naira", "Re-computed from the data."),
        ("Difference", f"=SUM(tStores25[Sales])-{int(ctrl['Total Sales (₦)'])}", SS.Sales.sum() - ctrl["Total Sales (₦)"], "naira", "Must be exactly 0."),
        ("Check", f'=IF(SUM(tStores25[Sales])={int(ctrl["Total Sales (₦)"])},"PASS","FAIL - investigate")',
         "PASS" if SS.Sales.sum() == ctrl["Total Sales (₦)"] else "FAIL", None, "Put a check like this at the top of every report.")])

s.band("1.3.12", "Common mistakes in the process - seen in our own data")
s.grid([["Summing a text column", "Sales typed as '₦1,234,560' are ignored by SUM - total silently too low", "ISNUMBER check; VALUE/SUBSTITUTE"],
        ["Counting duplicates twice", "60 duplicate lines in the FY2025 POS extract inflate sales", "Remove Duplicates; compare COUNTA vs distinct Row_ID"],
        ["Averaging averages", "Average of 12 store margins ≠ chain margin (big stores weigh more)", "Divide total profit by total sales"],
        ["Treating blank as zero", "30 blank Quantity cells are unknown, not zero", "COUNTBLANK; flag and recover from source"],
        ["Answering a different question", "Reporting sales when the MD asked about profit", "Write the question at the top of the sheet"]],
       header=["Mistake", "Where it shows up in our data", "Prevention"])

# ======================================================================================
# 1.4 Applications across industries
# ======================================================================================
s = b.sheet("1.4 Industry Applications", "1.4  Applications of Data Analytics Across Industries",
            "One live KPI per industry, each computed from a bootcamp dataset", "1.4.1 - 1.4.10")
s.band("1.4.1", "Retail & e-commerce - average order value by zone (SuperMart FY2025)")
zz = s25.groupby("Zone").agg(Sales=("Sales", "sum"), Orders=("Order_ID", "nunique")).reset_index()
s.table(zz, "tZoneAOV", fmts={"Sales": "naira", "Orders": "int"},
        calc={"Avg_Order_Value": ("=[@Sales]/[@Orders]", "naira", (zz.Sales / zz.Orders).tolist())})
tr = o["transactions"]
pm = tr.groupby("payment_method").agg(Orders=("order_id", "count"), Revenue=("total_price", "sum")).reset_index().sort_values("Revenue", ascending=False)
s.band("1.4.4", "Banking & payments - how SuperMart Online customers pay")
s.table(pm, "tPayMix", fmts={"Orders": "int", "Revenue": "naira"},
        calc={"Share_of_Revenue": ("=[@Revenue]/SUM([Revenue])", "pct", (pm.Revenue / pm.Revenue.sum()).tolist())}, widths={"payment_method": 18})
s.band("1.4.2", "Education - pass rate by faculty (Crestview, pass = score ≥ 50)")
fac = cv.groupby("Faculty").agg(Students=("Student_ID", "count"), Passed=("Exam_Score", lambda x: (x >= 50).sum()),
                                Avg_Score=("Exam_Score", "mean")).reset_index()
s.table(fac, "tFacPass", fmts={"Students": "int", "Passed": "int", "Avg_Score": "dec1"},
        calc={"Pass_Rate": ("=[@Passed]/[@Students]", "pct", (fac.Passed / fac.Students).tolist())}, widths={"Faculty": 26})
s.band("1.4.3", "Healthcare - average length of stay and bill by department (Lifecare)")
dep = lc.groupby("Department").agg(Admissions=("Admission_ID", "count"), Bed_Days=("Length_of_Stay", "sum"),
                                   Billing=("Billing_Amount", "sum")).reset_index()
s.table(dep, "tDeptALOS", fmts={"Admissions": "int", "Bed_Days": "int", "Billing": "nairam"},
        calc={"ALOS_days": ("=[@Bed_Days]/[@Admissions]", "dec1", (dep.Bed_Days / dep.Admissions).tolist()),
              "Avg_Bill": ("=[@Billing]/[@Admissions]", "naira", (dep.Billing / dep.Admissions).tolist())}, widths={"Department": 22})
s.band("1.4.6", "Marketing & customer management - spend by loyalty tier (SuperMart Online)")
cu = o["customers"]
tier = cu.groupby("membership_tier").agg(Customers=("customer_id", "count"), Total_Spent=("total_spent", "sum")).reset_index()
s.table(tier, "tTier", fmts={"Customers": "int", "Total_Spent": "naira"},
        calc={"Spend_per_Customer": ("=[@Total_Spent]/[@Customers]", "naira", (tier.Total_Spent / tier.Customers).tolist())},
        widths={"membership_tier": 16})
s.band("1.4.7", "Human resources - pay by seniority (ETEH Talent Salary Survey)")
sen = sal.groupby("Seniority").agg(Respondents=("Respondent_ID", "count"), Avg_Salary=("Annual_Salary_NGN", "mean"),
                                   Avg_Experience=("Years_of_Experience", "mean")).reset_index()
order = ["Junior", "Mid", "Senior", "Executive"]
sen["o"] = sen.Seniority.map({k: i for i, k in enumerate(order)}).fillna(99)
sen = sen.sort_values("o").drop(columns="o")
s.table(sen, "tSeniority", fmts={"Respondents": "int", "Avg_Salary": "naira", "Avg_Experience": "dec1"},
        calc={"Salary_per_Year_Exp": ('=IFERROR([@Avg_Salary]/[@Avg_Experience],"")', "naira",
                                      [a / e if e else "" for a, e in zip(sen.Avg_Salary, sen.Avg_Experience)])}, widths={"Seniority": 20})
s.band("1.4.8", "Logistics - delivery outcome by courier (SuperMart Online)")
dv = o["deliveries"].copy()
dv["Days"] = (dv.delivered_date - dv.shipped_date).dt.days
cr = dv.groupby("courier_name").agg(Shipments=("delivery_id", "count"),
                                    Delivered=("delivery_status", lambda x: (x == "Delivered").sum()),
                                    Avg_Days=("Days", "mean")).reset_index()
s.table(cr, "tCourier", fmts={"Shipments": "int", "Delivered": "int", "Avg_Days": "dec1"},
        calc={"Delivered_Rate": ("=[@Delivered]/[@Shipments]", "pct", (cr.Delivered / cr.Shipments).tolist())}, widths={"courier_name": 18})
s.band("1.4.5 / 1.4.10", "Hospitality & public-facing services - cancellations by channel (Eko Suites)")
hs = ht.groupby("Market_Segment").agg(Bookings=("Booking_ID", "count"),
                                      Cancelled=("Booking_Status", lambda x: (x == "Cancelled").sum())).reset_index()
s.table(hs, "tHotelSeg", fmts={"Bookings": "int", "Cancelled": "int"},
        calc={"Cancel_Rate": ("=[@Cancelled]/[@Bookings]", "pct", (hs.Cancelled / hs.Bookings).tolist())}, widths={"Market_Segment": 16})
s.insight("Same pattern every time: COUNT or SUM a group, divide by its total, compare groups. The industry changes; the method does not. "
          "Geology & mining (1.4.9) uses the identical pattern on assay or production data - e.g. grade per tonne by pit.")

# ======================================================================================
# 1.5 Introduction to Excel
# ======================================================================================
s = b.sheet("1.5 Excel Interface", "1.5  Introduction to Microsoft Excel",
            "The grid, file formats, the Name Box and named cells, and a clean workbook structure", "1.5.1 - 1.5.17")
s.band("1.5.4", "File formats - the same data saved different ways")
files = [("SuperMart_Stores_Sales_2022_2025.csv", CORE), ("SuperMart_Online_RAW_Export.xlsx", W2D),
         ("SuperMart_Sales_2025_DIRTY.csv", W2D), ("supermart_online.db", W3D), ("Lifecare_Admissions_2019_2024.csv", CORE),
         ("lifecare_hospital.db", W3D), ("SuperMart_Online_Sellers.json", W2D), ("SuperMart_Monthly_Sales_Bulletin_2025.pdf", W2D)]
frow = []
for fn, d in files:
    p = d / fn
    if p.exists():
        frow.append([fn, p.suffix, p.stat().st_size / 1024, {".csv": "Plain text - values only, no formulas or formats",
                                                             ".xlsx": "Excel workbook - sheets, formulas, formats",
                                                             ".db": "SQLite database - related tables (Week 3)",
                                                             ".json": "Nested text - web/API exports (Week 2)",
                                                             ".pdf": "Fixed layout - for reading, not analysis"}.get(p.suffix, "")])
fdf = pd.DataFrame(frow, columns=["File", "Format", "Size_KB", "What it holds"])
tf = s.table(fdf, "tFiles", fmts={"Size_KB": "int"}, widths={"File": 38, "What it holds": 40})
s.warn("Opening a CSV by double-click lets Excel guess types: Order_ID-style codes survive, but dates like 03/04/2025 may flip day and month. Use Data > From Text/CSV (2.1).")

s.band("1.5.11", "The Name Box and named cells - VAT at 7.5%")
s.scenario("Click the yellow cell and look at the Name Box (left of the formula bar): it says VAT_Rate, not D-something. "
           "Named cells make formulas read like sentences.")
vr = s.inputs([("VAT rate (Nigeria)", 0.075, "pct", "Named VAT_Rate - see Formulas > Name Manager")])
b.wb.define_name("VAT_Rate", f"='1.5 Excel Interface'!{vr['VAT rate (Nigeria)']}")
ex = s25.nlargest(5, "Sales")[["Order_ID", "Store_Name", "Product_Name", "Sales"]].reset_index(drop=True)
s.table(ex, "tVatDemo", fmts={"Sales": "naira"},
        calc={"VAT": ("=[@Sales]*VAT_Rate", "naira", (ex.Sales * 0.075).round(2).tolist()),
              "Gross": ("=[@Sales]+[@VAT]", "naira", (ex.Sales * 1.075).round(2).tolist())},
        widths={"Product_Name": 40, "Store_Name": 22, "Order_ID": 16})
s.band("1.5.14", "The grid - how big is a worksheet, and do our datasets fit?")
s.calc([("Rows in a worksheet", "=ROWS(A:A)", 1048576, "int", "Every modern Excel version."),
        ("Columns in a worksheet", "=COLUMNS(1:1)", 16384, "int", "A to XFD."),
        ("Largest dataset - diabetes screening rows", 100000, 100000, "int", "Fits comfortably."),
        ("Share of the grid it uses", "=100000/ROWS(A:A)", 100000 / 1048576, "pct", "Above ~1m rows, use Power Query + Data Model (Weeks 2-3).")])
s.band("1.5.16", "Workbook structure we use all bootcamp")
s.grid([["00_README", "Purpose, author, sources, last refresh, how to use"], ["01_Data", "Raw data as Excel Tables - never typed over"],
        ["02_Calc", "Helper columns, lookups, checks"], ["03_Report", "KPIs, charts, the page people read"],
        ["99_Checks", "Control totals and PASS/FAIL flags"]], header=["Sheet", "What goes on it"], merge_last=True)

# ======================================================================================
# 1.6 Working with Data
# ======================================================================================
s = b.sheet("1.6 Working with Data", "1.6  Working with Data in Excel",
            "Data types, Flash Fill as formulas, formatting, conditional formatting, validation", "1.6.1 - 1.6.14")
s.band("1.6.1-2", "The six data types - what Excel really sees in a messy column")
vals = [("₦1,234,560", "Sales typed with symbol"), ("1234560", "Sales stored as text (imported)"), (1234560, "True number"),
        ("2025-03-14", "ISO date stored as text"), ("TRUE", "Logical typed as text"), ("#N/A", "Error text")]
s.ws.write(s.r, 1, "Value as entered", b.F["hdr"]); s.ws.write(s.r, 2, "Description", b.F["hdr"])
for j, h_ in enumerate(["ISNUMBER", "ISTEXT", "TYPE (1=num 2=text)", "LEN"]):
    s.ws.write(s.r, 4 + j, h_, b.F["hdr"])
s.ws.merge_range(s.r, 2, s.r, 3, "Description", b.F["hdr"]); s.r += 1
for v, d in vals:
    s._w(s.r, 1, v, b.F["cell"]); s.ws.merge_range(s.r, 2, s.r, 3, d, b.F["cell"])
    a = s.addr(s.r, 1)
    isn = isinstance(v, (int, float))
    s.ws.write_formula(s.r, 4, f"=ISNUMBER({a})", b.F["cell"], isn)
    s.ws.write_formula(s.r, 5, f"=ISTEXT({a})", b.F["cell"], not isn)
    s.ws.write_formula(s.r, 6, f"=TYPE({a})", b.F["cell"], 1 if isn else 2)
    s.ws.write_formula(s.r, 7, f"=LEN({a})", b.F["cell"], len(str(v)))
    s.r += 1
s.r += 1
s.insight("Only one of six 'numbers' is a number. SUM would ignore the rest without warning - always check ISNUMBER on imported columns.")

s.band("1.6.7", "Flash Fill (Ctrl+E) - and the formula that does the same job")
s.scenario("Lifecare stores patient names as 'First M. Surname'. Flash Fill can split them after you type one example, "
           "but the result is static. The formula versions recalculate when names change.")
pn = lc.Patient_Name.drop_duplicates().head(8).reset_index(drop=True)
pdf = pd.DataFrame({"Patient_Name": pn})
first = [x.split(" ")[0] for x in pn]
lastn = [x.split(" ")[-1] for x in pn]
s.table(pdf, "tNames",
        calc={"First_Name": ('=TEXTBEFORE([@Patient_Name]," ")', None, first),
              "Surname": ('=TEXTAFTER([@Patient_Name]," ",-1)', None, lastn),
              "First_Name_Classic": ('=LEFT([@Patient_Name],FIND(" ",[@Patient_Name])-1)', None, first),
              "Display_Name": ('=UPPER([@Surname])&", "&[@First_Name]', None, [f"{l_.upper()}, {f}" for f, l_ in zip(first, lastn)])},
        widths={"Patient_Name": 24, "Display_Name": 24, "First_Name_Classic": 18})
s.para("TEXTBEFORE/TEXTAFTER need Microsoft 365 / Excel 2024; LEFT+FIND works in every version.", "note")

s.band("1.6.8-9", "Formatting changes the look, never the value")
v = float(s25.Sales.sum())
s.ws.write(s.r, 1, "Stored value", b.F["hdr"]); s.ws.merge_range(s.r, 2, s.r, 3, "Number format", b.F["hdr"])
s.ws.merge_range(s.r, 4, s.r, 5, "How it displays", b.F["hdr"]); s.ws.merge_range(s.r, 6, s.r, 8, "TEXT() formula equivalent", b.F["hdr"]); s.r += 1
for fmt_code, txt_fmt, shown in [("#,##0", "#,##0", f"{v:,.0f}"), ("₦#,##0", "₦#,##0", f"₦{v:,.0f}"),
                                 ('₦#,##0.0,,"m"', '₦#,##0.0,,"m"', f"₦{v / 1e6:,.1f}m"), ("0.00E+00", "0.00E+00", f"{v:.2E}")]:
    s.ws.write_number(s.r, 1, v, b.nf("gen"))
    s.ws.merge_range(s.r, 2, s.r, 3, fmt_code, b.F["cell"])
    s.ws.merge_range(s.r, 4, s.r, 5, v, b.nf(fmt_code))
    s.ws.merge_range(s.r, 6, s.r, 8, "", b.F["cell"])
    s.ws.write_formula(s.r, 6, f'=TEXT({s.addr(s.r, 1)},"{txt_fmt.replace(chr(34), chr(34) * 2)}")', b.F["cell"], shown)
    s.r += 1
s.r += 1
s.warn("TEXT() returns text - you cannot SUM its output. Format cells for display; keep the underlying value numeric.")

s.band("1.6.11", "Conditional formatting - FY2025 attainment heat-map")
ca = SS[["Store_Name", "Sales", "Target"]].copy()
ca["Attainment"] = ca.Sales / ca.Target
top_cf = s.r + 1
tca = s.table(ca, "tCF", fmts={"Sales": "naira", "Target": "naira", "Attainment": "pct"}, widths={"Store_Name": 26})
s.ws.conditional_format(tca.rng("Attainment", absolute=False), {"type": "3_color_scale", "min_color": "#F8696B", "mid_color": "#FFEB84", "max_color": "#63BE7B",
                                                                "mid_type": "num", "mid_value": 1})
s.ws.conditional_format(tca.rng("Sales", absolute=False), {"type": "data_bar", "bar_color": "#13807A", "bar_solid": True})
s.para("Applied: data bars on Sales; a 3-colour scale on Attainment with the midpoint fixed at 100% (Home > Conditional Formatting > Manage Rules).", "note")

s.band("1.6.14", "Data validation - stop bad zones being typed in the first place")
s.scenario("Type into the yellow cells. The list only accepts the four zones where SuperMart trades; the quantity cell rejects anything that is not a whole number from 1 to 500.")
s.ws.merge_range(s.r, 1, s.r, 3, "Zone (pick from list)", b.F["lab"]); s.ws.write(s.r, 4, "South-West", b.nf("gen", style="input"))
s.ws.data_validation(s.r, 4, s.r, 4, {"validate": "list", "source": zones, "error_title": "Unknown zone", "error_message": "Choose one of the four SuperMart zones."})
s.r += 1
s.ws.merge_range(s.r, 1, s.r, 3, "Quantity (1-500)", b.F["lab"]); s.ws.write(s.r, 4, 3, b.nf("int", style="input"))
s.ws.data_validation(s.r, 4, s.r, 4, {"validate": "integer", "criteria": "between", "minimum": 1, "maximum": 500,
                                     "error_title": "Invalid quantity", "error_message": "Whole number from 1 to 500. Negative quantities are returns - use the returns form."})
s.r += 2
s.tryit("Type 'SW' or -2 into the yellow cells and read the error message. The dirty extract has 8 negative quantities that validation would have blocked.")

# ======================================================================================
# 1.7 Formulas and core functions
# ======================================================================================
s = b.sheet("1.7 Formulas & Functions", "1.7  Excel Formulas and Core Functions",
            "Operators and BODMAS, relative/absolute/mixed references, the six essentials, errors", "1.7.1 - 1.7.14")
s.band("1.7.1-5", "Formula anatomy on one SuperMart line")
line = s25[(s25.Discount == 0.2) & (s25.Quantity >= 3)].iloc[0]
inp = s.inputs([("Sales (net, ₦)", int(line.Sales), "naira", f"Order {line.Order_ID}, {line.Product_Name[:40]}"),
                ("Quantity", int(line.Quantity), "int", ""), ("Discount", float(line.Discount), "pct0", ""), ("Profit (₦)", int(line.Profit), "naira", "")])
S_, Q_, D_, P_ = (inp[k] for k in ["Sales (net, ₦)", "Quantity", "Discount", "Profit (₦)"])
gross = line.Sales / (1 - line.Discount)
s.calc([("Net price per unit", f"={S_}/{Q_}", line.Sales / line.Quantity, "naira", "Division operator /"),
        ("List value before discount", f"={S_}/(1-{D_})", gross, "naira", "Brackets first (BODMAS): 1-20% = 80%, then divide."),
        ("WRONG: no brackets", f"={S_}/1-{D_}", line.Sales - line.Discount, "naira", "Excel divides by 1 then subtracts 0.2 - a silent error."),
        ("Discount given (₦)", f"={S_}/(1-{D_})-{S_}", gross - line.Sales, "naira", ""),
        ("Cost of goods", f"={S_}-{P_}", line.Sales - line.Profit, "naira", "Subtraction"),
        ("Margin", f"={P_}/{S_}", line.Profit / line.Sales, "pct", ""),
        ("Exponent demo: growth over 3 yrs at 10%", "=(1+10%)^3-1", 1.1 ** 3 - 1, "pct", "Exponent ^ is evaluated before * and /")])

s.band("1.7.6-7", "Relative, absolute and mixed references - a price-increase grid")
s.scenario("Head office is testing price increases of 2.5%-10% on FY2025 category sales. One formula in the top-left of the grid, "
           "copied across and down, fills every cell - because of the $ signs: =$B{row}*(1+C$top).")
pct_row = s.r
s.ws.write(s.r, 1, "Category ↓ / Increase →", b.F["hdr"]); s.ws.write(s.r, 2, "FY2025 Sales", b.F["hdr"])
incs = [0.025, 0.05, 0.075, 0.10]
for j, x in enumerate(incs):
    s.ws.write_number(s.r, 3 + j, x, b.nf("pct", style="input"))
s.r += 1
for _, rr in cg.iterrows():
    s.ws.write(s.r, 1, rr.Category, b.F["lab"]); s.ws.write_number(s.r, 2, rr.Sales, b.nf("naira"))
    for j, x in enumerate(incs):
        f = f"=$C{s.r + 1}*(1+{chr(68 + j)}${pct_row + 1})"
        s.ws.write_formula(s.r, 3 + j, f, b.nf("naira"), rr.Sales * (1 + x))
    s.r += 1
s.ws.merge_range(s.r, 1, s.r, 6, f"Cell D{pct_row + 2} contains  =$C{pct_row + 2}*(1+D${pct_row + 1})  - $C locks the column, $row locks the row. Press F4 to cycle.", b.F["ftext"])
s.r += 2

s.band("1.7.9-11", "AutoSum and the six essential functions - 20 Lifecare bills")
bills = lc[lc.Department == "Cardiology"].head(20)[["Admission_ID", "Patient_Name", "Admission_Type", "Length_of_Stay", "Billing_Amount"]].reset_index(drop=True)
bills.loc[[4, 11], "Billing_Amount"] = np.nan  # two bills not yet finalised
tb = s.table(bills, "tBills", fmts={"Billing_Amount": "naira2", "Length_of_Stay": "int"}, widths={"Patient_Name": 22, "Admission_ID": 16})
bb = bills.Billing_Amount
s.calc([("SUM", "=SUM(tBills[Billing_Amount])", bb.sum(), "naira", "Total billed so far"),
        ("AVERAGE", "=AVERAGE(tBills[Billing_Amount])", bb.mean(), "naira", "Ignores the 2 blank (not finalised) bills"),
        ("MIN", "=MIN(tBills[Billing_Amount])", bb.min(), "naira", ""),
        ("MAX", "=MAX(tBills[Billing_Amount])", bb.max(), "naira", ""),
        ("COUNT", "=COUNT(tBills[Billing_Amount])", int(bb.count()), "int", "Counts numbers only: 18"),
        ("COUNTA", "=COUNTA(tBills[Admission_ID])", 20, "int", "Counts anything non-blank: 20 admissions"),
        ("Bills outstanding", "=COUNTA(tBills[Admission_ID])-COUNT(tBills[Billing_Amount])", 2, "int", "COUNTA - COUNT = blanks to chase")])

s.band("1.7.12", "Understanding and fixing formula errors")
s.scenario("Each row deliberately produces an error, then shows the fix. The 'With error' column really errors - that is the point.")
s.ws.merge_range(s.r, 1, s.r, 2, "Error", b.F["hdr"]); s.ws.merge_range(s.r, 3, s.r, 5, "Cause (our data)", b.F["hdr"])
s.ws.write(s.r, 6, "With error", b.F["hdr"]); s.ws.merge_range(s.r, 7, s.r, 9, "Fixed formula", b.F["hdr"]); s.ws.write(s.r, 10, "Fixed result", b.F["hdr"]); s.r += 1
zero_orders_row = s.r
errs = [("#DIV/0!", "Margin for a new store with zero sales", "=0/0", '=IFERROR(0/0,"n/a - no sales")', "n/a - no sales", None),
        ("#VALUE!", "Adding a text amount '₦25,000' to a number", '="₦25,000"+1000', '=VALUE(SUBSTITUTE(SUBSTITUTE("₦25,000","₦",""),",",""))+1000', 26000, "naira"),
        ("#N/A", "Looking up store ST13, which does not exist", '=MATCH("ST13",tStores25[Store_ID],0)', '=IFNA(MATCH("ST13",tStores25[Store_ID],0),"not found")', "not found", None),
        ("#NAME?", "Misspelt function name", "=SUMM(1,2)", "=SUM(1,2)", 3, "int"),
        ("#REF!", "Referring to a deleted row (simulated)", "=#REF!+1", "=1+1", 2, "int")]
for e, cause, bad, good, res, kind in errs:
    s.ws.merge_range(s.r, 1, s.r, 2, e, b.F["lab"]); s.ws.merge_range(s.r, 3, s.r, 5, cause, b.F["cell"])
    s.ws.write_formula(s.r, 6, bad, b.nf("gen", style="bad"), e)
    s.mtext(s.r, 7, 9, good, b.F["ftext"])
    s.ws.write_formula(s.r, 10, good, b.nf(kind or "gen", style="result"), res)
    s.r += 1
s.r += 1

s.band("1.7.14", "Practical multi-sector calculations")
occ = L("lc_census")
occ23 = occ[pd.to_datetime(occ.Census_Date).dt.year == 2023]
h24 = ht[(ht.Arrival_Date.dt.year == 2024) & (ht.Booking_Status == "Honoured")]
rooms, nights = 120, 366
room_rev = h24.Booking_Value_NGN.sum()
inp = s.inputs([("SuperMart FY2025 profit (₦)", int(SS.Profit.sum()), "naira", ""), ("SuperMart FY2025 sales (₦)", int(SS.Sales.sum()), "naira", ""),
                ("Lifecare bed-days used, 2023", int(occ23.Patients_In_House.sum()), "int", "Sum of daily census"),
                ("Lifecare staffed beds", 520, "int", ""), ("Days in 2023", 365, "int", ""),
                ("Crestview students passing (score ≥ 50)", int((cv.Exam_Score >= 50).sum()), "int", ""), ("Crestview students", len(cv), "int", ""),
                ("Eko Suites room revenue 2024 (₦, non-cancelled)", float(room_rev), "naira", ""), ("Eko Suites rooms", rooms, "int", "Assumed inventory"),
                ("Nights in 2024", nights, "int", "Leap year")])
k = list(inp.values())
s.calc([("Retail: profit margin", f"={k[0]}/{k[1]}", SS.Profit.sum() / SS.Sales.sum(), "pct", ""),
        ("Hospital: average bed occupancy 2023", f"={k[2]}/({k[3]}*{k[4]})", occ23.Patients_In_House.sum() / (520 * 365), "pct", "Bed-days used ÷ bed-days available"),
        ("University: pass rate", f"={k[5]}/{k[6]}", (cv.Exam_Score >= 50).mean(), "pct", ""),
        ("Hotel: RevPAR (revenue per available room-night)", f"={k[7]}/({k[8]}*{k[9]})", room_rev / (rooms * nights), "naira", "The hotel industry's headline KPI")])

# ======================================================================================
# 1.8 Excel Tables
# ======================================================================================
s = b.sheet("1.8 Excel Tables", "1.8  Organising Data with Excel Tables",
            "Records and fields, Ctrl+T, structured references, calculated columns, total row, SUBTOTAL", "1.8.1 - 1.8.12")
s.band("1.8.1-4", "An official Excel Table: Port Harcourt - GRA, December 2025")
s.scenario("Click any cell in the table and the Table Design tab appears. Each ROW is one record (a sales line); each COLUMN is one field. "
           "The table is named tPH_Dec - formulas below refer to it by name, so they keep working when rows are added.")
ph = s25[(s25.Store_ID == "ST04") & (s25.Order_Date.dt.month == 12)][
    ["Order_ID", "Order_Date", "Customer_Name", "Segment", "Category", "Sub_Category", "Sales", "Quantity", "Discount", "Profit"]].reset_index(drop=True)
tph = s.table(ph, "tPH_Dec", fmts={"Order_Date": "date", "Sales": "naira", "Profit": "naira", "Discount": "pct0", "Quantity": "int"},
              calc={"Unit_Price": ("=[@Sales]/[@Quantity]", "naira", (ph.Sales / ph.Quantity).tolist()),
                    "Margin": ('=IF([@Sales]=0,"",[@Profit]/[@Sales])', "pct", (ph.Profit / ph.Sales).tolist())},
              total={"Order_ID": "Total", "Sales": "sum", "Quantity": "sum", "Profit": "sum"},
              widths={"Customer_Name": 18, "Order_ID": 16, "Sub_Category": 13, "Category": 14})
s.band("1.8.6", "Structured references - formulas that read like English")
cats_ph = ph.groupby("Category").Sales.sum()
s.calc([("Lines in the table", "=ROWS(tPH_Dec)", len(ph), "int", "Grows automatically when you paste new rows under the table"),
        ("Total December sales", "=SUM(tPH_Dec[Sales])", ph.Sales.sum(), "naira", ""),
        ("Technology sales", '=SUMIFS(tPH_Dec[Sales],tPH_Dec[Category],"Technology")', cats_ph.get("Technology", 0), "naira", ""),
        ("Lines sold at a loss", '=COUNTIF(tPH_Dec[Profit],"<0")', int((ph.Profit < 0).sum()), "int", ""),
        ("Corporate-segment average sale", '=AVERAGEIFS(tPH_Dec[Sales],tPH_Dec[Segment],"Corporate")',
         ph[ph.Segment == "Corporate"].Sales.mean() if (ph.Segment == "Corporate").any() else 0, "naira", ""),
        ("Visible-rows total (respects filters)", "=SUBTOTAL(109,tPH_Dec[Sales])", ph.Sales.sum(), "naira",
         "Filter the Category column - this changes, SUM does not. 109 = SUM ignoring hidden rows.")])
s.band("1.8.7-10", "Sort, filter and slice - do it on the table above")
s.bullets(["Sort Sales largest-to-smallest: the biggest line is at the top - is it a Technology order?",
           "Filter Profit < 0 (Number Filters > Less Than): which sub-categories lose money in December?",
           "Table Design > Insert Slicer > Category: click 'Furniture' and watch the SUBTOTAL result change.",
           "Table Design > Total Row: switch the Profit total to Average - Excel writes SUBTOTAL(101, ...) for you."])
s.band("1.8.11", "Naming conventions used in every ETEH workbook")
s.grid([["Tables", "t + Subject (+ Scope)", "tSales, tPH_Dec, tStores25"], ["Named cells", "Words_With_Underscores", "VAT_Rate, Bed_Count"],
        ["Sheets", "NN_Purpose", "01_Data, 03_Report"], ["Columns", "No spaces, units in name", "Billing_Amount, Length_of_Stay, Sales_NGN"]],
       header=["Object", "Pattern", "Examples"])

# ======================================================================================
# 1.9 Tidy data laws
# ======================================================================================
s = b.sheet("1.9 Tidy Data Laws", "1.9  Data Entry Best Practices: The Twelve Tidy Data Laws",
            "Before/after examples built from the bootcamp's own messy extracts", "1.9.1 - 1.9.12")
s.band("1.9.2-5", "Laws 1, 2, 4: one record per row, one variable per column, one header row")
s.scenario("Store managers often send 'wide' reports: one column per month. It looks tidy to a human but every month is a separate variable. "
           "The long (tidy) version has one row per Store x Month - which PivotTables, SUMIFS and Power Query all expect.")
s25["Month"] = s25.Order_Date.dt.month
wide = s25[s25.Store_ID.isin(["ST01", "ST04", "ST07", "ST10"])].pivot_table(index="Store_Name", columns="Month", values="Sales", aggfunc="sum").iloc[:, :4]
wide.columns = ["Jan", "Feb", "Mar", "Apr"]
wide = wide.reset_index()
s.h2("BEFORE - wide (one column per month)")
s.frame(wide, fmts={m: "naira" for m in ["Jan", "Feb", "Mar", "Apr"]})
long_ = wide.melt(id_vars="Store_Name", var_name="Month", value_name="Sales")
s.h2("AFTER - long / tidy (Store, Month, Sales)")
tl = s.table(long_, "tLong", fmts={"Sales": "naira"}, widths={"Store_Name": 28})
s.calc([("Check: wide total = long total", "=SUM(tLong[Sales])", long_.Sales.sum(), "naira", "Reshaping must never change the total (reconciliation)."),
        ("Ikeja, March (now one simple formula)", '=SUMIFS(tLong[Sales],tLong[Store_Name],"Lagos - Ikeja",tLong[Month],"Mar")',
         long_[(long_.Store_Name == "Lagos - Ikeja") & (long_.Month == "Mar")].Sales.sum(), "naira", "")])

s.band("1.9.4", "Law 3: atomic values - 'Benin City, Edo' is two facts")
loc = dirty["Store_Location"].drop_duplicates().head(6).reset_index(drop=True)
ldf = pd.DataFrame({"Store_Location": loc})
s.table(ldf, "tLoc", calc={"City": ('=TEXTBEFORE([@Store_Location],", ")', None, [x.split(", ")[0] for x in loc]),
                           "State": ('=TEXTAFTER([@Store_Location],", ")', None, [x.split(", ")[1] for x in loc])},
        widths={"Store_Location": 26})

s.band("1.9.7", "Law 6: ISO dates - one real order date typed three ways")
dts = [("03/04/2025", "UK reading: 3 April 2025. US reading: 4 March 2025"), ("03-Apr-2025", "Unambiguous, but text"), ("2025.04.03", "Text - Excel will not sort it as a date")]
s.ws.merge_range(s.r, 1, s.r, 2, "As typed (text)", b.F["hdr"]); s.ws.merge_range(s.r, 3, s.r, 6, "Problem", b.F["hdr"])
s.ws.merge_range(s.r, 7, s.r, 9, "Repair formula", b.F["hdr"]); s.ws.write(s.r, 10, "ISO result", b.F["hdr"]); s.r += 1
import datetime as _dt
fixes = ['=DATE(RIGHT({a},4),MID({a},4,2),LEFT({a},2))', '=DATEVALUE({a})', '=DATE(LEFT({a},4),MID({a},6,2),RIGHT({a},2))']
for (txt, prob), fx in zip(dts, fixes):
    s.ws.merge_range(s.r, 1, s.r, 2, txt, b.nf("text"))
    s.ws.merge_range(s.r, 3, s.r, 6, prob, b.F["cell"])
    a = s.addr(s.r, 1)
    s.mtext(s.r, 7, 9, fx.format(a=a), b.F["ftext"])
    s.ws.write_formula(s.r, 10, fx.format(a=a), b.nf("date", style="result"), 45750)
    s.r += 1
s.r += 1

s.band("1.9.8", "Law 7: blank is not zero - Lifecare bills again")
s.scenario("Two Cardiology bills (sheet 1.7) are blank because they are not finalised. If someone 'tidies' them to 0, the average bill drops - and the hospital under-reports revenue per patient.")
bz = bills.Billing_Amount.fillna(0)
s.calc([("Average bill - blanks left blank", "=AVERAGE('1.7 Formulas & Functions'!" + tb.rng("Billing_Amount") + ")", bb.mean(), "naira", "Correct: unknowns are excluded."),
        ("Average bill - blanks typed as 0", "=SUM('1.7 Formulas & Functions'!" + tb.rng("Billing_Amount") + ")/20", bz.mean(), "naira", "Wrong: understates by ~10%."),
        ("Blanks to chase", "=COUNTBLANK('1.7 Formulas & Functions'!" + tb.rng("Billing_Amount") + ")", 2, "int", "Report the count - do not hide it.")])

s.band("1.9.10", "Law 9: controlled vocabulary - the same ward, five spellings")
dl_ = L("dirty_lc")
wv = dl_.Ward.value_counts().reset_index()
wv.columns = ["Ward_as_typed", "Rows"]
wv = wv.head(16)
canon = {"emerald": "Emerald Ward", "sapphire": "Sapphire Ward", "saphire": "Sapphire Ward", "ruby": "Ruby Ward", "onyx": "Onyx Ward",
         "topaz": "Topaz Ward", "diamond": "Diamond Ward", "jade": "Jade Ward", "pearl": "Pearl Ward"}


def _canon(x):
    k = " ".join(w for w in x.lower().split() if w != "ward")
    return canon.get(k, x.strip().title())


wv["Standard_Ward"] = [_canon(x) for x in wv.Ward_as_typed]
s.table(wv, "tWardVariants", fmts={"Rows": "int"}, widths={"Ward_as_typed": 22, "Standard_Ward": 18},
        calc={"Length": ("=LEN([@Ward_as_typed])", "int", [len(x) for x in wv.Ward_as_typed]),
              "Is_Standard": ("=EXACT([@Ward_as_typed],[@Standard_Ward])", None, [a == b_ for a, b_ in zip(wv.Ward_as_typed, wv.Standard_Ward)])})
s.insight(f"Typed variants add {int(wv[wv.Ward_as_typed != wv.Standard_Ward].Rows.sum()):,} rows that a PivotTable would show as extra 'wards'. "
          "LEN exposes the invisible trailing space: 'Emerald Ward ' is 13 characters, not 12.")
s.para("The Standard_Ward column is the mapping table. Week 2 uses it with XLOOKUP (2.8.58) and Power Query Merge (3.4) instead of hand-editing.", "note")

s.band("1.9.11", "Law 10: store calculations, not typed results")
s.scenario("A typed total looks the same as a formula total - until the data changes. Below, someone typed last month's total; the formula is live.")
s.calc([("Typed total (static)", int(cg.Sales.sum() - 1250000), int(cg.Sales.sum() - 1250000), "naira", "Someone typed this from an old printout."),
        ("Formula total", "=SUM(tCat25[Sales])".replace("tCat25", "'1.3 Analytics Process'!tCat25") if False else "=SUM(tCat25[Sales])", cg.Sales.sum(), "naira", "Live - always right."),
        ("Difference", f"=SUM(tCat25[Sales])-{int(cg.Sales.sum() - 1250000)}", 1250000, "naira", "A ₦1.25m silent error.")])
s.band("1.9.12", "Laws 11-12: no merged cells; document everything")
s.grid([["Merged header 'Q1 2025' over three month columns", "Sorting fails ('requires merged cells to be identically sized'); PivotTables see blank headers", "One header per column: Jan_2025, Feb_2025... or go long"],
        ["No README", "Nobody knows the Sales column is net of discount", "00_README sheet + data dictionary (00_Data_Dictionary_and_Catalogue.xlsx)"]],
       header=["Bad practice", "What goes wrong", "Tidy fix"])

b.cover("Week 1 Section Examples - Foundations of Data Analytics and Excel",
        ["One worksheet per module (1.1 - 1.9). Each worked example is tagged with the slide topic number it supports.",
         "All numbers come from the bootcamp datasets in 00_Core_Datasets - nothing is invented. Every result cell is a live formula; the blue "
         "column beside it shows the exact formula text so you can read it without clicking.",
         "Instructor tip: project the sheet while teaching the matching slides, then let learners change the yellow cells."],
        [(k, DATASETS_TEXT[k]) for k in ["SuperMart Stores", "SuperMart Online", "Lifecare", "Crestview", "Salary Survey", "Eko Suites"]])
b.close()
print("Week 1 ->", OUT)
