"""Week 4 Section Examples workbook (modules 4.1 - 4.14): statistics, regression, forecasting, what-if, storytelling,
reporting, auditing, AI co-pilots, verification, ethics and internal benchmarks."""
from __future__ import annotations

import os
import sqlite3
import sys

import numpy as np
import pandas as pd
from scipy import optimize, stats
from xlsxwriter.utility import xl_cell_to_rowcol, xl_rowcol_to_cell

sys.path.insert(0, os.path.dirname(__file__))
from common import SEC, W3D  # noqa: E402
from wk_common import DATASETS_TEXT, L, hotel, lc_adm, sm_sales, smo  # noqa: E402
from xlhelp import Book  # noqa: E402

OUT = SEC / "ETEH_Week4_Section_Examples.xlsx"
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
XL0 = pd.Timestamp("1899-12-30")


def serial(d):
    return float((pd.Timestamp(d) - XL0).days)


def A(sheet, addr):
    """Absolute, sheet-qualified address from a relative 'E12'."""
    r, c = xl_cell_to_rowcol(addr)
    return f"'{sheet}'!{xl_rowcol_to_cell(r, c, True, True)}"


def AB(addr):
    r, c = xl_cell_to_rowcol(addr.replace("$", ""))
    return xl_rowcol_to_cell(r, c, True, True)


# ---------------------------------------------------------------------- data
sm = sm_sales()
sm["Year"] = sm.Order_Date.dt.year
s25 = sm[sm.Year == 2025]
s24 = sm[sm.Year == 2024]
stores = L("sm_stores")
sal = L("sal_clean").copy()
sal["Is_Masters"] = (sal.Education_Level == "Master's").astype(int)
sal["Is_PhD"] = (sal.Education_Level == "PhD").astype(int)
sal = sal[["Respondent_ID", "Gender", "Education_Level", "Job_Title", "Job_Family", "Seniority", "Age", "Years_of_Experience",
           "Is_Masters", "Is_PhD", "Annual_Salary_NGN"]]
cv = L("cv_perf")
cvd = cv[["Student_ID", "Faculty", "Gender", "Hours_Studied_Per_Week", "Attendance_Pct", "Sleep_Hours", "Previous_Score",
          "Tutoring_Sessions", "Physical_Activity_Hrs", "Exam_Score", "Grade", "At_Risk"]].copy()
lc = lc_adm()
los = lc[lc.Admission_Date.dt.year == 2023][["Admission_ID", "Admission_Date", "Department", "Admission_Type", "Age", "Gender",
                                             "Length_of_Stay", "Billing_Amount", "HMO", "Readmit_30d"]].reset_index(drop=True)
census = L("lc_census").copy()
census["Census_Date"] = pd.to_datetime(census.Census_Date)
c23 = census[census.Census_Date.dt.year == 2023][["Census_Date", "Patients_In_House", "Occupancy_Pct"]].reset_index(drop=True)
o = smo()
tx, pr, cu = o["transactions"], o["products"], o["customers"]
orders = tx.merge(pr[["product_id", "category"]], on="product_id")[["order_id", "order_date", "customer_id", "category", "quantity",
                                                                     "total_price", "order_status"]]
ho = hotel()
ho["Arrival_Year"] = ho.Arrival_Date.dt.year
hod = ho[["Booking_ID", "Arrival_Date", "Arrival_Year", "Room_Type", "Market_Segment", "Lead_Time_Days", "Total_Nights",
          "Special_Requests", "Repeat_Guest", "Avg_Room_Rate_NGN", "Booking_Status", "Booking_Value_NGN"]].copy()
ROOMS = 320

b = Book(str(OUT), "ETEH Week 4 - Section Examples", "Week 4 · Statistics, Forecasting, What-If, Communication, AI and Ethics")
F = b.F

# ======================================================================================
# 4.1 Descriptive statistics
# ======================================================================================
SH = "4.1 Descriptive Stats"
s = b.sheet(SH, "4.1  Descriptive Statistics and the Analysis ToolPak",
            "Centre, spread, shape and outliers - salaries, order values and hospital stays", "4.1.1 - 4.1.14")
s.scenario("ETEH Talent's salary survey (1,781 professionals, D_Salary) and SuperMart Online's 11,000 orders (D_Orders). "
           "Every number is a live formula over the Excel Tables - change a value in the data and the statistics update.")
y = sal.Annual_Salary_NGN
ov = orders.total_price
s.band("4.1.2-4", "Centre: mean, median and mode")
s.calc([("Mean salary", "=AVERAGE(tSal[Annual_Salary_NGN])", y.mean(), "naira", "Pulled up by a few executive salaries"),
        ("Median salary", "=MEDIAN(tSal[Annual_Salary_NGN])", y.median(), "naira", "Half earn less, half earn more"),
        ("Most common salary (mode)", "=MODE.SNGL(tSal[Annual_Salary_NGN])", float(y.mode().iloc[0]), "naira", "Survey answers cluster on round numbers"),
        ("Mean order value - SuperMart Online", "=AVERAGE(tOrders[total_price])", ov.mean(), "naira", ""),
        ("Median order value", "=MEDIAN(tOrders[total_price])", ov.median(), "naira", "Mean is over 4x the median: a few big electronics orders"),
        ("Mean ÷ median", "=AVERAGE(tOrders[total_price])/MEDIAN(tOrders[total_price])", ov.mean() / ov.median(), "dec2", "> 1.2 → report the median as 'typical'")])
s.band("4.1.5-7", "Spread: range, IQR, variance, standard deviation, coefficient of variation")
q1, q3 = y.quantile(.25), y.quantile(.75)
s.calc([("Minimum", "=MIN(tSal[Annual_Salary_NGN])", y.min(), "naira", ""),
        ("Maximum", "=MAX(tSal[Annual_Salary_NGN])", y.max(), "naira", ""),
        ("Range", "=MAX(tSal[Annual_Salary_NGN])-MIN(tSal[Annual_Salary_NGN])", y.max() - y.min(), "naira", "Sensitive to one extreme value"),
        ("Q1 (25th percentile)", "=QUARTILE.INC(tSal[Annual_Salary_NGN],1)", q1, "naira", ""),
        ("Q3 (75th percentile)", "=QUARTILE.INC(tSal[Annual_Salary_NGN],3)", q3, "naira", ""),
        ("Interquartile range (IQR)", "=QUARTILE.INC(tSal[Annual_Salary_NGN],3)-QUARTILE.INC(tSal[Annual_Salary_NGN],1)", q3 - q1, "naira", "The middle 50% spread - robust"),
        ("Sample variance", "=VAR.S(tSal[Annual_Salary_NGN])", y.var(), "int", "In ₦² - hard to read, used inside other formulas"),
        ("Sample standard deviation", "=STDEV.S(tSal[Annual_Salary_NGN])", y.std(), "naira", "Typical distance from the mean"),
        ("Coefficient of variation", "=STDEV.S(tSal[Annual_Salary_NGN])/AVERAGE(tSal[Annual_Salary_NGN])", y.std() / y.mean(), "pct", "Lets you compare spread across measures")])
s.h2("4.1.7  Which is more variable - salaries, order values or hospital stays? (CV removes the units)")
cvt = pd.DataFrame({"Measure": ["Annual salary (₦)", "Online order value (₦)", "Length of stay 2023 (days)", "Exam score (points)"],
                    "Column": ["tSal[Annual_Salary_NGN]", "tOrders[total_price]", "tLOS[Length_of_Stay]", "tCV[Exam_Score]"]})
vals = [y, ov, los.Length_of_Stay, cvd.Exam_Score]
s.table(cvt, "tCVcompare", widths={"Measure": 28, "Column": 26},
        calc={"Mean": ("=AVERAGE(INDIRECT([@Column]))", "dec1", [v.mean() for v in vals]),
              "Std_Dev": ("=STDEV.S(INDIRECT([@Column]))", "dec1", [v.std() for v in vals]),
              "CV": ("=[@Std_Dev]/[@Mean]", "pct", [v.std() / v.mean() for v in vals])})
s.para("INDIRECT turns the text in 'Column' into a live reference - handy for a stats panel, but it is volatile, so avoid it in big models.", "note")
s.band("4.1.8-9", "Shape: skewness and kurtosis")
s.calc([("Skewness - order value", "=SKEW(tOrders[total_price])", stats.skew(ov, bias=False), "dec2", "> 1 = strong right tail (few very large orders)"),
        ("Skewness - salary", "=SKEW(tSal[Annual_Salary_NGN])", stats.skew(y, bias=False), "dec2", "Near 0 = roughly symmetric"),
        ("Skewness - exam score", "=SKEW(tCV[Exam_Score])", stats.skew(cvd.Exam_Score, bias=False), "dec2", "A handful of very high scores"),
        ("Kurtosis - order value", "=KURT(tOrders[total_price])", stats.kurtosis(ov, bias=False), "dec2", "> 0 = heavier tails than normal"),
        ("Kurtosis - exam score", "=KURT(tCV[Exam_Score])", stats.kurtosis(cvd.Exam_Score, bias=False), "dec2", "Very peaked: most students score 63-71")])
s.band("4.1.10", "Percentiles and quartiles - salary bands for HR")
pct_rows = [(f"P{p}", p / 100) for p in (10, 25, 50, 75, 90)]
pt = pd.DataFrame({"Percentile": [a for a, _ in pct_rows], "k": [k for _, k in pct_rows]})
s.table(pt, "tSalPct", fmts={"k": "pct0"}, calc={"Salary": ("=PERCENTILE.INC(tSal[Annual_Salary_NGN],[@k])", "naira", [y.quantile(k) for _, k in pct_rows])})
s.band("4.1.11", "Outliers with the 1.5 × IQR rule - SuperMart Online order values")
oq1, oq3 = ov.quantile(.25), ov.quantile(.75)
oiqr = oq3 - oq1
up = oq3 + 1.5 * oiqr
lo = oq1 - 1.5 * oiqr
o_rows = s.calc([("Q1 order value", "=QUARTILE.INC(tOrders[total_price],1)", oq1, "naira", ""),
                 ("Q3 order value", "=QUARTILE.INC(tOrders[total_price],3)", oq3, "naira", ""),
                 ("IQR", "=QUARTILE.INC(tOrders[total_price],3)-QUARTILE.INC(tOrders[total_price],1)", oiqr, "naira", ""),
                 ("Upper fence = Q3 + 1.5 × IQR", "=QUARTILE.INC(tOrders[total_price],3)+1.5*(QUARTILE.INC(tOrders[total_price],3)-QUARTILE.INC(tOrders[total_price],1))", up, "naira", ""),
                 ("Lower fence = Q1 - 1.5 × IQR", "=QUARTILE.INC(tOrders[total_price],1)-1.5*(QUARTILE.INC(tOrders[total_price],3)-QUARTILE.INC(tOrders[total_price],1))", lo, "naira", "Negative → no low outliers possible")])
s.calc([("Orders above the upper fence", f'=COUNTIFS(tOrders[total_price],">"&{o_rows["Upper fence = Q3 + 1.5 × IQR"]})', int((ov > up).sum()), "int", "Genuine high-value orders, not errors - investigate before removing"),
        ("Share of revenue in those orders", f'=SUMIFS(tOrders[total_price],tOrders[total_price],">"&{o_rows["Upper fence = Q3 + 1.5 × IQR"]})/SUM(tOrders[total_price])',
         ov[ov > up].sum() / ov.sum(), "pct", "Deleting 'outliers' here would delete a big part of the business")], header=None)
s.band("4.1.12-13", "Analysis ToolPak replica - Data > Data Analysis > Descriptive Statistics (with Summary statistics ticked)")
s.para("The ToolPak pastes static numbers. This replica uses live formulas so you can see exactly what each line means.", "note")
cols = {"Annual salary": ("tSal[Annual_Salary_NGN]", y), "Exam score": ("tCV[Exam_Score]", cvd.Exam_Score), "Length of stay 2023": ("tLOS[Length_of_Stay]", los.Length_of_Stay)}
tp_rows = [("Mean", "AVERAGE({c})", lambda v: v.mean()), ("Standard Error", "STDEV.S({c})/SQRT(COUNT({c}))", lambda v: v.std() / np.sqrt(len(v))),
           ("Median", "MEDIAN({c})", lambda v: v.median()), ("Mode", "MODE.SNGL({c})", lambda v: float(v.mode().iloc[0])),
           ("Standard Deviation", "STDEV.S({c})", lambda v: v.std()), ("Sample Variance", "VAR.S({c})", lambda v: v.var()),
           ("Kurtosis", "KURT({c})", lambda v: stats.kurtosis(v, bias=False)), ("Skewness", "SKEW({c})", lambda v: stats.skew(v, bias=False)),
           ("Range", "MAX({c})-MIN({c})", lambda v: v.max() - v.min()), ("Minimum", "MIN({c})", lambda v: v.min()), ("Maximum", "MAX({c})", lambda v: v.max()),
           ("Sum", "SUM({c})", lambda v: v.sum()), ("Count", "COUNT({c})", lambda v: len(v)),
           ("Confidence Level (95.0%)", "CONFIDENCE.T(0.05,STDEV.S({c}),COUNT({c}))", lambda v: stats.t.ppf(.975, len(v) - 1) * v.std() / np.sqrt(len(v)))]
tpd = pd.DataFrame({"Statistic": [r[0] for r in tp_rows], **{k: 0.0 for k in cols}})
fx = {}
for i, (nm, f_, fn) in enumerate(tp_rows):
    for k, (cref, v) in cols.items():
        fx[(i, k)] = ("=" + f_.format(c=cref), float(fn(v)))
s.frame(tpd, fmts={k: "#,##0.00" for k in cols}, formulas=fx, col_widths=None)
s.band("4.1.14", "Reporting descriptive statistics honestly")
s.grid([["Skewed money values", "'Typical online order is ₦" + f"{ov.median():,.0f}" + " (median); mean ₦" + f"{ov.mean():,.0f}" + " is inflated by large electronics orders'"],
        ["Always give n", f"'n = {len(sal):,} respondents' - a mean of 7 'Other' gender respondents is not a finding"],
        ["Spread with centre", f"'Median salary ₦{y.median():,.0f}, IQR ₦{q1:,.0f} - ₦{q3:,.0f}'"],
        ["Round sensibly", "Mean salary ₦11,344,180.18 → ₦11.3m in a report; keep full precision in the workbook"],
        ["State exclusions", "'Cancelled orders included (gross bookings)' or 'excluded (net revenue)' - never leave it ambiguous"]],
       header=["Rule", "Example sentence from our data"], merge_last=True)

# ======================================================================================
# 4.2 Distributions, Pareto, percentiles, z-scores
# ======================================================================================
SH = "4.2 Distributions & Z"
s = b.sheet(SH, "4.2  Distributions, Pareto, Percentiles and Z-Scores",
            "Histogram vs normal curve, 80/20 customers, percentile bands, store z-scores and occupancy anomalies", "4.2.1 - 4.2.9")
s.band("4.2.1-3", "Distribution of Crestview exam scores vs the normal curve")
mu, sd = cvd.Exam_Score.mean(), cvd.Exam_Score.std()
edges = list(range(55, 80, 2)) + [101]
hb = pd.DataFrame({"From": edges[:-1], "To_below": edges[1:]})
hb["Bin"] = [f"{a}-{b_ - 1}" if b_ < 101 else f"{a}+" for a, b_ in zip(hb.From, hb.To_below)]
hb = hb[["Bin", "From", "To_below"]]
obs = [int(((cvd.Exam_Score >= a) & (cvd.Exam_Score < b_)).sum()) for a, b_ in zip(hb.From, hb.To_below)]
expn = [len(cvd) * (stats.norm.cdf(b_ - .5, mu, sd) - stats.norm.cdf(a - .5, mu, sd)) for a, b_ in zip(hb.From, hb.To_below)]
th = s.table(hb, "tHist", calc={"Students": ('=COUNTIFS(tCV[Exam_Score],">="&[@From],tCV[Exam_Score],"<"&[@To_below])', "int", obs),
                                "Normal_expected": ("=COUNT(tCV[Exam_Score])*(NORM.DIST([@To_below]-0.5,AVERAGE(tCV[Exam_Score]),STDEV.S(tCV[Exam_Score]),TRUE)"
                                                    "-NORM.DIST([@From]-0.5,AVERAGE(tCV[Exam_Score]),STDEV.S(tCV[Exam_Score]),TRUE))", "dec1", expn)})
s.chart("column", [{"name": "Students", "categories": th.ref("Bin"), "values": th.ref("Students"), "gap": 10}], "Exam scores: observed vs normal curve",
        secondary={"type": "line", "name": "Normal expected", "categories": th.ref("Bin"), "values": th.ref("Normal_expected"), "y2_fmt": "#,##0"}, rows=17)
s.insight("Scores are tightly peaked around 67 with a thin right tail up to 100 - the normal curve under-predicts the centre and the top bin. "
          "Use percentiles, not mean ± 2 SD, to set grade boundaries for this cohort.")
s.band("4.2.4", "Pareto: how many SuperMart customers make up 80% of FY2025 sales?")
cust = s25.groupby(["Customer_ID", "Customer_Name"]).Sales.sum().reset_index().sort_values("Sales", ascending=False).reset_index(drop=True)
cust.insert(0, "Rank", np.arange(1, len(cust) + 1))
tot = cust.Sales.sum()
cum = [cust.Sales[cust.Sales >= v].sum() / tot for v in cust.Sales]
s.calc([("Customers who bought in FY2025", "=COUNT(tPareto[Sales])", len(cust), "int", ""),
        ("Customers needed for 80% of sales", '=COUNTIFS(tPareto[Cum_Share],"<0.8")+1', int(sum(c_ < .8 for c_ in cum)) + 1, "int", ""),
        ("…as a share of customers", '=(COUNTIFS(tPareto[Cum_Share],"<0.8")+1)/COUNT(tPareto[Sales])', (sum(c_ < .8 for c_ in cum) + 1) / len(cust), "pct",
         "Not 20%: SuperMart's sales are less concentrated than the '80/20 rule' suggests"),
        ("Top 10 customers' share", "=SUMIFS(tPareto[Sales],tPareto[Rank],\"<=10\")/SUM(tPareto[Sales])", cust.Sales.head(10).sum() / tot, "pct", "Key-account list for the sales team")])
tpa = s.table(cust.head(60), "tPareto_top", fmts={"Sales": "naira", "Rank": "int"}, widths={"Customer_Name": 24})
s.para("Top 60 shown here for the chart; the full 693-customer table tPareto is on sheet D_Pareto_FY25.", "note")
s.chart("column", [{"name": "Sales", "categories": tpa.ref("Customer_Name"), "values": tpa.ref("Sales"), "gap": 20}], "Pareto - top 60 customers FY2025 (full table on D_Pareto_FY25)",
        y_fmt='₦#,##0.0,,"m"', legend=False, rows=17, size=(860, 300))
s.band("4.2.5-7", "Percentile rank, bands and z-scores: ranking the 12 stores on three measures at once")
st = s25.groupby(["Store_ID", "Store_Name"]).agg(Sales=("Sales", "sum"), Profit=("Profit", "sum")).reset_index().merge(stores[["Store_ID", "Annual_Target_2025"]], on="Store_ID")
st["Margin"] = st.Profit / st.Sales
st["Attainment"] = st.Sales / st.Annual_Target_2025
st = st[["Store_ID", "Store_Name", "Sales", "Margin", "Attainment"]]


def prank(arr, x):
    arr = np.sort(np.asarray(arr))
    r = (arr < x).sum()
    return np.floor(r / (len(arr) - 1) * 1000) / 1000


def z(v):
    return ((v - v.mean()) / v.std()).tolist()


pr_s = [prank(st.Sales, x) for x in st.Sales]
band = ["Top quartile" if p >= .75 else ("Upper middle" if p >= .5 else ("Lower middle" if p >= .25 else "Bottom quartile")) for p in pr_s]
zs, zm, za = z(st.Sales), z(st.Margin), z(st.Attainment)
comp = [(a + b_ + c) / 3 for a, b_, c in zip(zs, zm, za)]
crank = pd.Series(comp).rank(ascending=False, method="min").astype(int).tolist()
s.table(st, "tStoreZ", fmts={"Sales": "naira", "Margin": "pct", "Attainment": "pct"}, widths={"Store_Name": 24},
        calc={"Sales_PctRank": ("=PERCENTRANK.INC([Sales],[@Sales])", "pct0", pr_s),
              "Sales_Band": ('=IFS([@Sales_PctRank]>=0.75,"Top quartile",[@Sales_PctRank]>=0.5,"Upper middle",[@Sales_PctRank]>=0.25,"Lower middle",TRUE,"Bottom quartile")', None, band),
              "z_Sales": ("=STANDARDIZE([@Sales],AVERAGE([Sales]),STDEV.S([Sales]))", "dec2", zs),
              "z_Margin": ("=STANDARDIZE([@Margin],AVERAGE([Margin]),STDEV.S([Margin]))", "dec2", zm),
              "z_Attain": ("=STANDARDIZE([@Attainment],AVERAGE([Attainment]),STDEV.S([Attainment]))", "dec2", za),
              "Composite_z": ("=AVERAGE([@z_Sales],[@z_Margin],[@z_Attain])", "dec2", comp),
              "Overall_Rank": ("=RANK.EQ([@Composite_z],[Composite_z],0)", "int", crank)})
best = st.Store_Name[int(np.argmax(comp))]
_last_sales = st.Store_Name[int(np.argmin(st.Sales.values))]
_rk = dict(zip(st.Store_Name, crank))
s.insight(f"{best} ranks first on the composite z-score (equal weight on size, margin and target attainment), but the ranking reshuffles below it: "
          f"{_last_sales} is last on sales yet ranks {_rk[_last_sales]} overall thanks to its margin. "
          "Z-scores put ₦, % and % on one scale - always state the weights you chose.")
s.band("4.2.8", "Detecting anomalies systematically - Lifecare daily bed occupancy, 2023")
zc = ((c23.Occupancy_Pct - c23.Occupancy_Pct.mean()) / c23.Occupancy_Pct.std())
flag = np.where(zc.abs() > 2, "Anomaly", "")
s.calc([("Days in 2023 census", "=ROWS(tCensus23)", len(c23), "int", ""),
        ("Mean occupancy", "=AVERAGE(tCensus23[Occupancy_Pct])", c23.Occupancy_Pct.mean(), "pct", ""),
        ("Days with |z| > 2", '=COUNTIFS(tCensus23[Flag],"Anomaly")', int((zc.abs() > 2).sum()), "int", "Expect ~5% of days if roughly normal"),
        ("Highest occupancy day", "=INDEX(tCensus23[Census_Date],MATCH(MAX(tCensus23[Occupancy_Pct]),tCensus23[Occupancy_Pct],0))",
         serial(c23.Census_Date[c23.Occupancy_Pct.idxmax()]), "date", "Check: staffing roster, an outbreak, or a data-entry error?")])
tc = s.table(c23, "tCensus23", fmts={"Census_Date": "date", "Patients_In_House": "int", "Occupancy_Pct": "pct"},
             calc={"z": ("=STANDARDIZE([@Occupancy_Pct],AVERAGE([Occupancy_Pct]),STDEV.S([Occupancy_Pct]))", "dec2", zc.tolist()),
                   "Flag": ('=IF(ABS([@z])>2,"Anomaly","")', None, flag.tolist())})
s.ws.conditional_format(tc.r0 + 1, tc.ci("Flag"), tc.r0 + len(c23), tc.ci("Flag"),
                        {"type": "cell", "criteria": "==", "value": '"Anomaly"', "format": b.wb.add_format({"bg_color": "#F8D7DA", "font_color": "#7B1F28"})})
s.band("4.2.9", "Choosing the right distribution summary")
s.grid([["Symmetric, no outliers", "Crestview attendance %", "Mean ± SD"],
        ["Skewed money values", "Online order value, hospital bills", "Median + IQR (or P10-P90)"],
        ["Ranking / league table", "12 stores", "Percentile rank or bands"],
        ["Comparing different units", "Sales ₦, margin %, attainment %", "Z-scores (composite)"],
        ["Concentration", "Customers or products", "Pareto curve + 'top N = X%'"],
        ["Spotting unusual days", "Bed occupancy", "Z-score or IQR fences, then investigate"]], header=["Data shape / purpose", "Bootcamp example", "Summary to use"])

# ======================================================================================
# 4.3 Correlation
# ======================================================================================
SH = "4.3 Correlation"
s = b.sheet(SH, "4.3  Correlation Analysis",
            "Pearson, Spearman, a correlation matrix, scatterplots, confounding and a real Simpson's paradox", "4.3.1 - 4.3.10")
s.band("4.3.1-2", "Pearson correlation: what is linked to Crestview exam scores?")
vars_ = ["Hours_Studied_Per_Week", "Attendance_Pct", "Previous_Score", "Tutoring_Sessions", "Sleep_Hours", "Physical_Activity_Hrs"]
s.calc([(f"r( {v} , Exam_Score )", f"=CORREL(tCV[{v}],tCV[Exam_Score])", cvd[v].corr(cvd.Exam_Score), "dec3",
         {"Attendance_Pct": "Strongest link", "Hours_Studied_Per_Week": "Moderate positive", "Sleep_Hours": "≈ 0: no linear link"}.get(v, ""))
        for v in vars_])
s.band("4.3.4-5", "Correlation matrix - CORREL in every cell (Data Analysis > Correlation gives the same numbers, static)")
mv = vars_[:4] + ["Exam_Score"]
md = pd.DataFrame({"Variable": mv, **{v: 0.0 for v in mv}})
fxm = {(i, cj): (f"=CORREL(tCV[{ri}],tCV[{cj}])", float(cvd[ri].corr(cvd[cj]))) for i, ri in enumerate(mv) for cj in mv}
top = s.frame(md, fmts={v: "0.00" for v in mv}, formulas=fxm)
s.ws.conditional_format(top, 2, top + len(mv) - 1, 1 + len(mv), {"type": "3_color_scale", "min_color": "#F8D7DA", "mid_color": "#FFFFFF", "max_color": "#BFE3DA",
                                                                 "min_type": "num", "min_value": -1, "mid_type": "num", "mid_value": 0, "max_type": "num", "max_value": 1})
s.band("4.3.3 / 4.3.6", "Spearman rank correlation and the scatterplot (300 sampled students)")
smp = cvd.sample(300, random_state=7)[["Student_ID", "Hours_Studied_Per_Week", "Exam_Score"]].reset_index(drop=True)
rk_h = smp.Hours_Studied_Per_Week.rank(ascending=False, method="average").tolist()
rk_e = smp.Exam_Score.rank(ascending=False, method="average").tolist()
s.calc([("Pearson r (sample)", "=CORREL(tSpear[Hours_Studied_Per_Week],tSpear[Exam_Score])", smp.Hours_Studied_Per_Week.corr(smp.Exam_Score), "dec3", ""),
        ("Spearman ρ = CORREL of ranks", "=CORREL(tSpear[Rank_Hours],tSpear[Rank_Score])", float(np.corrcoef(rk_h, rk_e)[0, 1]), "dec3",
         "Close to Pearson → relationship is monotonic and roughly linear")])
tsp = s.table(smp, "tSpear", fmts={"Hours_Studied_Per_Week": "int", "Exam_Score": "int"},
              calc={"Rank_Hours": ("=RANK.AVG([@Hours_Studied_Per_Week],[Hours_Studied_Per_Week])", "dec1", rk_h),
                    "Rank_Score": ("=RANK.AVG([@Exam_Score],[Exam_Score])", "dec1", rk_e)})
s.chart("scatter", [{"name": "Students", "categories": tsp.ref("Hours_Studied_Per_Week"), "values": tsp.ref("Exam_Score"),
                     "trendline": {"type": "linear", "display_r_squared": True}}], "Hours studied vs exam score - look before you trust r",
        x_title="Hours studied per week", y_title="Exam score", legend=False, rows=17)
s.band("4.3.7", "Correlation is not causation - Eko Suites cancellations")
s.calc([("r( Lead time , Cancelled )", "=CORREL(tHotel[Lead_Time_Days],--(tHotel[Booking_Status]=\"Cancelled\"))",
         float(np.corrcoef(hod.Lead_Time_Days, hod.Booking_Status == "Cancelled")[0, 1]), "dec3", "Long lead times go with cancellations…"),
        ("r( Special requests , Cancelled )", "=CORREL(tHotel[Special_Requests],--(tHotel[Booking_Status]=\"Cancelled\"))",
         float(np.corrcoef(hod.Special_Requests, hod.Booking_Status == "Cancelled")[0, 1]), "dec3",
         "…but forcing guests to make requests will not stop cancellations: committed guests make requests")])
lb = pd.DataFrame({"Lead_band": ["0-7 days", "8-30 days", "31-90 days", "91-180 days", "180+ days"], "From": [0, 8, 31, 91, 181], "To": [7, 30, 90, 180, 9999]})
cnt = [int(((hod.Lead_Time_Days >= a) & (hod.Lead_Time_Days <= b_)).sum()) for a, b_ in zip(lb.From, lb.To)]
can = [int(((hod.Lead_Time_Days >= a) & (hod.Lead_Time_Days <= b_) & (hod.Booking_Status == "Cancelled")).sum()) for a, b_ in zip(lb.From, lb.To)]
tlb = s.table(lb, "tLeadBand", calc={"Bookings": ('=COUNTIFS(tHotel[Lead_Time_Days],">="&[@From],tHotel[Lead_Time_Days],"<="&[@To])', "int", cnt),
                                     "Cancelled": ('=COUNTIFS(tHotel[Lead_Time_Days],">="&[@From],tHotel[Lead_Time_Days],"<="&[@To],tHotel[Booking_Status],"Cancelled")', "int", can),
                                     "Cancel_rate": ("=[@Cancelled]/[@Bookings]", "pct", [c_ / n for c_, n in zip(can, cnt)])})
s.chart("column", [{"name": "Cancel rate", "categories": tlb.ref("Lead_band"), "values": tlb.ref("Cancel_rate"), "labels": True, "label_fmt": "0%"}],
        "Cancellation rate rises with booking lead time", legend=False, y_fmt="0%", rows=14, size=(600, 260))
s.band("4.3.8", "Confounding: test suspected third variables - seniority (no) and years of experience (yes) in the salary gender gap")
seni = ["Junior", "Mid", "Senior", "Executive"]
g = sal[sal.Gender.isin(["Female", "Male"])]
gt = pd.DataFrame({"Seniority": seni + ["All levels"]})
fem = [g[(g.Seniority == k) & (g.Gender == "Female")].Annual_Salary_NGN.mean() for k in seni] + [g[g.Gender == "Female"].Annual_Salary_NGN.mean()]
mal = [g[(g.Seniority == k) & (g.Gender == "Male")].Annual_Salary_NGN.mean() for k in seni] + [g[g.Gender == "Male"].Annual_Salary_NGN.mean()]
fsh = [(g[(g.Seniority == k)].Gender == "Female").mean() for k in seni] + [(g.Gender == "Female").mean()]
s.table(gt, "tGapSeniority",
        calc={"Female_avg": ('=IF([@Seniority]="All levels",AVERAGEIFS(tSal[Annual_Salary_NGN],tSal[Gender],"Female"),AVERAGEIFS(tSal[Annual_Salary_NGN],tSal[Gender],"Female",tSal[Seniority],[@Seniority]))', "naira", fem),
              "Male_avg": ('=IF([@Seniority]="All levels",AVERAGEIFS(tSal[Annual_Salary_NGN],tSal[Gender],"Male"),AVERAGEIFS(tSal[Annual_Salary_NGN],tSal[Gender],"Male",tSal[Seniority],[@Seniority]))', "naira", mal),
              "Gap_pct": ("=1-[@Female_avg]/[@Male_avg]", "pct", [1 - f_ / m_ for f_, m_ in zip(fem, mal)]),
              "Female_share": ('=IF([@Seniority]="All levels",COUNTIFS(tSal[Gender],"Female")/(COUNTIFS(tSal[Gender],"Female")+COUNTIFS(tSal[Gender],"Male")),'
                               'COUNTIFS(tSal[Gender],"Female",tSal[Seniority],[@Seniority])/(COUNTIFS(tSal[Gender],"Female",tSal[Seniority],[@Seniority])+COUNTIFS(tSal[Gender],"Male",tSal[Seniority],[@Seniority])))', "pct", fsh)})
_n_s = [int(((g.Seniority == k)).sum()) for k in seni]
_adj_s = sum(n_ * (1 - f_ / m_) for n_, f_, m_ in zip(_n_s, fem[:4], mal[:4])) / sum(_n_s)
s.insight(f"Seniority is NOT the confounder here: women and men are spread across levels almost equally, and the gap weighted by level ({_adj_s:.1%}) "
          f"is the same as the raw gap ({1 - fem[4] / mal[4]:.1%}). Testing a suspected confounder and ruling it out is a valid result.")
xb = pd.DataFrame({"Experience_band": ["0-2 yrs", "3-5 yrs", "6-10 yrs", "11-15 yrs", "16+ yrs"], "From": [0, 3, 6, 11, 16], "To": [2, 5, 10, 15, 99]})
def _band(a, b_, gen=None):
    m_ = (g.Years_of_Experience >= a) & (g.Years_of_Experience <= b_)
    return g[m_ & (g.Gender == gen)] if gen else g[m_]
xn = [len(_band(a, b_)) for a, b_ in zip(xb.From, xb.To)]
xf = [_band(a, b_, "Female").Annual_Salary_NGN.mean() for a, b_ in zip(xb.From, xb.To)]
xm = [_band(a, b_, "Male").Annual_Salary_NGN.mean() for a, b_ in zip(xb.From, xb.To)]
xgap = [1 - f_ / m_ for f_, m_ in zip(xf, xm)]
_crit = 'tSal[Years_of_Experience],">="&[@From],tSal[Years_of_Experience],"<="&[@To]'
s.table(xb, "tGapExperience",
        calc={"People": (f'=COUNTIFS(tSal[Gender],"Female",{_crit})+COUNTIFS(tSal[Gender],"Male",{_crit})', "int", xn),
              "Female_avg": (f'=AVERAGEIFS(tSal[Annual_Salary_NGN],tSal[Gender],"Female",{_crit})', "naira", xf),
              "Male_avg": (f'=AVERAGEIFS(tSal[Annual_Salary_NGN],tSal[Gender],"Male",{_crit})', "naira", xm),
              "Gap_pct": ("=1-[@Female_avg]/[@Male_avg]", "pct", xgap)})
_adj_x = sum(n_ * g_ for n_, g_ in zip(xn, xgap)) / sum(xn)
s.calc([("Raw gap (all women vs all men)", "=1-AVERAGEIFS(tSal[Annual_Salary_NGN],tSal[Gender],\"Female\")/AVERAGEIFS(tSal[Annual_Salary_NGN],tSal[Gender],\"Male\")",
             1 - fem[4] / mal[4], "pct", "Women average fewer years of experience than men in this survey"),
            ("Gap with experience held constant", "=SUMPRODUCT(tGapExperience[People],tGapExperience[Gap_pct])/SUM(tGapExperience[People])",
             _adj_x, "pct", "About half the raw gap is explained by experience")])
s.insight(f"Years of experience IS a confounder: it is linked to both gender (women average {g[g.Gender=='Female'].Years_of_Experience.mean():.1f} years, men {g[g.Gender=='Male'].Years_of_Experience.mean():.1f}) "
          f"and salary. Holding it constant cuts the gap from {1 - fem[4] / mal[4]:.1%} to {_adj_x:.1%}. Report both numbers - the remaining gap is still real.")
s.band("4.3.9", "Simpson's paradox in Eko Suites bookings: Deluxe vs Standard Twin")
sp_rows = []
for yr in ["2023", "2024", "All years"]:
    for rt in ["Deluxe", "Standard Twin"]:
        sub = hod[(hod.Room_Type == rt) & ((hod.Arrival_Year == int(yr)) if yr != "All years" else True)]
        sp_rows.append([yr, rt, len(sub), int((sub.Booking_Status == "Cancelled").sum())])
spd = pd.DataFrame(sp_rows, columns=["Arrival_year", "Room_type", "b", "c"])[["Arrival_year", "Room_type"]]
s.table(spd, "tSimpson", widths={"Arrival_year": 12, "Room_type": 16},
        calc={"Bookings": ('=IF([@Arrival_year]="All years",COUNTIFS(tHotel[Room_Type],[@Room_type]),COUNTIFS(tHotel[Room_Type],[@Room_type],tHotel[Arrival_Year],VALUE([@Arrival_year])))', "int", [r[2] for r in sp_rows]),
              "Cancelled": ('=IF([@Arrival_year]="All years",COUNTIFS(tHotel[Room_Type],[@Room_type],tHotel[Booking_Status],"Cancelled"),COUNTIFS(tHotel[Room_Type],[@Room_type],tHotel[Arrival_Year],VALUE([@Arrival_year]),tHotel[Booking_Status],"Cancelled"))', "int", [r[3] for r in sp_rows]),
              "Cancel_rate": ("=[@Cancelled]/[@Bookings]", "pct", [r[3] / r[2] for r in sp_rows])})
s.warn("Deluxe cancels LESS than Standard Twin in 2023 and in 2024, yet MORE overall. Why? 91% of Deluxe bookings fall in 2024 (a high-cancellation year) "
       "versus 79% of Standard Twin. Always check whether a hidden grouping variable is driving an overall comparison.")
s.band("4.3.10", "Reporting correlation responsibly")
s.grid([["Give r, n and the scatter", f"'r = {cvd.Attendance_Pct.corr(cvd.Exam_Score):.2f} between attendance and exam score (n = {len(cvd):,})' + chart"],
        ["Describe strength in words", "|r| < 0.3 weak, 0.3-0.7 moderate, > 0.7 strong - and say 'linear'"],
        ["Say 'associated with', not 'causes'", "'Longer lead times are associated with more cancellations'"],
        ["Name possible confounders", "Seniority in the salary gap; arrival year in the room-type comparison"],
        ["Check subgroups", "Run the same correlation by faculty or year before generalising"]], header=["Rule", "Example from our data"], merge_last=True)

# ======================================================================================
# 4.4 Regression
# ======================================================================================
SH = "4.4 Regression"
s = b.sheet(SH, "4.4  Linear Regression - what drives salary in the ETEH Talent survey?",
            "Simple and multiple regression with LINEST, p-values, residuals, prediction and multicollinearity", "4.4.1 - 4.4.12")
X1 = sal.Years_of_Experience.values
Y = sal.Annual_Salary_NGN.values


def linest(Yv, Xm):
    """Replicates Excel LINEST(y, X, TRUE, TRUE) -> 5 x (k+1) matrix (coefficients in reverse X order, intercept last)."""
    n, k = Xm.shape
    Xd = np.column_stack([Xm, np.ones(n)])
    beta, *_ = np.linalg.lstsq(Xd, Yv, rcond=None)
    resid = Yv - Xd @ beta
    ssr = float(resid @ resid)
    sst = float(((Yv - Yv.mean()) ** 2).sum())
    df = n - k - 1
    sey = np.sqrt(ssr / df)
    cov = sey ** 2 * np.linalg.inv(Xd.T @ Xd)
    se = np.sqrt(np.diag(cov))
    r2 = 1 - ssr / sst
    Fv = ((sst - ssr) / k) / (ssr / df)
    order = list(range(k - 1, -1, -1)) + [k]
    out = [[beta[i] for i in order], [se[i] for i in order], [r2, sey] + [None] * (k - 1), [Fv, df] + [None] * (k - 1), [sst - ssr, ssr] + [None] * (k - 1)]
    return out, beta, se, df, r2


L1, beta1, se1, df1, r2_1 = linest(Y, X1.reshape(-1, 1))
s.band("4.4.2 / 4.4.4", "Simple regression: Salary = a + b × Years of experience")
r1 = s.calc([("Slope b (₦ per extra year)", "=SLOPE(tSal[Annual_Salary_NGN],tSal[Years_of_Experience])", beta1[0], "naira", "Each extra year of experience ≈ this much more pay"),
             ("Intercept a (₦ at 0 years)", "=INTERCEPT(tSal[Annual_Salary_NGN],tSal[Years_of_Experience])", beta1[1], "naira", "Starting salary implied by the line"),
             ("R² (4.4.5)", "=RSQ(tSal[Annual_Salary_NGN],tSal[Years_of_Experience])", r2_1, "pct", "Share of salary variation explained by experience"),
             ("Standard error of estimate", "=STEYX(tSal[Annual_Salary_NGN],tSal[Years_of_Experience])", L1[2][1], "naira", "Typical prediction miss, in ₦")])
s.h2("Full LINEST statistics (dynamic array - spills 5 rows × 2 columns)")
lab_r = s.r
s.grid([["Coefficient"], ["Standard error"], ["R² | SE of estimate"], ["F | degrees of freedom"], ["SS regression | SS residual"]], header=None, c0=1)
s.ws.write(lab_r - 1 if lab_r > 0 else 0, 3, "Years_of_Experience | Intercept", F["note"])
s.spill(lab_r, 3, "=LINEST(tSal[Annual_Salary_NGN],tSal[Years_of_Experience],TRUE,TRUE)", L1, ["#,##0.00"] * 2)
s.band("4.4.3 / 4.4.6-7", "Multiple regression: add Age, Master's and PhD (0/1 dummy columns, Bachelor's/High School = baseline)")
Xm = sal[["Age", "Years_of_Experience", "Is_Masters", "Is_PhD"]].values.astype(float)
Lm, betam, sem, dfm, r2m = linest(Y, Xm)
s.para("LINEST returns coefficients in REVERSE order of the X columns: Is_PhD, Is_Masters, Years_of_Experience, Age, Intercept.", "note")
hdr_r = s.r
for j, nm in enumerate(["Is_PhD", "Is_Masters", "Years_of_Experience", "Age", "Intercept"]):
    s.ws.write(hdr_r, 3 + j, nm, F["hdr"])
s.ws.write(hdr_r, 1, "LINEST row", F["hdr"])
s.ws.write(hdr_r, 2, "", F["hdr"])
s.r += 1
mrow = s.r
for i, nm in enumerate(["Coefficient", "Standard error", "R² | SE", "F | df", "SSreg | SSresid"]):
    s.ws.write(mrow + i, 1, nm, F["lab"])
s.spill(mrow, 3, "=LINEST(tSal[Annual_Salary_NGN],tSal[[Age]:[Is_PhD]],TRUE,TRUE)", Lm, ["#,##0.00"] * 5)
s.r = mrow + 5
pv_cached = [2 * stats.t.sf(abs(Lm[0][j] / Lm[1][j]), dfm) for j in range(5)]
s.ws.write(s.r, 1, "t statistic", F["lab"])
s.ws.write(s.r + 1, 1, "p-value", F["lab"])
for j in range(5):
    cc = xl_rowcol_to_cell(mrow, 3 + j)
    ce = xl_rowcol_to_cell(mrow + 1, 3 + j)
    s.ws.write_formula(s.r, 3 + j, f"={cc}/{ce}", b.nf("dec2", style="result"), Lm[0][j] / Lm[1][j])
    s.ws.write_formula(s.r + 1, 3 + j, f"=T.DIST.2T(ABS({cc}/{ce}),{xl_rowcol_to_cell(mrow + 3, 4, True, True)})", b.nf("0.0000", style="result"), pv_cached[j])
s.r += 3
r2c, dfc = xl_rowcol_to_cell(mrow + 2, 3, True, True), xl_rowcol_to_cell(mrow + 3, 4, True, True)
n_ = len(sal)
adj = 1 - (1 - r2m) * (n_ - 1) / (n_ - 4 - 1)
s.calc([("R² (multiple)", f"={r2c}", r2m, "pct", f"vs {r2_1:.1%} with experience alone"),
        ("Adjusted R² (4.4.6)", f"=1-(1-{r2c})*(COUNT(tSal[Annual_Salary_NGN])-1)/{dfc}", adj, "pct", "Penalises extra predictors; use it to compare models"),
        ("PhD premium (₦, holding age & experience)", f"={xl_rowcol_to_cell(mrow, 3, True, True)}", Lm[0][0], "naira", f"p = {pv_cached[0]:.4f}"),
        ("Master's premium (₦)", f"={xl_rowcol_to_cell(mrow, 4, True, True)}", Lm[0][1], "naira", f"p = {pv_cached[1]:.4f}")])
s.insight("p-values below 0.05 mean the coefficient is unlikely to be zero in the population. Look at the Age coefficient: once experience is in the model, "
          "age adds little or even turns negative - that is multicollinearity (4.4.10), not 'older people earn less'.")
s.band("4.4.9", "Making predictions - type a profile in the yellow cells")
pin = s.inputs([("Years of experience", 6, "gen", "Stay within the data range: 0 - 34 years"), ("Age", 31, "gen", ""), ("Education (Bachelor's / Master's / PhD)", "Master's", "text", "")])
ED = pin["Education (Bachelor's / Master's / PhD)"]
YR, AG = pin["Years of experience"], pin["Age"]
pe = (Lm[0][4] + Lm[0][3] * 31 + Lm[0][2] * 6 + Lm[0][1] * 1 + Lm[0][0] * 0)
cm = [xl_rowcol_to_cell(mrow, 3 + j, True, True) for j in range(5)]
s.calc([("Predicted salary - simple model", f"=INTERCEPT(tSal[Annual_Salary_NGN],tSal[Years_of_Experience])+SLOPE(tSal[Annual_Salary_NGN],tSal[Years_of_Experience])*{YR}",
         beta1[1] + beta1[0] * 6, "naira", "Same as FORECAST.LINEAR(x, known_y, known_x)"),
        ("Predicted salary - multiple model", f'={cm[4]}+{cm[3]}*{AG}+{cm[2]}*{YR}+{cm[1]}*({ED}="Master\'s")+{cm[0]}*({ED}="PhD")',
         pe, "naira", "Uses the LINEST coefficients above"),
        ("Approx. 95% prediction range (±2 × SE)", f'=TEXT({cm[4]}+{cm[3]}*{AG}+{cm[2]}*{YR}+{cm[1]}*({ED}="Master\'s")+{cm[0]}*({ED}="PhD")-2*{xl_rowcol_to_cell(mrow + 2, 4, True, True)},"₦#,##0")&" to "&TEXT({cm[4]}+{cm[3]}*{AG}+{cm[2]}*{YR}+{cm[1]}*({ED}="Master\'s")+{cm[0]}*({ED}="PhD")+2*{xl_rowcol_to_cell(mrow + 2, 4, True, True)},"₦#,##0")',
         f"₦{pe - 2 * Lm[2][1]:,.0f} to ₦{pe + 2 * Lm[2][1]:,.0f}", None, "Individual salaries vary a lot around the line")])
s.band("4.4.8", "Residual analysis - 200 respondents (residual = actual - predicted)")
rs = sal.sample(200, random_state=3)[["Respondent_ID", "Years_of_Experience", "Annual_Salary_NGN"]].reset_index(drop=True)
pred = beta1[1] + beta1[0] * rs.Years_of_Experience
sl_c, in_c = AB(r1["Slope b (₦ per extra year)"]), AB(r1["Intercept a (₦ at 0 years)"])
trs = s.table(rs, "tResid", fmts={"Years_of_Experience": "dec1", "Annual_Salary_NGN": "naira"},
              calc={"Predicted": (f"={in_c}+{sl_c}*[@Years_of_Experience]", "naira", pred.tolist()),
                    "Residual": ("=[@Annual_Salary_NGN]-[@Predicted]", "naira", (rs.Annual_Salary_NGN - pred).tolist())})
s.chart("scatter", [{"name": "Residual", "categories": trs.ref("Years_of_Experience"), "values": trs.ref("Residual")}], "Residuals vs experience - look for curves and funnels",
        x_title="Years of experience", y_title="Residual (₦)", legend=False, y_fmt='₦#,##0.0,,"m"', rows=17)
s.insight("The residuals bend: the line OVER-predicts at both ends (mostly negative residuals for 0-2 years and for 20+ years) and UNDER-predicts in the middle (6-15 years). "
          "Pay rises fast early and then levels off - a curved (log or polynomial) term would fit better. Residual plots reveal what R² hides.")
s.band("4.4.10-12", "Multicollinearity, limitations and honest reporting")
rag = sal.Age.corr(sal.Years_of_Experience)
s.calc([("r( Age , Years of experience )", "=CORREL(tSal[Age],tSal[Years_of_Experience])", rag, "dec3", "Above 0.9 → the two variables carry almost the same information"),
        ("Variance inflation factor (VIF)", "=1/(1-CORREL(tSal[Age],tSal[Years_of_Experience])^2)", 1 / (1 - rag ** 2), "dec1", "VIF > 5 is a warning, > 10 serious - drop one or combine them")])
s.grid([["Range", "Do not predict for 40 years' experience - the data stops at 34"],
        ["Survey sample", "Self-reported salaries of 1,781 respondents - not every Nigerian professional"],
        ["Missing drivers", "Company size, city and sector are not in the data"],
        ["Report", f"'Each additional year of experience is associated with about ₦{beta1[0]:,.0f} higher annual salary (R² = {r2_1:.2f}, n = {n_:,}).'"]],
       header=["Limitation", "In this example"], merge_last=True)

# ======================================================================================
# 4.5 Time series and forecasting
# ======================================================================================
SH = "4.5 Forecasting"
s = b.sheet(SH, "4.5  Time Series and Forecasting - SuperMart monthly sales 2022-2025",
            "Moving averages, smoothing, seasonal indices, FORECAST.ETS with confidence bands and a 2025 backtest", "4.5.1 - 4.5.13")
mser = sm.groupby(sm.Order_Date.dt.to_period("M")).Sales.sum()
months = [p.to_timestamp() for p in mser.index]
yv = mser.values.astype(float)
N = len(yv)


def hw_fit(y_, m=12):
    def run(params, y_, h=0):
        a, b_, g_ = params
        lvl = y_[:m].mean()
        tr = (y_[m:2 * m].mean() - y_[:m].mean()) / m
        se_ = list(y_[:m] - lvl)
        fitted, err = [], []
        for t in range(len(y_)):
            f_ = lvl + tr + se_[t]
            fitted.append(f_)
            e = y_[t] - f_
            err.append(e)
            nl = a * (y_[t] - se_[t]) + (1 - a) * (lvl + tr)
            tr = b_ * (nl - lvl) + (1 - b_) * tr
            se_.append(g_ * (y_[t] - nl) + (1 - g_) * se_[t])
            lvl = nl
        fc = [lvl + (k + 1) * tr + se_[len(y_) - m + (k % m)] for k in range(h)]
        return np.array(err), fc

    res = optimize.minimize(lambda p: (run(p, y_)[0] ** 2).sum(), x0=[.3, .05, .2], bounds=[(0, 1)] * 3, method="L-BFGS-B")
    a, b_, g_ = res.x
    err, fc = run(res.x, y_, 12)
    sig = np.sqrt((err[m:] ** 2).mean())
    ci = []
    for h in range(1, 13):
        v = 1 + sum((a + j * a * b_ + (g_ if j % m == 0 else 0)) ** 2 for j in range(1, h))
        ci.append(1.96 * sig * np.sqrt(v))
    return np.array(fc), np.array(ci), res.x


alpha_default = 0.3
s.band("4.5.1-3", "The series: trend + strong Q4 seasonality")
alpha_in = s.inputs([("Smoothing constant α (4.5.5)", alpha_default, "dec2", "0 = very smooth, 1 = follows the last month; try 0.1 and 0.6")])["Smoothing constant α (4.5.5)"]
tsd = pd.DataFrame({"Month": months, "t": np.arange(1, N + 1), "Month_No": [m.month for m in months], "Sales": yv})
tsd["MA3"] = np.nan
tsd["MA12"] = np.nan
tsd["SES_forecast"] = np.nan
top = s.r + 1  # first data row (0-based) after header
ses = [yv[0]]
for i in range(1, N):
    ses.append(alpha_default * yv[i - 1] + (1 - alpha_default) * ses[-1])
fx = {}
for i in range(N):
    r_ = top + i + 1  # excel row number
    if i >= 2:
        fx[(i, "MA3")] = (f"=AVERAGE(E{r_ - 2}:E{r_})", float(yv[i - 2:i + 1].mean()))
    if i >= 11:
        fx[(i, "MA12")] = (f"=AVERAGE(E{r_ - 11}:E{r_})", float(yv[i - 11:i + 1].mean()))
    fx[(i, "SES_forecast")] = ((f"=E{r_}" if i == 0 else f"={alpha_in}*E{r_ - 1}+(1-{alpha_in})*H{r_ - 1}"), float(ses[i]))
s.frame(tsd, fmts={"Month": "mon", "t": "int", "Month_No": "int", "Sales": "naira", "MA3": "naira", "MA12": "naira", "SES_forecast": "naira"}, formulas=fx)
for cc_, w in [(1, 11), (2, 5), (3, 9), (4, 14), (5, 14), (6, 14), (7, 14)]:
    s.ws.set_column(cc_, cc_, w)
rng = lambda col, a=1, z=N: f"${col}${top + a}:${col}${top + z}"  # noqa: E731
s.chart("line", [{"name": "Sales", "categories": f"='{SH}'!{rng('B')}", "values": f"='{SH}'!{rng('E')}"},
                 {"name": "12-month moving average (trend)", "categories": f"='{SH}'!{rng('B')}", "values": f"='{SH}'!{rng('G')}"}],
        "Monthly sales with 12-month moving average: trend rising, November/December peaks", y_fmt='₦#,##0,,"m"', rows=17,
        extra=lambda ch: ch.set_x_axis({"date_axis": True, "num_format": "mmm yy", "major_unit": 6, "major_unit_type": "months"}))
s.band("4.5.3", "Seasonal index by month (training years 2022-2024)")
tr_mask = np.arange(N) < 36
si = [yv[tr_mask & (tsd.Month_No.values == m)].mean() / yv[tr_mask].mean() for m in range(1, 13)]
sid = pd.DataFrame({"Month_No": range(1, 13), "Month": MONTHS})
tsi = s.table(sid, "tSeasonIdx", calc={"Seasonal_index": (f'=AVERAGEIFS({rng("E")},{rng("D")},[@Month_No],{rng("C")},"<=36")/AVERAGE({rng("E", 1, 36)})', "dec2", si)})
s.chart("column", [{"name": "Index", "categories": tsi.ref("Month"), "values": tsi.ref("Seasonal_index"), "labels": True, "label_fmt": "0.00"}],
        "Seasonal index (1.00 = average month)", legend=False, rows=14, size=(600, 260))
s.band("4.5.6-10", "Backtest: train on 2022-2024, forecast 2025, compare with what actually happened")
fc_bt, _, par_bt = hw_fit(yv[:36])
sl, ic = np.polyfit(np.arange(1, 37), yv[:36], 1)
bt = []
for k in range(12):
    i = 36 + k
    trend = ic + sl * (i + 1)
    bt.append([months[i], yv[i], yv[i - 12], trend * si[k], float(fc_bt[k])])
btd = pd.DataFrame(bt, columns=["Month", "Actual", "Seasonal_naive", "Trend_x_Season", "ETS"])
fxb = {}
for k in range(12):
    i = 36 + k
    r_ = top + i + 1
    fxb[(k, "Actual")] = (f"=E{r_}", float(yv[i]))
    fxb[(k, "Seasonal_naive")] = (f"=E{r_ - 12}", float(yv[i - 12]))
    fxb[(k, "Trend_x_Season")] = (f"=FORECAST.LINEAR(C{r_},{rng('E', 1, 36)},{rng('C', 1, 36)})*INDEX(tSeasonIdx[Seasonal_index],D{r_})", bt[k][3])
    fxb[(k, "ETS")] = (f"=FORECAST.ETS(B{r_},{rng('E', 1, 36)},{rng('B', 1, 36)},12,1)", bt[k][4])
s.para("ETS column: Excel's AAA exponential smoothing. Cached numbers were produced by an equivalent Holt-Winters fit outside Excel; "
       "Excel recalculates on open and may differ slightly.", "note")
btop = s.frame(btd, fmts={"Month": "mon", "Actual": "naira", "Seasonal_naive": "naira", "Trend_x_Season": "naira", "ETS": "naira"}, formulas=fxb)
act = btd.Actual.values
acc = []
for meth, col in [("Seasonal naive", "D"), ("Trend × season", "E"), ("FORECAST.ETS", "F")]:
    fcv = btd[{"D": "Seasonal_naive", "E": "Trend_x_Season", "F": "ETS"}[col]].values
    ra = f"$C${btop + 1}:$C${btop + 12}"
    rf = f"${col}${btop + 1}:${col}${btop + 12}"
    acc.append((meth, ra, rf, np.mean(np.abs(act - fcv)), np.mean(np.abs(act - fcv) / act), np.sum(fcv - act) / np.sum(act)))
s.h2("4.5.9  Forecast accuracy, 2025")
s.calc([(f"MAPE - {m}", f"=SUMPRODUCT(ABS({ra}-{rf})/{ra})/12", mape, "pct", "Average % miss per month") for m, ra, rf, mae, mape, bias in acc] +
       [(f"Bias - {m}", f"=SUM({rf})/SUM({ra})-1", bias, "pct", "+ over-forecast, − under-forecast") for m, ra, rf, mae, mape, bias in acc])
best_m = min(acc, key=lambda a_: a_[4])[0]
s.insight(f"Lowest MAPE in the 2025 backtest: {best_m}. Monthly retail sales are noisy - even the best method misses individual months by a wide margin, "
          "so plan with ranges, not single numbers.")
s.band("4.5.6-8 / 4.5.11", "2026 forecast with FORECAST.ETS, 95% confidence band and detected seasonality")
fc26, ci26, par26 = hw_fit(yv)
m26 = pd.date_range("2026-01-01", periods=12, freq="MS")
s.calc([("Detected season length", f"=FORECAST.ETS.SEASONALITY({rng('E')},{rng('B')})", 12, "int", "12 = yearly cycle in monthly data"),
        ("Forecast FY2026 total", f"=SUM(tForecast26[Forecast])", float(fc26.sum()), "naira", f"vs FY2025 actual ₦{yv[36:].sum():,.0f}")])
f26 = pd.DataFrame({"Month": m26})
tf = s.table(f26, "tForecast26", fmts={"Month": "mon"},
             calc={"Forecast": (f"=FORECAST.ETS([@Month],{rng('E')},{rng('B')},12,1)", "naira", fc26.tolist()),
                   "CI_95": (f"=FORECAST.ETS.CONFINT([@Month],{rng('E')},{rng('B')},0.95,12,1)", "naira", ci26.tolist()),
                   "Lower": ("=[@Forecast]-[@CI_95]", "naira", (fc26 - ci26).tolist()),
                   "Upper": ("=[@Forecast]+[@CI_95]", "naira", (fc26 + ci26).tolist())})
s.chart("line", [{"name": "Forecast", "categories": tf.ref("Month"), "values": tf.ref("Forecast")},
                 {"name": "Lower 95%", "categories": tf.ref("Month"), "values": tf.ref("Lower"), "color": "#9DB4C0"},
                 {"name": "Upper 95%", "categories": tf.ref("Month"), "values": tf.ref("Upper"), "color": "#9DB4C0"}],
        "FY2026 monthly forecast with 95% band", y_fmt='₦#,##0,,"m"', rows=17)
s.para("Data > Forecast Sheet draws the same chart automatically: select tsd Month + Sales, set end date Dec 2026, confidence 95%, seasonality 12.", "note")
s.band("4.5.12-13", "Limitations and choosing a method")
s.grid([["Seasonal naive", "Very short history, quick sanity check", "Ignores growth"],
        ["Moving average", "Showing the underlying trend", "Lags turning points; no seasonality"],
        ["Trend × seasonal index", "Explainable to managers; works in any Excel", "Assumes straight-line growth"],
        ["FORECAST.ETS", "≥ 2-3 full seasons of regular monthly data", "Black box; needs evenly spaced dates"],
        ["Any method", "Store openings, price changes, FX shocks are not in the history", "Add judgement and scenarios (4.6)"]], header=["Method", "Use when", "Watch out for"])

# ======================================================================================
# 4.6 What-if
# ======================================================================================
SH = "4.6 What-If & Solver"
s = b.sheet(SH, "4.6  Goal Seek, Data Tables, Scenarios and Solver - Eko Suites revenue plan",
            "A small, transparent model built from 2024 bookings, then every what-if tool applied to it", "4.6.1 - 4.6.12")
h24 = hod[(hod.Arrival_Year == 2024) & (hod.Booking_Status == "Honoured")]
rn24 = int(h24.Total_Nights.sum())
rev24 = float(h24.Booking_Value_NGN.sum())
s.band("4.6.1", "The model - inputs (yellow) from the 2024 data, outputs (green) are formulas")
s.scenario(f"Eko Suites, Victoria Island: modelled as a {ROOMS}-room hotel (peak nights in the data reach 307 rooms in use). "
           "2024 honoured stays give the base occupancy and average daily rate (ADR).")
base = s.calc([("2024 room-nights sold", '=SUMIFS(tHotel[Total_Nights],tHotel[Arrival_Year],2024,tHotel[Booking_Status],"Honoured")', rn24, "int", "Honoured bookings only"),
               ("2024 room revenue", '=SUMIFS(tHotel[Booking_Value_NGN],tHotel[Arrival_Year],2024,tHotel[Booking_Status],"Honoured")', rev24, "naira", ""),
               ("2024 ADR", '=SUMIFS(tHotel[Booking_Value_NGN],tHotel[Arrival_Year],2024,tHotel[Booking_Status],"Honoured")/SUMIFS(tHotel[Total_Nights],tHotel[Arrival_Year],2024,tHotel[Booking_Status],"Honoured")',
                rev24 / rn24, "naira", "Average daily rate per occupied room"),
               ("2024 occupancy", f'=SUMIFS(tHotel[Total_Nights],tHotel[Arrival_Year],2024,tHotel[Booking_Status],"Honoured")/({ROOMS}*366)', rn24 / (ROOMS * 366), "pct", "2024 had 366 days")])
occ0, adr0 = round(rn24 / (ROOMS * 366), 3), round(rev24 / rn24, -2)
cost_per = 28000.0
fixed = 3_600_000_000.0
inp = s.inputs([("Rooms", ROOMS, "int", "Model assumption"), ("Occupancy 2025", occ0, "pct", "Start from 2024 actual"),
                ("ADR 2025 (₦)", adr0, "naira", "Start from 2024 actual, rounded"), ("Variable cost per occupied room-night (₦)", cost_per, "naira", "Housekeeping, laundry, amenities, utilities"),
                ("Fixed costs per year (₦)", fixed, "naira", "Salaries, lease, maintenance, marketing, F&B overheads (assumption)")])
R_, O_, P_, V_, X_ = inp["Rooms"], inp["Occupancy 2025"], inp["ADR 2025 (₦)"], inp["Variable cost per occupied room-night (₦)"], inp["Fixed costs per year (₦)"]
rev0 = ROOMS * 365 * occ0 * adr0
prof0 = ROOMS * 365 * occ0 * (adr0 - cost_per) - fixed
outs = s.calc([("Room-nights sold", f"={R_}*365*{O_}", ROOMS * 365 * occ0, "int", ""),
               ("Room revenue", f"={R_}*365*{O_}*{P_}", rev0, "naira", ""),
               ("Operating profit", f"={R_}*365*{O_}*({P_}-{V_})-{X_}", prof0, "naira", "The cell Goal Seek and Solver will target"),
               ("RevPAR (revenue per available room-night)", f"={O_}*{P_}", occ0 * adr0, "naira", "Hotel industry headline KPI")])
PROFIT = AB(outs["Operating profit"])
s.band("4.6.2-3", "Goal Seek: what ADR gives ₦1bn operating profit?")
tgt = 1_000_000_000.0
adr_needed = (tgt + fixed) / (ROOMS * 365 * occ0) + cost_per
s.grid([["1", f"Data > What-If Analysis > Goal Seek. Set cell: {PROFIT}   To value: 1000000000   By changing: {P_}"],
        ["2", "Excel iterates until profit = ₦1bn and writes the answer into the ADR input"],
        ["3", "Check the answer algebraically (below) - Goal Seek finds ONE input for ONE target; it cannot respect constraints"]], header=["Step", "Action"], merge_last=True)
s.calc([("ADR needed (algebra check)", f"=(1000000000+{X_})/({R_}*365*{O_})+{V_}", adr_needed, "naira", "Goal Seek should land on this number"),
        ("Increase vs current ADR input", f"=(1000000000+{X_})/({R_}*365*{O_})+{V_}-{P_}", adr_needed - adr0, "naira", "Is a rise this large realistic against OTA competitors?")])
s.band("4.6.4-6", "Data tables / sensitivity: profit for every occupancy × ADR combination")
s.para("Excel's Data Table ({=TABLE()}) cannot be created from a script, so each cell below computes the model directly - identical results. "
       "To build the real one: put the profit formula in the corner, select the grid, Data > What-If > Data Table, Row input = ADR cell, Column input = Occupancy cell.", "note")
occs = [0.40, 0.45, 0.50, 0.55, 0.60, 0.65, 0.70]
adrs = [80000, 90000, 100000, 110000, 120000, 130000]
dt = pd.DataFrame({"Occupancy ↓ / ADR →": occs, **{f"₦{a // 1000}k": 0.0 for a in adrs}})
fxd = {}
r_top = s.r + 1
for i, oc in enumerate(occs):
    for j, a in enumerate(adrs):
        fxd[(i, f"₦{a // 1000}k")] = (f"={R_}*365*B{r_top + i + 1}*({a}-{V_})-{X_}", ROOMS * 365 * oc * (a - cost_per) - fixed)
t0 = s.frame(dt, fmts={"Occupancy ↓ / ADR →": "pct0", **{f"₦{a // 1000}k": '₦#,##0,,"m"' for a in adrs}}, formulas=fxd)
s.ws.conditional_format(t0, 2, t0 + len(occs) - 1, 1 + len(adrs), {"type": "3_color_scale", "min_color": "#F8D7DA", "mid_color": "#FFFFFF", "max_color": "#BFE3DA",
                                                                   "mid_type": "num", "mid_value": 0})
s.insight("Reading the table (4.6.6): the red cells are losses. Five points of occupancy is worth roughly as much as a ₦10k ADR rise - "
          "so a campaign that fills rooms can matter as much as a price rise.")
s.band("4.6.7-8", "Scenarios: pick one in the yellow cell (Scenario Manager stores the same three sets of inputs)")
scen = pd.DataFrame({"Scenario": ["Pessimistic", "Base", "Optimistic"], "Occupancy": [occ0 - .06, occ0, occ0 + .06], "ADR": [adr0 * .92, adr0, adr0 * 1.08],
                     "Fixed_costs": [fixed * 1.05, fixed, fixed]})
sc_prof = (ROOMS * 365 * scen.Occupancy * (scen.ADR - cost_per) - scen.Fixed_costs).tolist()
sc_rev = (ROOMS * 365 * scen.Occupancy * scen.ADR).tolist()
s.table(scen, "tScen", fmts={"Occupancy": "pct", "ADR": "naira", "Fixed_costs": "naira"},
        calc={"Revenue": (f"={R_}*365*[@Occupancy]*[@ADR]", "naira", sc_rev),
              "Profit": (f"={R_}*365*[@Occupancy]*([@ADR]-{V_})-[@Fixed_costs]", "naira", sc_prof)})
pick = s.inputs([("Selected scenario", "Base", "text", "Pessimistic / Base / Optimistic")])["Selected scenario"]
s.ws.data_validation(pick.replace("$", ""), {"validate": "list", "source": ["Pessimistic", "Base", "Optimistic"]})
s.calc([("Profit - selected scenario", f"=INDEX(tScen[Profit],MATCH({pick},tScen[Scenario],0))", sc_prof[1], "naira", "Scenario summary report = this table"),
        ("Spread (optimistic - pessimistic)", "=MAX(tScen[Profit])-MIN(tScen[Profit])", max(sc_prof) - min(sc_prof), "naira", "The range management must plan for")])
s.band("4.6.9-11", "Solver: how should rooms be allocated across market segments each night?")
seg = ho[ho.Arrival_Year == 2024].groupby("Market_Segment").agg(ADR=("Avg_Room_Rate_NGN", "mean"), Cancel_rate=("Booking_Status", lambda v: (v == "Cancelled").mean())).reset_index()
seg = seg[seg.Market_Segment != "Complimentary"].reset_index(drop=True)
comm = {"Online Travel Agent": 0.18, "Walk-in / Travel Agent": 0.08, "Corporate": 0.0, "Airline Crew": 0.0}
mins = {"Online Travel Agent": 0, "Walk-in / Travel Agent": 0, "Corporate": 40, "Airline Crew": 10}
caps = {"Online Travel Agent": 220, "Walk-in / Travel Agent": 110, "Corporate": 70, "Airline Crew": 10}
seg["Commission"] = seg.Market_Segment.map(comm)
seg["Min_rooms"] = seg.Market_Segment.map(mins)
seg["Max_demand"] = seg.Market_Segment.map(caps)
val = (seg.ADR * (1 - seg.Cancel_rate) * (1 - seg.Commission)).values
lp = optimize.linprog(-val, A_ub=[np.ones(len(seg)), (seg.Market_Segment == "Online Travel Agent").astype(float).values],
                      b_ub=[ROOMS, 0.5 * ROOMS], bounds=list(zip(seg.Min_rooms, seg.Max_demand)), method="highs")
sol = np.round(lp.x).astype(int)
seg["Rooms_allocated"] = sol
segd = seg[["Market_Segment", "ADR", "Cancel_rate", "Commission", "Min_rooms", "Max_demand", "Rooms_allocated"]]
tsg = s.table(segd, "tSegments", fmts={"ADR": "naira", "Cancel_rate": "pct", "Commission": "pct0", "Min_rooms": "int", "Max_demand": "int", "Rooms_allocated": "int"},
              widths={"Market_Segment": 22},
              calc={"Net_value_per_room": ("=[@ADR]*(1-[@Cancel_rate])*(1-[@Commission])", "naira", val.tolist()),
                    "Nightly_value": ("=[@Rooms_allocated]*[@Net_value_per_room]", "naira", (sol * val).tolist()),
                    "Within_limits": ("=AND([@Rooms_allocated]>=[@Min_rooms],[@Rooms_allocated]<=[@Max_demand])", None, [True] * len(seg))},
              total={"Market_Segment": "Total", "Rooms_allocated": "sum", "Nightly_value": "sum"})
s.ws.conditional_format(tsg.r0 + 1, tsg.ci("Rooms_allocated"), tsg.r0 + len(segd), tsg.ci("Rooms_allocated"),
                        {"type": "no_errors", "format": b.wb.add_format({"bg_color": "#FFF4CC", "border": 1, "border_color": "#E0A526"})})
s.calc([("Objective: expected net value per night", "=SUM(tSegments[Nightly_value])", float((sol * val).sum()), "naira", "Solver: Set Objective → Max"),
        ("Constraint: total rooms ≤ capacity", f"=SUM(tSegments[Rooms_allocated])<={R_}", bool(sol.sum() <= ROOMS), None, f"Uses {int(sol.sum())} of {ROOMS}"),
        ("Constraint: OTA ≤ 50% of rooms", f'=SUMIFS(tSegments[Rooms_allocated],tSegments[Market_Segment],"Online Travel Agent")<=0.5*{R_}',
         bool(sol[seg.Market_Segment == "Online Travel Agent"].sum() <= .5 * ROOMS), None, "Rate-parity / dependency policy")])
s.grid([["Objective", "Maximise SUM(tSegments[Nightly_value])"],
        ["By changing", "tSegments[Rooms_allocated] (yellow column)"],
        ["Subject to", "Rooms_allocated ≥ Min_rooms; Rooms_allocated ≤ Max_demand; integer; SUM ≤ Rooms; OTA ≤ 50% of Rooms"],
        ["Method", "Simplex LP (the model is linear) - Solver's answer is already filled in; reset the column to 0 and re-run to see it"]],
       header=["Solver box", "Entry"], merge_last=True)
s.insight("OTA bookings have the highest headline ADR but the LOWEST net value per room once 36% cancellations and 18% commission are removed. "
          "Solver fills corporate and walk-in demand first - a result managers rarely guess.")
s.band("4.6.12", "Documenting what-if analysis")
s.grid([["Inputs", "Yellow cells only; source and date beside each (2024 bookings, extracted from D_Hotel)"],
        ["Assumptions", f"{ROOMS} rooms, ₦{cost_per:,.0f} variable cost per room-night, ₦{fixed/1e9:.2f}bn fixed costs, commission rates, demand caps"],
        ["Tools used", "Goal Seek (ADR for ₦1bn), sensitivity table, 3 scenarios, Solver LP allocation"],
        ["Limits", "Annual averages; ignores seasonality (see 4.5) and group bookings"]], header=["Item", "Recorded"], merge_last=True)

# ======================================================================================
# 4.7 Storytelling
# ======================================================================================
SH = "4.7 Data Storytelling"
s = b.sheet(SH, "4.7  Data Storytelling - SuperMart FY2025 in one story",
            "Headline built from live numbers, three-act structure, finding-titles, annotation and misleading vs honest axes", "4.7.1 - 4.7.9")
g25 = s25.Sales.sum() / s24.Sales.sum() - 1
stf = s25.groupby(["Store_ID", "Store_Name", "Zone"]).Sales.sum().reset_index().merge(stores[["Store_ID", "Annual_Target_2025"]], on="Store_ID")
stf["Attainment"] = stf.Sales / stf.Annual_Target_2025
missed = int((stf.Attainment < 1).sum())
s.band("4.7.3", "The headline message - written by a formula, so it can never disagree with the numbers")
tst = s.table(stf.drop(columns=["Attainment"]), "tStory", fmts={"Sales": "naira", "Annual_Target_2025": "naira"}, widths={"Store_Name": 24},
              calc={"Attainment": ("=[@Sales]/[@Annual_Target_2025]", "pct", stf.Attainment.tolist()),
                    "Gap_to_target": ("=[@Sales]-[@Annual_Target_2025]", "naira", (stf.Sales - stf.Annual_Target_2025).tolist())})
worst = stf.sort_values("Sales").assign(g=lambda d: d.Sales - d.Annual_Target_2025).sort_values("g").iloc[0]
head = (f"FY2025 sales grew {g25:.0%} to ₦{s25.Sales.sum() / 1e6:,.0f}m, but {missed} of 12 stores missed target - "
        f"{worst.Store_Name} has the largest gap (₦{-worst.g / 1e6:,.1f}m)")
s.calc([("Headline", f'="FY2025 sales grew "&TEXT(SUM(tStory[Sales])/{s24.Sales.sum():.0f}-1,"0%")&" to ₦"&TEXT(SUM(tStory[Sales])/1000000,"#,##0")&"m, but "&COUNTIFS(tStory[Attainment],"<1")&" of 12 stores missed target - "'
         f'&INDEX(tStory[Store_Name],MATCH(MIN(tStory[Gap_to_target]),tStory[Gap_to_target],0))&" has the largest gap (₦"&TEXT(-MIN(tStory[Gap_to_target])/1000000,"0.0")&"m)"',
         head, None, "FY2024 sales (₦609.2m) typed as the comparison base")])
s.band("4.7.2 / 4.7.4-5", "Three acts, message hierarchy and evidence by audience")
s.grid([["Act 1 - Situation", f"SuperMart grew {g25:.0%} in FY2025; every zone grew", "Line: monthly sales 2022-2025"],
        ["Act 2 - Complication", f"{missed} of 12 stores missed target; the North-Central and South-East stores trail", "Bar: attainment by store, target line at 100%"],
        ["Act 3 - Resolution", "Re-phase targets for new stores; replicate Ikeja/Lekki promotions in Wuse and Owerri", "Table: 3 actions, owner, date"]],
       header=["Act", "Message", "Evidence"])
s.grid([["Board", "Headline + 3 numbers + decision needed", "1 slide"], ["Regional managers", "Their stores vs target, monthly trend", "Dashboard (Week 3, 3.8)"],
        ["Store managers", "Line-level detail, loss-making products", "Filtered table"]], header=["Audience", "Evidence they need", "Format"])
s.band("4.7.6-7", "Titles that state the finding + annotation")
st_sorted = stf.sort_values("Attainment").reset_index(drop=True)
tat = s.table(st_sorted[["Store_Name", "Attainment"]].rename(columns={"Attainment": "Attainment_FY25"}), "tAttain", fmts={"Attainment_FY25": "pct"}, widths={"Store_Name": 24},
              calc={"Target": ("=1", "pct0", [1] * 12)})
pts = [{"fill": {"color": "#B23A48" if a < 1 else "#13807A"}} for a in st_sorted.Attainment]
s.chart("bar", [{"name": "Attainment", "categories": tat.ref("Store_Name"), "values": tat.ref("Attainment_FY25"), "labels": True, "label_fmt": "0%", "points": pts}],
        f"{missed} of 12 stores missed their FY2025 target (red = below 100%)", legend=False, y_fmt="0%", rows=18, size=(720, 360))
s.para("Weak title: 'Store attainment FY2025'. Strong title: the sentence above - the reader gets the finding before reading the chart.", "note")
s.band("4.7.8", "Avoiding misleading presentations - same numbers, two y-axes")
zt = s25.groupby("Zone").Sales.sum().reset_index().sort_values("Sales", ascending=False)
tz = s.table(zt, "tZoneStory", fmts={"Sales": "naira"})
mn = zt.Sales.min()
s.chart("column", [{"name": "Sales", "categories": tz.ref("Zone"), "values": tz.ref("Sales")}], "MISLEADING: axis starts at ₦140m - South-West looks 4x larger",
        y_fmt='₦#,##0,,"m"', legend=False, rows=15, size=(520, 260), extra=lambda ch: ch.set_y_axis({"min": 140e6, "num_format": '₦#,##0,,"m"'}))
s.chart("column", [{"name": "Sales", "categories": tz.ref("Zone"), "values": tz.ref("Sales")}], "HONEST: zero baseline - South-West is about 1.7x the next zone",
        y_fmt='₦#,##0,,"m"', legend=False, rows=15, size=(520, 260), extra=lambda ch: ch.set_y_axis({"min": 0, "num_format": '₦#,##0,,"m"'}))
s.band("4.7.9", "Structuring the written narrative")
s.grid([["1. Headline", head], ["2. Why it matters", "Missed stores represent ₦" + f"{-stf[stf.Attainment < 1].eval('Sales-Annual_Target_2025').sum() / 1e6:,.1f}" + "m of unmet target"],
        ["3. Evidence", "Attainment chart; zone table; monthly trend"], ["4. So what / recommendation", "Three actions with owners and dates"],
        ["5. Caveats", "Targets were set before two store refurbishments; FY2025 figures are unaudited"]], header=["Paragraph", "Content"], merge_last=True)

# ======================================================================================
# 4.8 Report writing
# ======================================================================================
SH = "4.8 Report Writing"
s = b.sheet(SH, "4.8  Report Writing and PDF Production - Lifecare 2023 Operations Report",
            "An executive summary written from live numbers; this sheet is set up to print as a clean A4 PDF", "4.8.1 - 4.8.10")
adm23 = len(los)
alos23 = los.Length_of_Stay.mean()
rr23 = (los.Readmit_30d == "Yes").mean()
occ23 = c23.Occupancy_Pct.mean()
d95 = int((c23.Occupancy_Pct > .95).sum())
top_dep = los.groupby("Department").size().idxmax()
s.band("4.8.1", "Report structure")
s.grid([["1", "Title page", "Lifecare General Hospital - 2023 Operations Report; version, date, author, classification"],
        ["2", "Executive summary", "One page: findings, implications, recommendations (below)"],
        ["3", "Method & data", "Source extracts, period, exclusions, definitions (ALOS, readmission)"],
        ["4", "Findings", "Admissions, bed occupancy, length of stay, readmissions, billing"],
        ["5", "Recommendations", "Numbered, owned, dated"],
        ["6", "Appendix", "Full tables, data dictionary, reconciliation"]], header=["#", "Section", "Lifecare 2023 content"])
s.band("4.8.2", "Executive summary - every number is a formula over tLOS / tCensus23")
s.calc([("Finding 1", f'="Lifecare admitted "&TEXT(ROWS(tLOS),"#,##0")&" patients in 2023; "&INDEX(tLOSdep[Department],MATCH(MAX(tLOSdep[Admissions]),tLOSdep[Admissions],0))&" was the busiest department."',
         f"Lifecare admitted {adm23:,} patients in 2023; {top_dep} was the busiest department.", None, ""),
        ("Finding 2", f'="Average bed occupancy was "&TEXT(AVERAGE(tCensus23[Occupancy_Pct]),"0.0%")&", above the 85% safe level, with "&COUNTIFS(tCensus23[Occupancy_Pct],">0.95")&" days above 95%."',
         f"Average bed occupancy was {occ23:.1%}, above the 85% safe level, with {d95} days above 95%.", None, ""),
        ("Finding 3", f'="Average length of stay was "&TEXT(AVERAGE(tLOS[Length_of_Stay]),"0.0")&" days and the 30-day readmission rate "&TEXT(COUNTIFS(tLOS[Readmit_30d],"Yes")/ROWS(tLOS),"0.0%")&"."',
         f"Average length of stay was {alos23:.1f} days and the 30-day readmission rate {rr23:.1%}.", None, ""),
        ("Recommendation", "Open 20 surge beds for Q4 and pilot a discharge-planning nurse in the highest-readmission department.",
         "Open 20 surge beds for Q4 and pilot a discharge-planning nurse in the highest-readmission department.", None, "Recommendations are typed - judgement, not formula")])
s.band("4.8.4", "Table design for reports - department summary")
dep = los.groupby("Department").agg(Admissions=("Admission_ID", "count")).reset_index().sort_values("Admissions", ascending=False)
alos_d = [los[los.Department == d].Length_of_Stay.mean() for d in dep.Department]
rr_d = [(los[los.Department == d].Readmit_30d == "Yes").mean() for d in dep.Department]
bill_d = [los[los.Department == d].Billing_Amount.mean() for d in dep.Department]
s.table(dep, "tLOSdep", fmts={"Admissions": "int"}, widths={"Department": 26},
        calc={"ALOS_days": ("=AVERAGEIFS(tLOS[Length_of_Stay],tLOS[Department],[@Department])", "dec1", alos_d),
              "Readmit_rate": ('=COUNTIFS(tLOS[Department],[@Department],tLOS[Readmit_30d],"Yes")/[@Admissions]', "pct", rr_d),
              "Avg_bill": ("=AVERAGEIFS(tLOS[Billing_Amount],tLOS[Department],[@Department])", "naira", bill_d)},
        total={"Department": "Total", "Admissions": "sum"})
s.grid([["Right-align numbers, left-align text", "Consistent decimals per column (1 for ALOS, 1 for %)"],
        ["Units in headers, not in cells", "'ALOS (days)', 'Avg bill (₦)'"], ["Sort by the measure that matters", "Busiest department first"],
        ["Light gridlines, bold totals", "Table Style Medium 2 + total row"]], header=["Rule", "Applied above"])
s.band("4.8.5-8", "Page setup, PDF export, fonts and metadata (applied to THIS sheet - try File > Print preview)")
s.grid([["Orientation / size", "Landscape A4, fit to 1 page wide"], ["Print titles", "Rows 1-3 repeat on every page"],
        ["Header / footer", "Header: report title. Footer: 'Confidential - Lifecare internal' | page X of Y | file name"],
        ["Export", "File > Export > Create PDF/XPS > Options: 'Active sheet' or 'Entire workbook'; Standard quality"],
        ["Fonts", "Calibri / Arial embed everywhere; the ₦ symbol needs a font that contains it (Calibri does)"],
        ["Metadata", "File > Info > Properties: Title, Author, Company, Tags - this workbook already has them set"]], header=["Setting", "Value"])
s.ws.set_landscape()
s.ws.set_paper(9)
s.ws.fit_to_pages(1, 0)
s.ws.repeat_rows(0, 2)
s.ws.set_header("&L&\"Calibri,Bold\"Lifecare General Hospital - 2023 Operations Report&RETEH Week 4")
s.ws.set_footer("&LConfidential - Lifecare internal&CPage &P of &N&R&F")
s.band("4.8.9-10", "Version control and the review checklist")
s.grid([["v0.1", "2024-01-08", "Analyst", "First draft from extract LC_2023_v1"], ["v0.2", "2024-01-10", "Analyst", "Added readmission by department after QA comment"],
        ["v1.0", "2024-01-12", "Head of Quality", "Approved for board pack"]], header=["Version", "Date", "By", "Change"])
s.grid([["Numbers reconcile to source (admissions count = extract rows)", "☐"], ["Every chart has a finding title and source line", "☐"],
        ["No patient names or IDs in the report (aggregates only)", "☐"], ["Spell-check, consistent ₦ and % formats", "☐"],
        ["PDF opened and checked on another computer", "☐"]], header=["Check", "Done"])

# ======================================================================================
# 4.9 Slide presentations
# ======================================================================================
SH = "4.9 Slide Presentations"
s = b.sheet(SH, "4.9  Slide Presentations - a 10-slide storyboard on Eko Suites cancellations",
            "One message per slide, each with the number that proves it; title word counts and timing", "4.9.1 - 4.9.10")
hc = hod.copy()
hc["c"] = hc.Booking_Status == "Cancelled"
cr_all = hc.c.mean()
cr_ota = hc[hc.Market_Segment == "Online Travel Agent"].c.mean()
cr_rep = hc[hc.Repeat_Guest == "Yes"].c.mean()
lost = hc[hc.c].Booking_Value_NGN.sum()
sb = [["1", "Title", "Cancellations cost Eko Suites ₦" + f"{lost / 1e9:,.2f}" + "bn of bookings in 2023-24", '=SUMIFS(tHotel[Booking_Value_NGN],tHotel[Booking_Status],"Cancelled")', lost, "naira", 0.5],
      ["2", "Situation", f"One booking in three is cancelled ({cr_all:.0%})", '=COUNTIFS(tHotel[Booking_Status],"Cancelled")/ROWS(tHotel)', cr_all, "pct", 1],
      ["3", "Where", f"OTA bookings cancel most ({cr_ota:.0%})", '=COUNTIFS(tHotel[Market_Segment],"Online Travel Agent",tHotel[Booking_Status],"Cancelled")/COUNTIFS(tHotel[Market_Segment],"Online Travel Agent")', cr_ota, "pct", 1.5],
      ["4", "When", "Bookings made 6+ months ahead cancel three times in four", '=COUNTIFS(tHotel[Lead_Time_Days],">180",tHotel[Booking_Status],"Cancelled")/COUNTIFS(tHotel[Lead_Time_Days],">180")',
       ((hc.Lead_Time_Days > 180) & hc.c).sum() / (hc.Lead_Time_Days > 180).sum(), "pct", 1.5],
      ["5", "Who", f"Repeat guests almost never cancel ({cr_rep:.1%})", '=COUNTIFS(tHotel[Repeat_Guest],"Yes",tHotel[Booking_Status],"Cancelled")/COUNTIFS(tHotel[Repeat_Guest],"Yes")', cr_rep, "pct2", 1],
      ["6", "Signal", "Guests who make special requests cancel less", '=COUNTIFS(tHotel[Special_Requests],">0",tHotel[Booking_Status],"Cancelled")/COUNTIFS(tHotel[Special_Requests],">0")',
       ((hc.Special_Requests > 0) & hc.c).sum() / (hc.Special_Requests > 0).sum(), "pct", 1],
      ["7", "Option A", "Deposit for bookings > 90 days ahead", '=COUNTIFS(tHotel[Lead_Time_Days],">90")', int((hc.Lead_Time_Days > 90).sum()), "int", 1.5],
      ["8", "Option B", "Loyalty offer to convert first-time guests", '=COUNTIFS(tHotel[Repeat_Guest],"No")', int((hc.Repeat_Guest == "No").sum()), "int", 1],
      ["9", "Recommendation", "Pilot deposits on OTA bookings > 90 days from Q1", "", "", None, 1.5],
      ["10", "Next steps", "Owner, date, and the KPI we will track: cancellation rate", "", "", None, 0.5]]
sbd = pd.DataFrame([[a, b_, c] for a, b_, c, *_ in sb], columns=["Slide", "Role", "Title (one message)"])
sbd["Proof_number"] = ""
sbd["Minutes"] = [r[6] for r in sb]
fxs = {(i, "Proof_number"): (r[3], r[4]) for i, r in enumerate(sb) if r[3]}
top = s.frame(sbd, fmts={"Minutes": "dec1"}, formulas=fxs)
for i, r in enumerate(sb):
    if r[5]:
        s.ws.write_formula(top + i, 4, r[3], b.nf(r[5], style="result"), float(r[4]))
s.ws.set_column(3, 3, 60)
s.ws.set_column(4, 4, 16)
s.calc([("Total talk time (minutes)", f"=SUM(F{top + 1}:F{top + len(sb)})", sum(r[6] for r in sb), "dec1", "Target 12 minutes + 5 for questions (4.9.8)"),
        ("Words in slide 4 title", f'=LEN(TRIM(D{top + 4}))-LEN(SUBSTITUTE(TRIM(D{top + 4})," ",""))+1', len(sb[3][2].split()), "int", "Keep finding-titles under ~15 words (4.9.4)")])
s.band("4.9.3-9", "Projection, identity, tables, questions and accessibility")
s.grid([["Charts for projection", "≥ 18pt labels, one highlighted series, no 3-D, direct labels instead of legends"],
        ["Visual identity", "ETEH navy #0B2545, teal #13807A, gold #E0A526 only for emphasis"],
        ["Tables on slides", "Max 5 rows × 4 columns; move the full table to the appendix"],
        ["Handling questions", "Back-up slides: lead-time bands (4.3), segment Solver result (4.6), method notes"],
        ["Accessible decks", "Alt text on every chart, reading order checked, colour + label (not colour alone)"]], header=["Topic", "Rule for this deck"], merge_last=True)

# ======================================================================================
# 4.10 Auditing, protection, automation (this sheet is protected - no password)
# ======================================================================================
SH = "4.10 Audit & Protect"
s = b.sheet(SH, "4.10  Formula Auditing, Protection and Light Automation",
            "Find three planted errors, then see a protected sheet in action (Review > Unprotect Sheet, no password)", "4.10.1 - 4.10.8")
s.band("4.10.1-4", "Audit exercise: zone summary with THREE planted errors")
s.scenario("The 'Submitted' column came from a colleague. Use Formulas > Trace Precedents, Show Formulas (Ctrl+`) and Evaluate Formula to find what is wrong. "
           "The 'Correct' column and the check column show the answer.")
zs_ = s25.groupby("Zone").Sales.sum().reindex(["South-West", "South-South", "North-Central", "South-East"])
z0 = s.r + 1
s.ws.write_row(s.r, 1, ["Zone", "Sales (data)", "Share - submitted", "Share - correct", "Error found?"], F["hdr"])
s.r += 1
tot_row = z0 + 4
for i, (zn, v) in enumerate(zs_.items()):
    r_ = z0 + i + 1
    s.ws.write(r_ - 1, 1, zn, F["cell"])
    s.ws.write_formula(r_ - 1, 2, f'=SUMIFS(tDash25[Sales],tDash25[Zone],B{r_})', b.nf("naira"), float(v))
    if i == 2:
        s.ws.write_number(r_ - 1, 3, 0.15, b.nf("pct", style="bad"))   # hard-coded value
        sub_v = 0.15
    elif i == 3:
        s.ws.write_formula(r_ - 1, 3, f"=C{r_}/C{tot_row}", b.nf("pct", style="bad"), 1.0)   # relative ref drifted onto its own row
        sub_v = 1.0
    else:
        s.ws.write_formula(r_ - 1, 3, f"=C{r_}/C${tot_row + 1}", b.nf("pct"), float(v / zs_.iloc[:3].sum()))
        sub_v = v / zs_.iloc[:3].sum()
    s.ws.write_formula(r_ - 1, 4, f"=C{r_}/SUM($C${z0 + 1}:$C${z0 + 4})", b.nf("pct", style="result"), float(v / zs_.sum()))
    s.ws.write_formula(r_ - 1, 5, f'=IF(ROUND(D{r_}-E{r_},6)=0,"OK","ERROR")', F["cell"], "OK" if abs(sub_v - v / zs_.sum()) < 1e-6 else "ERROR")
s.ws.write(tot_row, 1, "Total", F["lab"])
s.ws.write_formula(tot_row, 2, f"=SUM(C{z0 + 1}:C{z0 + 3})", b.nf("naira", style="bad"), float(zs_.iloc[:3].sum()))
s.ws.write_formula(tot_row, 4, f"=SUM(E{z0 + 1}:E{z0 + 4})", b.nf("pct", style="result"), 1.0)
s.r = tot_row + 2
s.grid([["1", "Total =SUM(C..:C..) stops one row short - South-East is excluded, so every share is overstated", "Trace Precedents on the total shows the blue box missing a row"],
        ["2", "North-Central share is a typed 15% - not a formula", "Show Formulas (Ctrl+`) or Go To Special > Constants"],
        ["3", "South-East divides by itself (the relative reference to the total drifted up one row) - shows 100%", "Evaluate Formula steps into the wrong denominator; fix with a $C$ absolute reference"]],
       header=["Error", "What is wrong", "Tool that finds it"])
s.para("Watch Window (4.10.4): Formulas > Watch Window > Add the Total cell, then edit data on another sheet and watch it change.", "note")
s.band("4.10.5", "Protection - this sheet is protected: only the yellow cell can be edited")
unl = b.wb.add_format({"bg_color": "#FFF4CC", "border": 1, "border_color": "#E0A526", "locked": False, "num_format": "0.0%"})
s.ws.merge_range(s.r, 1, s.r, 3, "Growth assumption for FY2026 (editable)", F["lab"])
s.ws.write_number(s.r, 4, 0.12, unl)
gcell = xl_rowcol_to_cell(s.r, 4, True, True)
s.r += 1
s.ws.merge_range(s.r, 1, s.r, 3, "FY2026 sales plan (locked formula)", F["lab"])
s.ws.write_formula(s.r, 4, f"=SUM(tDash25[Sales])*(1+{gcell})", b.nf("naira", style="result"), float(s25.Sales.sum() * 1.12))
s.r += 2
s.ws.protect("", {"select_locked_cells": True, "select_unlocked_cells": True, "format_columns": True, "format_rows": True})
s.grid([["Lock cells", "All cells are locked by default; unlock inputs via Format Cells > Protection"],
        ["Protect sheet", "Review > Protect Sheet - stops formula overwrites (not a security feature)"],
        ["Protect workbook structure", "Stops sheets being deleted, renamed or unhidden"],
        ["Encrypt with password", "File > Info > Protect Workbook > Encrypt - the only real confidentiality control"]], header=["Level", "What it does"], merge_last=True)
s.band("4.10.6-8", "Light automation: a recorded macro and an Office Script")
s.code('''Sub Refresh_And_Export_Report()
    ' Recorded, then tidied: refresh all queries, recalculate, export the report sheet to PDF
    ThisWorkbook.RefreshAll
    Application.CalculateFull
    Worksheets("4.8 Report Writing").ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=ThisWorkbook.Path & "\\Lifecare_2023_Report_" & Format(Date, "yyyy-mm-dd") & ".pdf", _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, OpenAfterPublish:=False
End Sub''')
s.code('''// Office Script (Excel for the web > Automate > New Script): flag stores below target
function main(workbook: ExcelScript.Workbook) {
  const t = workbook.getTable("tStory");
  const att = t.getColumnByName("Attainment").getRangeBetweenHeaderAndTotal().getValues();
  let missed = 0;
  att.forEach(r => { if ((r[0] as number) < 1) missed++; });
  console.log(`Stores below target: ${missed}`);   // run from Power Automate on a schedule
}''')
s.grid([["Name", "Refresh_And_Export_Report"], ["Purpose", "Monthly refresh + PDF of the Lifecare report"], ["Trigger", "Manual button / first working day"],
        ["Inputs / outputs", "Power Query connections → PDF in the workbook folder"], ["Owner / last tested", "Analyst / 2024-01-12"],
        ["Save as", ".xlsm (macro-enabled); Office Scripts live in OneDrive"]], header=["Automation log", "Entry"], merge_last=True)

# ======================================================================================
# 4.11 AI co-pilots
# ======================================================================================
SH = "4.11 AI Co-pilots"
s = b.sheet(SH, "4.11  AI Co-pilots for Analysts - prompts, schema prompts, formulas and SQL",
            "Templates you can paste, an AI SQL answer that looks right but is wrong, and what must never be pasted", "4.11.1 - 4.11.10")
s.band("4.11.2", "Prompt structure for analytical tasks")
s.grid([["Role", "You are an Excel analyst helping a Nigerian retail chain."], ["Context", "Table tSal has columns Seniority, Gender, Annual_Salary_NGN (one row per respondent)."],
        ["Task", "Write ONE formula returning the average salary of female Senior respondents."], ["Format", "Return only the formula, then one sentence explaining it."],
        ["Constraints", "Excel 365; structured references; no helper columns."]], header=["Part", "Example"], merge_last=True)
s.band("4.11.3", "Schema prompt for supermart_online.db (generated from the real database - no data rows, so no personal data)")
with sqlite3.connect(W3D / "supermart_online.db") as con:
    schema_lines = []
    for (tname,) in con.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"):
        cols_ = [f"{r_[1]} {r_[2]}" for r_ in con.execute(f"PRAGMA table_info({tname})")]
        schema_lines.append(f"{tname}({', '.join(cols_)})")
s.code("Database: SQLite. Tables and columns:\n" + "\n".join(schema_lines) +
       "\nKeys: transactions.customer_id → customers; .product_id → products; .seller_id → sellers; .delivery_id → deliveries (NULL when cancelled).\n"
       "Revenue rule: net revenue = SUM(total_price) WHERE order_status IN ('Delivered','Completed').")
s.band("4.11.4", "AI-generated Excel formula - verified before use")
fem_sen = sal[(sal.Gender == "Female") & (sal.Seniority == "Senior")].Annual_Salary_NGN.mean()
s.calc([("AI answer (correct)", '=AVERAGEIFS(tSal[Annual_Salary_NGN],tSal[Gender],"Female",tSal[Seniority],"Senior")', fem_sen, "naira", "Check by filtering D_Salary: same number ✔"),
        ("AI answer (plausible but wrong)", '=AVERAGEIF(tSal[Gender],"Female",tSal[Annual_Salary_NGN])', sal[sal.Gender == "Female"].Annual_Salary_NGN.mean(), "naira", "Ignores Seniority - looks reasonable, answers a different question ✘")])
s.band("4.11.5", "AI-generated SQL: 'revenue and number of deliveries per customer zone'")
ai_sql = ("SELECT c.zone, SUM(t.total_price) AS revenue, COUNT(d.delivery_id) AS deliveries\nFROM transactions t\nJOIN customers c ON c.customer_id = t.customer_id\n"
          "JOIN deliveries d ON d.customer_id = t.customer_id      -- joins on CUSTOMER, not on the delivery key\nGROUP BY c.zone ORDER BY revenue DESC")
ok_sql = ("SELECT c.zone, SUM(t.total_price) AS revenue, COUNT(t.delivery_id) AS deliveries\nFROM transactions t\nJOIN customers c ON c.customer_id = t.customer_id\n"
          "GROUP BY c.zone ORDER BY revenue DESC")
with sqlite3.connect(W3D / "supermart_online.db") as con:
    bad = pd.read_sql_query(ai_sql, con)
    good = pd.read_sql_query(ok_sql, con)
cmpd = good.merge(bad, on="zone", suffixes=("_correct", "_AI"))
s.code(ai_sql)
s.code(ok_sql)
s.table(cmpd, "tAiSql", fmts={"revenue_correct": "naira", "deliveries_correct": "int", "revenue_AI": "naira", "deliveries_AI": "int"},
        calc={"Inflation": ("=[@revenue_AI]/[@revenue_correct]", "dec1", (cmpd.revenue_AI / cmpd.revenue_correct).tolist())},
        total={"zone": "Total", "revenue_correct": "sum", "revenue_AI": "sum"})
s.warn(f"The AI query returns ₦{bad.revenue.sum() / 1e9:,.2f}bn - about {bad.revenue.sum() / good.revenue.sum():.1f}x the true ₦{good.revenue.sum() / 1e9:,.2f}bn - "
       "because each order is repeated for every delivery that customer ever had. Control total check: SUM(total_price) over transactions must match.")
s.band("4.11.6-10", "Explaining, narrative, prohibited data, limitations and workflow")
s.grid([["Explain", "Paste a formula (not data): 'Explain =INDEX(tStory[Store_Name],MATCH(MIN(...),...,0)) step by step'"],
        ["Narrative", "Give the AI the aggregated table (12 store totals) and ask for a 3-sentence summary; then check every number"],
        ["Never paste", "Patient names, Patient_ID, phone, email, date of birth, Suicidal_Thoughts_Ever, salaries with names - into any public AI tool"],
        ["Limitations", "No access to your file unless you share it; can invent column names, functions (e.g. 'SUMIFSX') and statistics"],
        ["Workflow", "1 schema prompt → 2 ask → 3 run on a copy → 4 verify against a control total → 5 log prompt + check in the provenance log (4.12)"]],
       header=["Use", "Practice"], merge_last=True)

# ======================================================================================
# 4.12 Verifying AI output
# ======================================================================================
SH = "4.12 Verify AI Output"
s = b.sheet(SH, "4.12  Verifying AI Output and Handling Hallucinations",
            "Six AI claims about our data, each checked by a formula; control totals; provenance and sign-off", "4.12.1 - 4.12.8")
s.band("4.12.2-4", "Verification protocol in action - claims an AI assistant made about the bootcamp data")
claims = [
    ("SuperMart FY2025 sales were ₦733.2m", 733.2e6, "=SUM(tDash25[Sales])", s25.Sales.sum(), "naira", 0.001),
    ("Lagos stores (Ikeja + Lekki) made 30% of FY2025 sales", 0.30, '=(SUMIFS(tDash25[Sales],tDash25[Store_ID],"ST01")+SUMIFS(tDash25[Sales],tDash25[Store_ID],"ST02"))/SUM(tDash25[Sales])',
     s25[s25.Store_ID.isin(["ST01", "ST02"])].Sales.sum() / s25.Sales.sum(), "pct", 0.005),
    ("Median salary in the survey is ₦12.5m", 12.5e6, "=MEDIAN(tSal[Annual_Salary_NGN])", sal.Annual_Salary_NGN.median(), "naira", 0.001),
    ("Lifecare 2023 readmission rate was 10.3%", 0.103, '=COUNTIFS(tLOS[Readmit_30d],"Yes")/ROWS(tLOS)', rr23, "pct", 0.0005),
    ("Eko Suites cancellation rate is 22%", 0.22, '=COUNTIFS(tHotel[Booking_Status],"Cancelled")/ROWS(tHotel)', cr_all, "pct", 0.005),
    ("Correlation of attendance with exam score is 0.81", 0.81, "=CORREL(tCV[Attendance_Pct],tCV[Exam_Score])", cvd.Attendance_Pct.corr(cvd.Exam_Score), "dec3", 0.005)]
cl = pd.DataFrame([[c[0], c[1], float(c[3]), c[5]] for c in claims], columns=["AI claim", "Claimed_value", "Verified_value", "Tolerance"])
tcl = s.table(cl, "tClaims", widths={"AI claim": 50, "Claimed_value": 15},
              calc={"Verdict": ('=IF(ABS([@Verified_value]-[@Claimed_value])<=[@Tolerance]*IF(ABS([@Claimed_value])>1,ABS([@Claimed_value]),1),"VERIFIED","WRONG - do not use")', None,
                                ["VERIFIED" if abs(c[3] - c[1]) <= c[5] * (abs(c[1]) if abs(c[1]) > 1 else 1) else "WRONG - do not use" for c in claims])})
for i, c in enumerate(claims):
    s.ws.write(tcl.r0 + 1 + i, tcl.ci("Claimed_value"), c[1], b.nf(c[4], style="plain"))
    s.ws.write_formula(tcl.r0 + 1 + i, tcl.ci("Verified_value"), c[2], b.nf(c[4], style="result"), float(c[3]))
    s.ws.write_number(tcl.r0 + 1 + i, tcl.ci("Tolerance"), c[5], b.nf("0.0%" if c[5] >= 0.001 else "0.00%", style="plain"))
s.ws.conditional_format(tcl.r0 + 1, tcl.ci("Verdict"), tcl.r0 + len(claims), tcl.ci("Verdict"),
                        {"type": "text", "criteria": "begins with", "value": "WRONG", "format": b.wb.add_format({"bg_color": "#F8D7DA", "font_color": "#7B1F28", "bold": True})})
s.para("Tolerance: relative for money/counts (0.1% of the claim), absolute for rates and correlations. Every claim is re-computed from the data, never from memory.", "note")
s.band("4.12.3", "The control total discipline - an AI-built zone summary that 'lost' a zone")
ai_zone = s25.groupby("Zone").Sales.sum().drop("South-East")
s.calc([("Control total: FY2025 sales (ledger)", "=SUM(tDash25[Sales])", s25.Sales.sum(), "naira", "From the source, never from the summary"),
        ("Sum of the AI's zone table", float(ai_zone.sum()), float(ai_zone.sum()), "naira", "AI summary listed only 3 zones"),
        ("Difference", f"=SUM(tDash25[Sales])-{ai_zone.sum():.0f}", s25.Sales.sum() - ai_zone.sum(), "naira", "≠ 0 → South-East was silently dropped; discard and redo")])
s.band("4.12.5-8", "Provenance, discard rules, sign-off and checklist")
s.grid([["2024-01-12 10:40", "AI assistant (web)", "Schema prompt + 'zone revenue SQL'", "Joined deliveries on customer_id - 56,725 rows, revenue ₦7.19bn (~5.5x)", "Discarded; rewrote join", "A. Analyst"],
        ["2024-01-12 11:05", "AI assistant (web)", "'Explain RANK.EQ ties'", "Explanation matched Microsoft docs", "Used in notes", "A. Analyst"]],
       header=["When", "Tool", "Prompt (no data)", "Check performed", "Outcome", "By"])
s.grid([["Discard entirely", "Invented columns/functions; fails a control total twice; cannot explain its own logic"],
        ["Sign-off", "The analyst who publishes owns the number - 'the AI said so' is never a defence"],
        ["Checklist", "☐ schema only  ☐ re-ran on data  ☐ control total matches  ☐ edge cases (blanks, cancelled)  ☐ logged  ☐ reviewer initials"]],
       header=["Rule", "Detail"], merge_last=True)

# ======================================================================================
# 4.13 Ethics, PII, PHI
# ======================================================================================
SH = "4.13 Ethics PII PHI"
s = b.sheet(SH, "4.13  Data Ethics, PII and PHI",
            "Classify every sensitive column in the bootcamp data, minimise, suppress small cells and test for unfair gaps", "4.13.1 - 4.13.9")
s.band("4.13.2-3", "Inventory: personal (PII) and health (PHI) data in the bootcamp datasets")
pii = [["SuperMart Online customers", "first_name, last_name, email, phone_number", "Direct identifier (PII)", "Never leave the database; analyse by customer_id only"],
       ["SuperMart Online customers", "date_of_birth, city", "Quasi-identifier", "Band age (18-24, 25-34…); use state/zone, not city, for small groups"],
       ["Lifecare patients / admissions", "Patient_Name, Patient_ID", "Direct identifier (PHI)", "Remove before any extract leaves the hospital"],
       ["Lifecare admissions", "Medical_Condition, Medication, Test_Result", "Health data (PHI, special category)", "Aggregate only; minimum 5 per cell"],
       ["Lifecare screening", "HbA1c_Level, Blood_Glucose_mg_dL, Diabetes", "Health data (PHI)", "Pseudonymised Screening_ID only"],
       ["Crestview wellbeing survey", "Suicidal_Thoughts_Ever, Depression_Screen_Positive", "Highly sensitive", "Restricted file; aggregate tables with suppression only"],
       ["ETEH salary survey", "Annual_Salary_NGN + Job_Title + Age", "Quasi-identifiers in combination", "Publish bands, not exact salaries, for rare job titles"]]
s.table(pd.DataFrame(pii, columns=["Dataset", "Columns", "Classification", "Handling rule"]), "tPII",
        widths={"Dataset": 26, "Columns": 40, "Classification": 26, "Handling rule": 50})
s.band("4.13.4", "Data minimisation - question: 'Which departments have the highest readmission rate?'")
needed = ["Department", "Readmit_30d", "Admission_Date (year only)"]
s.calc([("Columns available in the admissions extract", len(lc.columns), len(lc.columns), "int", ", ".join(lc.columns[:8]) + " …"),
        ("Columns actually needed", len(needed), len(needed), "int", ", ".join(needed)),
        ("Columns to drop before analysis", f"={len(lc.columns)}-{len(needed)}", len(lc.columns) - len(needed), "int", "Including Patient_Name, Doctor_Name, Billing_Amount")])
s.band("4.13.5", "Anonymisation and aggregation: wellbeing survey with small-cell suppression (n < 5 → '<5')")
wb_ = L("cv_wellbeing")
agg = wb_.groupby("Programme").agg(Respondents=("Respondent_ID", "count"), Screen_positive=("Depression_Screen_Positive", lambda v: (v == "Yes").sum())).reset_index()
agg = agg.sort_values("Respondents", ascending=False).reset_index(drop=True)
s.table(agg, "tSuppress", fmts={"Respondents": "int", "Screen_positive": "int"}, widths={"Programme": 18},
        calc={"Publishable_count": ('=IF([@Screen_positive]<5,"<5",[@Screen_positive])', None, ["<5" if v < 5 else int(v) for v in agg.Screen_positive]),
              "Publishable_rate": ('=IF([@Respondents]<10,"suppressed",[@Screen_positive]/[@Respondents])', "pct", ["suppressed" if n < 10 else p / n for n, p in zip(agg.Respondents, agg.Screen_positive)])})
qi = wb_.assign(AgeBand=pd.cut(wb_.Age, [0, 20, 25, 30, 100]).astype(str)).groupby(["Gender", "Age", "State_of_Origin", "Programme"]).size()
qi2 = wb_.assign(AgeBand=pd.cut(wb_.Age, [0, 20, 25, 30, 100]).astype(str)).groupby(["Gender", "AgeBand", "Programme"]).size()
s.calc([("Respondents unique on Gender + exact Age + State + Programme", int((qi == 1).sum()), int((qi == 1).sum()), "int", "Each could be re-identified by a classmate"),
        ("…unique after banding age and dropping State", int((qi2 == 1).sum()), int((qi2 == 1).sum()), "int", "Generalisation shrinks the risk (k-anonymity)")])
s.para("Only aggregated counts appear in this workbook - the respondent-level sensitive file stays in 00_Core_Datasets with restricted access.", "note")
s.band("4.13.7", "Bias and fairness check - salary gap by seniority (from 4.3, restated as a fairness test)")
s.grid([["Raw gap (all levels)", f"{1 - fem[-1] / mal[-1]:.1%}", "What a headline would quote"],
        ["Largest within-level gap", f"{max(1 - f_ / m_ for f_, m_ in zip(fem[:-1], mal[:-1])):.1%}", "The fairness issue to investigate"],
        ["'Other' gender respondents", f"{int((sal.Gender == 'Other').sum())}", "Too few to report - suppress, do not drop silently"]], header=["Measure", "Value", "Interpretation"])
s.band("4.13.6 / 4.13.8-9", "Legal basis, retention and governance (Nigeria Data Protection Act 2023)")
s.grid([["Purpose & legal basis", "Lifecare: provision of care / public interest in health; SuperMart: contract + legitimate interest; survey: consent"],
        ["Retention", "Training extracts deleted at the end of the bootcamp; production data per organisation policy"],
        ["Secure handling", "Encrypted laptops, no personal data in email or public AI tools, shared links restricted"],
        ["Accountability", "Data owner approves extracts; analyst documents minimisation; DPO reviews new uses"]], header=["Principle", "In practice"], merge_last=True)

# ======================================================================================
# 4.14 Internal benchmarks, PDF extraction, reconciliation (rewritten - no external sources)
# ======================================================================================
SH = "4.14 Benchmarks & Recon"
s = b.sheet(SH, "4.14  Internal Benchmarks, PDF Extraction and Reconciliation",
            "Benchmark stores against the chain, reconcile the PDF bulletin to the ledger, align definitions, quantify uncertainty", "4.14.1 - 4.14.5")
s.band("4.14.1", "Beginner: benchmark fitness - each store against the chain median (same period, same definitions)")
bm = s25.groupby(["Store_ID", "Store_Name"]).agg(Sales=("Sales", "sum"), Profit=("Profit", "sum")).reset_index().merge(stores[["Store_ID", "Floor_Area_sqm"]], on="Store_ID")
spsq = (bm.Sales / bm.Floor_Area_sqm)
mar = bm.Profit / bm.Sales
s.table(bm, "tBench", fmts={"Sales": "naira", "Profit": "naira", "Floor_Area_sqm": "int"}, widths={"Store_Name": 24},
        calc={"Sales_per_sqm": ("=[@Sales]/[@Floor_Area_sqm]", "naira", spsq.tolist()),
              "vs_median_sqm": ("=[@Sales_per_sqm]/MEDIAN([Sales_per_sqm])-1", "pct", (spsq / spsq.median() - 1).tolist()),
              "Margin": ("=[@Profit]/[@Sales]", "pct", mar.tolist()),
              "vs_median_margin_pts": ("=[@Margin]-MEDIAN([Margin])", "pct", (mar - mar.median()).tolist())})
s.grid([["Same definition", "Net sales after discount, excl. VAT, for every store"], ["Same period", "1 Jan - 31 Dec 2025 for every store"],
        ["Comparable units", "Sales per m² removes the size effect - Ikeja is big, not necessarily efficient"],
        ["Median, not mean", "One very large store would drag a mean benchmark"]], header=["Fitness test", "Applied here"], merge_last=True)
s.band("4.14.2", "Beginner: extract and verify the PDF - SuperMart_Monthly_Sales_Bulletin_2025.pdf (Week2_Data_Files)")
s.scenario("Extract Table 1 with Data > Get Data > From File > From PDF (Week 2, 2.4), keep the FY 2025 column, then compare with the ledger. "
           "The bulletin values below are exactly as printed (₦m, 1 decimal).")
ct = s25.assign(M=s25.Order_Date.dt.month).pivot_table(index="Store_Name", columns="M", values="Sales", aggfunc="sum", fill_value=0) / 1e6
ct = ct.reindex(stores.Store_Name)
bul = pd.DataFrame({"Store_Name": ct.index, "Bulletin_FY_m": [round(v, 1) for v in ct.sum(axis=1)],
                    "Sum_of_printed_months_m": [round(sum(round(x, 1) for x in r), 1) for r in ct.values]})
led = [s25[s25.Store_Name == n].Sales.sum() for n in bul.Store_Name]
diff = [bv - lv / 1e6 for bv, lv in zip(bul.Bulletin_FY_m, led)]
s.table(bul, "tBulletin", fmts={"Bulletin_FY_m": "0.0", "Sum_of_printed_months_m": "0.0"}, widths={"Store_Name": 24},
        calc={"Ledger_m": ("=SUMIFS(tDash25[Sales],tDash25[Store_Name],[@Store_Name])/1000000", "0.000", [v / 1e6 for v in led]),
              "Diff_m": ("=[@Bulletin_FY_m]-[@Ledger_m]", "0.000", diff),
              "Status": ('=IF(ABS([@Diff_m])<=0.05,"Rounding only","INVESTIGATE")', None, ["Rounding only" if abs(d) <= .05 else "INVESTIGATE" for d in diff]),
              "Months_vs_FY_m": ("=[@Sum_of_printed_months_m]-[@Bulletin_FY_m]", "0.0", (bul.Sum_of_printed_months_m - bul.Bulletin_FY_m).tolist())},
        total={"Store_Name": "Total", "Bulletin_FY_m": "sum", "Ledger_m": "sum"})
s.insight("Every store agrees within ±₦0.05m (rounding to one decimal). Adding the 12 printed monthly figures can differ from the printed FY total by ₦0.1m - "
          "that is why you reconcile to the ledger, not to a sum of rounded numbers.")
s.band("4.14.3", "Intermediate: reconcile time periods and definitions")
fy_apr_prev = sm[(sm.Order_Date >= "2024-04-01") & (sm.Order_Date <= "2025-03-31")].Sales.sum()
gross_ = tx[tx.order_date.dt.year == 2025].total_price.sum()
net_ = tx[(tx.order_date.dt.year == 2025) & tx.order_status.isin(["Delivered", "Completed"])].total_price.sum()
s.calc([("SuperMart calendar FY2025 (Jan-Dec)", "=SUM(tDash25[Sales])", s25.Sales.sum(), "naira", "What the bulletin reports"),
        ("SuperMart Apr 2024 - Mar 2025 (an April year-end)", float(fy_apr_prev), float(fy_apr_prev), "naira", "Different 12 months → different number; label the period"),
        ("SuperMart Online 2025 gross bookings", float(gross_), float(gross_), "naira", "All orders, including cancelled and returned"),
        ("SuperMart Online 2025 net revenue", float(net_), float(net_), "naira", "Delivered + Completed only"),
        ("Gross ÷ net", f"={gross_:.0f}/{net_:.0f}", gross_ / net_, "dec2", "Comparing one company's gross with another's net would overstate by this factor")])
s.band("4.14.4", "Intermediate: uncertainty and missing data")
dsm = L("dirty_sm")
qmiss = int(pd.to_numeric(dsm.Quantity, errors="coerce").isna().sum())
s.para(f"Missing data: the Week 2 dirty extract has {qmiss} lines with no Quantity. Report the total units as a range, not a single number, until the store confirms the values.", "plain")
dep_ci = los.groupby("Department").agg(n=("Admission_ID", "count"), k=("Readmit_30d", lambda v: (v == "Yes").sum())).reset_index()
dep_ci["p"] = dep_ci.k / dep_ci.n
moe = 1.96 * np.sqrt(dep_ci.p * (1 - dep_ci.p) / dep_ci.n)
s.table(dep_ci[["Department", "n", "k"]].rename(columns={"n": "Admissions", "k": "Readmitted"}), "tReadmitCI", fmts={"Admissions": "int", "Readmitted": "int"}, widths={"Department": 26},
        calc={"Rate": ("=[@Readmitted]/[@Admissions]", "pct", dep_ci.p.tolist()),
              "Margin_95": ("=1.96*SQRT([@Rate]*(1-[@Rate])/[@Admissions])", "pct", moe.tolist()),
              "Low_95": ("=[@Rate]-[@Margin_95]", "pct", (dep_ci.p - moe).tolist()),
              "High_95": ("=[@Rate]+[@Margin_95]", "pct", (dep_ci.p + moe).tolist())})
_ci = dep_ci.assign(lo=dep_ci.p - moe, hi=dep_ci.p + moe).set_index("Department")
_top = _ci.p.idxmax()
_sep = [d for d in _ci.index if d != _top and _ci.loc[d, "hi"] < _ci.loc[_top, "lo"]]
s.insight(f"{_top}'s interval ({_ci.loc[_top, 'lo']:.1%}-{_ci.loc[_top, 'hi']:.1%}) sits clearly above {' and '.join(_sep)} - a real difference worth a clinical review. "
          "The other departments' intervals overlap each other, so do not rank them as if their small gaps were real.")
s.band("4.14.5", "Advanced practical: benchmark decision memo")
low = bm.assign(sq=spsq).sort_values("sq").iloc[0]
s.grid([["Decision needed", "Which stores get the FY2026 refurbishment budget?"],
        ["Benchmark used", "Internal: chain median sales per m² and margin, FY2025 ledger (reconciled to the PDF bulletin within ±₦0.05m)"],
        ["Finding", f"{low.Store_Name} has the lowest sales per m² (₦{low.sq:,.0f}) - {abs(low.sq / spsq.median() - 1):.0%} below the chain median"],
        ["Uncertainty", "One year of data; store refurbishments and openings not adjusted; 2025 figures unaudited"],
        ["Recommendation", "Shortlist the bottom three stores on sales per m² for site review before allocating budget"],
        ["Sign-off", "Analyst / Finance reviewer / Date"]], header=["Memo section", "Content"], merge_last=True)

# ---------------------------------------------------------------------- data sheets
b.data_sheet("D_Salary", sal, "tSal", fmts={"Annual_Salary_NGN": "naira", "Age": "int", "Years_of_Experience": "dec1", "Is_Masters": "int", "Is_PhD": "int"},
             title="ETEH Talent Salary Survey (1,781) + 0/1 dummy columns for regression", note="Is_Masters / Is_PhD = 1 when Education_Level is Master's / PhD")
b.data_sheet("D_Students", cvd, "tCV", title="Crestview University - student performance 2024/25 (6,607)")
b.data_sheet("D_Orders", orders, "tOrders", fmts={"order_date": "date", "total_price": "naira"}, title="SuperMart Online orders (11,000) with product category")
b.data_sheet("D_LOS_2023", los, "tLOS", fmts={"Admission_Date": "date", "Billing_Amount": "naira"}, title="Lifecare admissions 2023 (no patient names or IDs - minimised extract)")
b.data_sheet("D_Hotel", hod, "tHotel", fmts={"Arrival_Date": "date", "Avg_Room_Rate_NGN": "naira", "Booking_Value_NGN": "naira"}, title="Eko Suites bookings 2023-2024 (36,275)")
b.data_sheet("D_Pareto_FY25", cust.assign(Cum_Share=cum), "tPareto", fmts={"Sales": "naira", "Cum_Share": "pct", "Rank": "int"},
             title="SuperMart FY2025 customers ranked by sales (Cum_Share = share of sales from customers at least this large)")
dash = s25[["Order_ID", "Order_Date", "Store_ID", "Store_Name", "Zone", "Category", "Sales", "Profit"]]
b.data_sheet("D_Sales_2025", dash, "tDash25", fmts={"Order_Date": "date", "Sales": "naira", "Profit": "naira"}, title="SuperMart FY2025 sales lines (3,312)")

b.cover("Week 4 Section Examples - Statistics, Forecasting, What-If, Communication, AI and Ethics",
        ["One worksheet per module (4.1 - 4.14). Every statistic is a live formula over the D_ tables; yellow cells are inputs you can change.",
         "4.5 forecasts SuperMart's monthly sales with FORECAST.ETS and backtests it on 2025; 4.6 is a full Eko Suites revenue model with Goal Seek, "
         "a sensitivity grid, scenarios and a Solver allocation.",
         "4.10 is a protected sheet (no password) with three planted formula errors to find. 4.14 uses internal benchmarks and the Week 2 PDF bulletin only.",
         "Sensitive data: only minimised or aggregated extracts appear here - no patient names, contact details or respondent-level wellbeing answers."],
        [(k, DATASETS_TEXT[k]) for k in ["Salary Survey", "Crestview", "SuperMart Stores", "SuperMart Online", "Lifecare", "Eko Suites", "Crestview Wellbeing"]])
b.close()
print("Week 4 ->", OUT)
