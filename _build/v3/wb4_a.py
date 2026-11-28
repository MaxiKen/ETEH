"""Week 4 workbook, modules 4.1 - 4.4 (edition 3): descriptive statistics, distributions, correlation and regression.

Every result is a live formula on the Data_ tables; cached values are worked out here with Excel's definitions (QUARTILE.INC
= linear interpolation, STDEV.S, SKEW and KURT with the sample corrections), and verify_xlsx.py recomputes them independently.
Charts: a histogram of exam scores, the Pareto curve, store z-scores, two scatterplots (one with a trendline) and residuals.
"""
from __future__ import annotations

import math

import numpy as np
import pandas as pd
from scipy import stats

from w1x import ModuleSheet
from wb4_common import N, f, tp

S_ = "tSal[Annual_Salary_NGN]"
O_ = "tOrders[total_price]"
E_ = "tCV[Final_Score]"
L_ = "tLOS[LOS_Days]"


def q(s, k):
    return float(np.percentile(s, k * 25))


def pc(s, p):
    return float(np.percentile(s, p * 100))


def sd(s):
    return float(pd.Series(s).std(ddof=1))


def mode1(s):
    """MODE.SNGL: ties go to the value that appears first."""
    s = list(s)
    c = pd.Series(s).value_counts()
    top = c.max()
    return next(v for v in s if c[v] == top)


def dref(sheet, df, col):
    """Chart reference to one column of a Data_ table that starts in A1."""
    j = list(df.columns).index(col)
    return [sheet, 1, j, len(df), j]


def prank(arr, x, sig=3):
    v = sorted(arr)
    r = sum(1 for a in v if a < x)
    fct = 10 ** sig
    return math.floor(r / (len(v) - 1) * fct) / fct


def build(b, D):
    m41(b, D)
    m42(b, D)
    m43(b, D)
    m44(b, D)


# =====================================================================================================
def m41(b, D):
    S, O, CV, L = D.S.Annual_Salary_NGN, D.O.total_price, D.CV.Final_Score, D.LOS.LOS_Days
    m = ModuleSheet(b, "4.1 Descriptive Stats", "4.1  Descriptive Statistics",
                    "Tutor walk-through after Module 4.1: centre, spread and shape, worked live on four real columns", "4.1.1 - 4.1.14")
    m.dataset([("tSal", "Sheet Data_Sal (A1)", f"{len(S):,} salary survey respondents: Annual_Salary_NGN is the main example."),
               ("tOrders", "Sheet Data_Orders (A1)", f"{len(O):,} SuperMart Online order lines: total_price is the skewed example."),
               ("tCV", "Sheet Data_CV (A1)", f"{len(CV):,} Kingsway students: Final_Score."),
               ("tLOS", "Sheet Data_LOS (A1)", f"{len(L):,} Mercyland 2023 admissions: LOS_Days (length of stay).")])

    tp(m, "4.1.1",
       [("Count (size)", f"=COUNT({S_})", len(S), "int", "How many values the summary describes."),
        ("Median (centre)", f"=MEDIAN({S_})", f(S.median()), N, "Half earn less, half more."),
        ("Q1 (spread: bottom of the middle 50%)", f"=QUARTILE.INC({S_},1)", q(S, 1), N, ""),
        ("Q3 (spread: top of the middle 50%)", f"=QUARTILE.INC({S_},3)", q(S, 3), N, ""),
        ("Skewness (shape)", f"=SKEW({S_})", f(S.skew()), "dec2", "Near 0: balanced.")],
       [("n", "How many exam scores are in tCV?", f"=COUNT({E_})", len(CV), "int", "Count of exam scores"),
        ("n", "What is the median exam score?", f"=MEDIAN({E_})", f(CV.median()), "dec1", "Median exam score")],
       steps=["Pick one numeric column.", "Report one number for each family: size, centre, spread and shape.",
              "Use the median and quartiles when the column is skewed (4.1.8)."])

    big10 = sorted(O, reverse=True)[9]
    trim = (O.sum() - O[O >= big10].sum()) / (len(O) - (O >= big10).sum())
    tp(m, "4.1.2",
       [("Mean salary", f"=AVERAGE({S_})", f(S.mean()), N, "Close to the median: salaries are balanced."),
        ("Mean order value", f"=AVERAGE({O_})", f(O.mean()), N, "Far above the median order (4.1.3)."),
        ("Mean without the 10 largest orders", f'=(SUM({O_})-SUMIF({O_},">="&LARGE({O_},10)))/(COUNT({O_})-COUNTIF({O_},">="&LARGE({O_},10)))',
         f(trim), N, "Ten orders out of 11,000 move the mean.")],
       [("n", "What is the mean exam score? (1 decimal)", f"=AVERAGE({E_})", f(CV.mean()), "dec1", "Mean exam score"),
        ("n", "What is the mean length of stay in 2023? (2 decimals)", f"=AVERAGE({L_})", f(L.mean()), "dec2", "Mean length of stay")],
       steps=["Type =AVERAGE( and select the column (a structured reference such as tSal[Annual_Salary_NGN]).",
              "Compare it with the median: a big gap means outliers or a long tail."])

    tp(m, "4.1.3",
       [("Median order (typical order)", f"=MEDIAN({O_})", f(O.median()), N, "Half of all orders are smaller."),
        ("Mean order", f"=AVERAGE({O_})", f(O.mean()), N, ""),
        ("Mean ÷ median", f"=AVERAGE({O_})/MEDIAN({O_})", f(O.mean() / O.median()), "dec2", "Above about 1.2: report the median.")],
       [("n", "What is the median length of stay in 2023 (days)?", f"=MEDIAN({L_})", f(L.median()), "dec1", "Median length of stay"),
        ("n", "What is the median salary?", f"=MEDIAN({S_})", f(S.median()), N, "Median salary")])

    ms, mq, ml = mode1(S), mode1(D.O.quantity), mode1(L)
    tp(m, "4.1.4",
       [("Most common salary", f"=MODE.SNGL({S_})", f(ms), N, "MODE.SNGL returns one value."),
        ("Respondents on that salary", f"=COUNTIF({S_},MODE.SNGL({S_}))", int((S == ms).sum()), "int", ""),
        ("Respondents at Senior level (the modal level)", '=COUNTIF(tSal[Seniority],"Senior")', int((D.S.Seniority == "Senior").sum()), "int",
         "For text, count each level; the largest count is the mode."),
        ("Most common order quantity", "=MODE.SNGL(tOrders[quantity])", int(mq), "int", "")],
       [("n", "What is the most common length of stay in 2023 (days)?", f"=MODE.SNGL({L_})", int(ml), "int", "Modal length of stay"),
        ("n", "How many 2023 admissions had exactly that length of stay?", f"=COUNTIF({L_},MODE.SNGL({L_}))", int((L == ml).sum()), "int",
         "Admissions at the mode")])

    tp(m, "4.1.5",
       [("Minimum", f"=MIN({S_})", f(S.min()), N, ""), ("Maximum", f"=MAX({S_})", f(S.max()), N, ""),
        ("Range", f"=MAX({S_})-MIN({S_})", f(S.max() - S.min()), N, "Uses only the two extremes."),
        ("IQR (Q3 - Q1)", f"=QUARTILE.INC({S_},3)-QUARTILE.INC({S_},1)", q(S, 3) - q(S, 1), N, "The middle 50%: not moved by outliers.")],
       [("n", "What is the IQR of exam scores?", f"=QUARTILE.INC({E_},3)-QUARTILE.INC({E_},1)", q(CV, 3) - q(CV, 1), "dec1", "IQR of exam scores"),
        ("n", "What is the range of length of stay in 2023 (days)?", f"=MAX({L_})-MIN({L_})", int(L.max() - L.min()), "int", "Range of stay")])

    tp(m, "4.1.6",
       [("Variance (sample)", f"=VAR.S({S_})", f(pd.Series(S).var(ddof=1)), N, "In squared naira: hard to read."),
        ("Standard deviation (sample)", f"=STDEV.S({S_})", sd(S), N, "Typical distance from the mean, in naira.")],
       [("n", "What is the standard deviation of exam scores? (2 decimals)", f"=STDEV.S({E_})", sd(CV), "dec2", "SD of exam scores"),
        ("n", "What is the standard deviation of length of stay? (2 decimals)", f"=STDEV.S({L_})", sd(L), "dec2", "SD of length of stay")],
       note="STDEV.S treats the data as a sample (divides by n - 1); STDEV.P divides by n. With thousands of rows the difference is tiny.")

    tp(m, "4.1.7",
       [("CV of order value", f"=STDEV.S({O_})/AVERAGE({O_})", sd(O) / O.mean(), "pct", "Varies most."),
        ("CV of exam score", f"=STDEV.S({E_})/AVERAGE({E_})", sd(CV) / CV.mean(), "pct", "Varies least.")],
       [("n", "What is the coefficient of variation of salary? (as a %)", f"=STDEV.S({S_})/AVERAGE({S_})", sd(S) / S.mean(), "pct", "CV of salary"),
        ("n", "What is the coefficient of variation of length of stay?", f"=STDEV.S({L_})/AVERAGE({L_})", sd(L) / L.mean(), "pct", "CV of stay")])

    tp(m, "4.1.8",
       [("Skewness, order value", f"=SKEW({O_})", f(O.skew()), "dec2", "Strong right tail."),
        ("Skewness, salary", f"=SKEW({S_})", f(S.skew()), "dec2", "Roughly balanced.")],
       [("n", "What is the skewness of length of stay? (2 decimals)", f"=SKEW({L_})", f(L.skew()), "dec2", "Skewness of stay"),
        ("t", "Is length of stay right-skewed, left-skewed or balanced?", "right-skewed", None, None, ("right", "right skewed", "positive"))],
       note="Rule of thumb: between -0.5 and +0.5 is roughly balanced; above +1 is a strong right tail (a few very large values).")

    tp(m, "4.1.9",
       [("Kurtosis, order value", f"=KURT({O_})", f(O.kurt()), "dec2", "Very heavy tails: extreme orders are common."),
        ("Kurtosis, salary", f"=KURT({S_})", f(S.kurt()), "dec2", "Below 0: flatter than a normal curve.")],
       [("n", "What is the kurtosis of length of stay? (2 decimals)", f"=KURT({L_})", f(L.kurt()), "dec2", "Kurtosis of stay"),
        ("n", "What is the kurtosis of exam scores? (2 decimals)", f"=KURT({E_})", f(CV.kurt()), "dec2", "Kurtosis of exam scores")],
       note="Excel's KURT reports excess kurtosis: a normal curve scores 0.")

    tp(m, "4.1.10",
       [("P10", f"=PERCENTILE.INC({S_},0.1)", pc(S, .1), N, "10% earn less."), ("P50 (the median)", f"=PERCENTILE.INC({S_},0.5)", pc(S, .5), N, ""),
        ("P90", f"=PERCENTILE.INC({S_},0.9)", pc(S, .9), N, "The top 10% start here.")],
       [("n", "What is the 90th percentile order value?", f"=PERCENTILE.INC({O_},0.9)", pc(O, .9), N, "P90 order value"),
        ("n", "What is the 25th percentile exam score?", f"=PERCENTILE.INC({E_},0.25)", pc(CV, .25), "dec1", "P25 exam score")])

    q1, q3 = q(O, 1), q(O, 3)
    fence = q3 + 1.5 * (q3 - q1)
    lq1, lq3 = q(L, 1), q(L, 3)
    lf = lq3 + 1.5 * (lq3 - lq1)
    fx_f = f"(QUARTILE.INC({O_},3)+1.5*(QUARTILE.INC({O_},3)-QUARTILE.INC({O_},1)))"
    tp(m, "4.1.11",
       [("Q1", f"=QUARTILE.INC({O_},1)", q1, N, ""), ("Q3", f"=QUARTILE.INC({O_},3)", q3, N, ""),
        ("Upper fence = Q3 + 1.5 x IQR", f"={fx_f}", fence, N, "Values above are outliers by this rule."),
        ("Orders above the fence", f'=COUNTIF({O_},">"&{fx_f})', int((O > fence).sum()), "int", ""),
        ("Their share of order value", f'=SUMIF({O_},">"&{fx_f})/SUM({O_})', f(O[O > fence].sum() / O.sum()), "pct",
         "Outliers here are real big orders: investigate, do not delete.")],
       [("n", "What is the upper fence for length of stay (days)?", f"=QUARTILE.INC({L_},3)+1.5*(QUARTILE.INC({L_},3)-QUARTILE.INC({L_},1))", lf,
         "dec1", "Upper fence for stay"),
        ("n", "How many 2023 admissions stayed longer than that fence?",
         f'=COUNTIF({L_},">"&(QUARTILE.INC({L_},3)+1.5*(QUARTILE.INC({L_},3)-QUARTILE.INC({L_},1))))', int((L > lf).sum()), "int", "Long-stay outliers")],
       steps=["Work out Q1 and Q3 with QUARTILE.INC.", "IQR = Q3 - Q1; fences = Q1 - 1.5 x IQR and Q3 + 1.5 x IQR.",
              "Count and inspect the values outside the fences before deciding what to do."])

    tp(m, "4.1.12", None, [],
       steps=True,
       do=[("Turn on the Analysis ToolPak (File › Options › Add-ins › Manage Excel Add-ins › Go › tick Analysis ToolPak).",
            "Data › Data Analysis appears at the right of the Data tab."),
           ("Open Data › Data Analysis and scroll the list of tools.", "Descriptive Statistics, Histogram, Correlation and Regression are listed.")])

    se = sd(S) / math.sqrt(len(S))
    tp(m, "4.1.13",
       [("Standard error = STDEV.S / SQRT(COUNT)", f"=STDEV.S({S_})/SQRT(COUNT({S_}))", se, N, "How precisely the mean is known."),
        ("Confidence level (95%) = CONFIDENCE.T", f"=CONFIDENCE.T(0.05,STDEV.S({S_}),COUNT({S_}))",
         f(stats.t.ppf(.975, len(S) - 1) * se), N, "The mean is within this of the true mean, 95% of the time.")],
       [("n", "What is the standard error of the mean exam score? (3 decimals)", f"=STDEV.S({E_})/SQRT(COUNT({E_}))", sd(CV) / math.sqrt(len(CV)), "dec3",
         "Standard error of exam mean", 0.0006)],
       do=[("Data › Data Analysis › Descriptive Statistics. Input range: Data_Sal column K with its header; tick Labels in first row, "
            "Summary statistics and Confidence Level (95%). Output to a new sheet.",
            "Mean, Standard Error, Median, Mode, SD, Kurtosis, Skewness and Count match the Worked solutions in 4.1.2 - 4.1.13.")])

    tp(m, "4.1.14",
       [("Median order", f"=MEDIAN({O_})", f(O.median()), N, "The typical customer."),
        ("Orders above the outlier fence", f'=COUNTIF({O_},">"&{fx_f})', int((O > fence).sum()), "int", "What lifts the mean."),
        ("Mean order", f"=AVERAGE({O_})", f(O.mean()), N, "Report it only with the median beside it.")],
       [("o", "Rewrite 'The average stay at Mercyland is 5 days' honestly, using the median, IQR and the long-stay outliers from 4.1.11.",
         f"Half of the {len(L):,} admissions in 2023 stayed {L.median():.0f} days or less (IQR {lq1:.0f} to {lq3:.0f} days); "
         f"{int((L > lf).sum())} long stays lift the mean to {L.mean():.1f} days.")])
    m.finish()


# =====================================================================================================
def m42(b, D):
    CV = D.CV.Final_Score
    S = D.S.Annual_Salary_NGN
    m = ModuleSheet(b, "4.2 Distributions", "4.2  Distribution Analysis",
                    "Tutor walk-through after Module 4.2: histograms, the normal curve, Pareto, percentile bands, z-scores and anomalies",
                    "4.2.1 - 4.2.9")
    m.dataset([("tCV", "Sheet Data_CV (A1)", "Kingsway exam scores: the histogram and normal-curve examples."),
               ("tCustRev", "Sheet Data_CustRev (A1)", f"{len(D.CR):,} SuperMart Online buyers, net revenue largest first: the Pareto example."),
               ("tSal", "Sheet Data_Sal (A1)", "Salary: percentile bands."),
               ("tDash25 / tStores", "Sheets Data_Dash25, Data_Stores (A1)", "FY2025 store sales, profit and targets: z-scores."),
               ("tCensus23", "Sheet Data_Census (A1)", "Mercyland 2023 daily bed occupancy: anomaly detection."),
               ("tBins, tStoreZ", "On this sheet", "Histogram bins and the store z-score table: every cell is a formula.")])
    lo = [0] + list(range(55, 100, 5))
    hi = [54] + list(range(59, 100, 5))
    hi[-1] = 100
    labels = ["Below 55"] + [f"{a}-{b_}" for a, b_ in zip(lo[1:], hi[1:])]
    labels[-1] = "95-100"
    cnt = [int(((CV >= a) & (CV <= b_)).sum()) for a, b_ in zip(lo, hi)]
    T = {}

    def bins(mm):
        df = pd.DataFrame({"Bin": labels, "Lower": lo, "Upper": hi})
        calc = {"Students": ("=COUNTIFS(tCV[Final_Score],\">=\"&[@Lower],tCV[Final_Score],\"<=\"&[@Upper])", "int", cnt),
                "Share": ("=[@Students]/COUNT(tCV[Final_Score])", "pct", [c / len(CV) for c in cnt])}
        mm.h("tBins: exam scores in 5-point bins (the histogram's source)")
        T["b"] = mm.s.table(df, "tBins", fmts={"Lower": "int", "Upper": "int"}, calc=calc, widths={"Bin": 10, "Students": 10, "Share": 9})
        mm.s.r += 1
        mm.s.chart("column", [{"name": "Students", "categories": T["b"].ref("Bin"), "values": T["b"].ref("Students"), "gap": 10}],
                   f"Exam scores peak at {labels[int(np.argmax(cnt))]}, with tails on both sides", "Final score", "Students", legend=False,
                   rows=16, size=(640, 290))

    pk = int(np.argmax(cnt))
    tp(m, "4.2.1",
       [("Students in the peak bin", "=MAX(tBins[Students])", max(cnt), "int", ""),
        ("Peak bin", "=INDEX(tBins[Bin],MATCH(MAX(tBins[Students]),tBins[Students],0))", labels[pk], "gen", "The most common range of scores."),
        ("Bins add up to every student", "=SUM(tBins[Students])", len(CV), "int", "Check: equals COUNT(tCV[Final_Score]).")],
       [("n", "How many students scored 70-74?", '=COUNTIFS(tCV[Final_Score],">=70",tCV[Final_Score],"<=74")', int(((CV >= 70) & (CV <= 74)).sum()), "int",
         "Students scoring 70-74"),
        ("n", "What share of students scored below 55?", '=COUNTIF(tCV[Final_Score],"<55")/COUNT(tCV[Final_Score])', f((CV < 55).mean()), "pct",
         "Share below 55")],
       steps=["List the bins with a lower and an upper limit.", "COUNTIFS between the limits gives each bin's count.",
              "Plot the counts as columns with no gaps: that is a histogram."], extra=bins)

    mu, s_ = CV.mean(), sd(CV)
    w1 = ((CV >= mu - s_) & (CV <= mu + s_)).mean()
    w2 = ((CV >= mu - 2 * s_) & (CV <= mu + 2 * s_)).mean()
    sw1 = ((S >= S.mean() - sd(S)) & (S <= S.mean() + sd(S))).mean()
    W1 = f'=COUNTIFS({E_},">="&(AVERAGE({E_})-STDEV.S({E_})),{E_},"<="&(AVERAGE({E_})+STDEV.S({E_})))/COUNT({E_})'
    tp(m, "4.2.2",
       [("Exam scores within 1 SD of the mean", W1, f(w1), "pct", "A normal curve has 68.3%."),
        ("Chance of a score at or below 75, if normal", f"=NORM.DIST(75,AVERAGE({E_}),STDEV.S({E_}),TRUE)", f(stats.norm.cdf(75, mu, s_)), "pct",
         "Compare: actual share at or below 75 is below."),
        ("Actual share at or below 75", f'=COUNTIF({E_},"<=75")/COUNT({E_})', f((CV <= 75).mean()), "pct", "")],
       [("n", "What share of exam scores lie within 2 SD of the mean?", W1.replace("STDEV.S", "2*STDEV.S"), f(w2), "pct", "Within 2 SD"),
        ("n", "What share of salaries lie within 1 SD of the mean salary?",
         f'=COUNTIFS({S_},">="&(AVERAGE({S_})-STDEV.S({S_})),{S_},"<="&(AVERAGE({S_})+STDEV.S({S_})))/COUNT({S_})', f(sw1), "pct", "Salary within 1 SD")],
       note="If far more than 68% sit within 1 SD, the curve is more peaked than normal (positive kurtosis, 4.1.9).")

    tp(m, "4.2.3",
       [("Lowest score", f"=MIN({E_})", int(CV.min()), "int", ""), ("Highest score", f"=MAX({E_})", int(CV.max()), "int", ""),
        ("Bins needed at width 5", f"=ROUNDUP((MAX({E_})-MIN({E_})+1)/5,0)", math.ceil((CV.max() - CV.min() + 1) / 5), "int", "About 10-15 bins usually shows the shape.")],
       [("n", "How many bins are needed at a width of 2 points?", f"=ROUNDUP((MAX({E_})-MIN({E_})+1)/2,0)", math.ceil((CV.max() - CV.min() + 1) / 2), "int",
         "Bins at width 2"),
        ("t", "In tBins, which bin holds the most students?", labels[pk], "=INDEX(tBins[Bin],MATCH(MAX(tBins[Students]),tBins[Students],0))", "Peak bin")],
       do=[("Select Data_CV column J (Final_Score). Insert › Insert Statistic Chart › Histogram. Then Format Axis › Bin width 5.",
            f"A histogram peaking at {labels[pk]}, like the chart under 4.2.1."),
           ("Change the bin width to 2, then 20.", "Width 2 looks spiky (noise); width 20 hides the shape in a few blocks.")])

    cr = D.CR
    k80 = int((cr.Cum_Share < 0.8).sum()) + 1
    top20 = round(0.2 * len(cr))

    def pareto(mm):
        mm.s.chart("line", [{"name": "Cumulative share of revenue", "categories": dref("Data_CustRev", cr, "Rank"),
                             "values": dref("Data_CustRev", cr, "Cum_Share")}],
                   f"{k80:,} of {len(cr):,} buyers ({k80 / len(cr):.0%}) bring 80% of net revenue", "Customers, largest first", "Cumulative share",
                   legend=False, y_fmt="0%", rows=15, size=(640, 270),
                   extra=lambda ch: ch.set_x_axis({"name": "Customers, largest first", "interval_unit": 400, "num_font": {"size": 8}}))

    tp(m, "4.2.4",
       [("Net revenue, all buyers", "=SUM(tCustRev[Net_Revenue])", f(cr.Net_Revenue.sum()), N, "Delivered + Completed."),
        ("Check against tOrders", '=SUMIFS(tOrders[total_price],tOrders[order_status],"Delivered")+SUMIFS(tOrders[total_price],tOrders[order_status],"Completed")',
         f(cr.Net_Revenue.sum()), N, "Must equal the line above."),
        ("Customers needed for 80% of revenue", '=COUNTIF(tCustRev[Cum_Share],"<0.8")+1', k80, "int", ""),
        ("Share of buyers that is", '=(COUNTIF(tCustRev[Cum_Share],"<0.8")+1)/COUNT(tCustRev[Rank])', k80 / len(cr), "pct", "Not 20%: the 80/20 rule is a rough guide.")],
       [("n", "What share of net revenue comes from the top 20% of buyers?",
         '=SUMIF(tCustRev[Rank],"<="&ROUND(0.2*COUNT(tCustRev[Rank]),0),tCustRev[Net_Revenue])/SUM(tCustRev[Net_Revenue])',
         f(cr.Net_Revenue[:top20].sum() / cr.Net_Revenue.sum()), "pct", "Top-20% share"),
        ("n", "How many buyers bring the first 50% of revenue?", '=COUNTIF(tCustRev[Cum_Share],"<0.5")+1', int((cr.Cum_Share < .5).sum()) + 1, "int",
         "Buyers for 50%")],
       steps=["Total each customer's net revenue (a PivotTable on tOrders, filtered to Delivered and Completed).",
              "Sort largest first and add a running total; divide by the grand total for the cumulative share.",
              "Count the customers until the cumulative share reaches 80%."], extra=pareto)

    q1s, q2s, q3s = q(S, 1), q(S, 2), q(S, 3)
    tp(m, "4.2.5",
       [("Percentile rank of ₦12,000,000", f"=PERCENTRANK.INC({S_},12000000)", prank(S, 12e6), "pct", "Share of respondents below ₦12m."),
        ("Band 1 (up to Q1)", f'=COUNTIF({S_},"<="&QUARTILE.INC({S_},1))', int((S <= q1s).sum()), "int", ""),
        ("Band 2 (Q1 to median)", f'=COUNTIFS({S_},">"&QUARTILE.INC({S_},1),{S_},"<="&QUARTILE.INC({S_},2))', int(((S > q1s) & (S <= q2s)).sum()), "int",
         "Bands are not exactly 25% when many people share a salary.")],
       [("n", "What is the percentile rank of an exam score of 75?", f"=PERCENTRANK.INC({E_},75)", prank(CV, 75), "pct", "Percentile rank of 75"),
        ("n", "How many respondents are in Band 4 (salary above Q3)?", f'=COUNTIF({S_},">"&QUARTILE.INC({S_},3))', int((S > q3s).sum()), "int", "Band 4 count")])

    st = D.STO.copy()
    g = D.D25.groupby("Store_Name").agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"))
    st["Sales"] = st.Store_Name.map(g.Sales)
    st["Profit"] = st.Store_Name.map(g.Profit)
    st["Margin"] = st.Profit / st.Sales
    st["Attain"] = st.Sales / st.Annual_Target_2025

    def z(c):
        return (st[c] - st[c].mean()) / st[c].std(ddof=1)
    st["zS"], st["zM"], st["zA"] = z("Sales"), z("Margin"), z("Attain")
    st["Comp"] = (st.zS + st.zM + st.zA) / 3

    def ztable(mm):
        df = pd.DataFrame({"Store_Name": st.Store_Name})
        calc = {"Sales": ("=SUMIF(tDash25[Store_Name],[@Store_Name],tDash25[Sales])", N, [f(x) for x in st.Sales]),
                "Margin": ("=SUMIF(tDash25[Store_Name],[@Store_Name],tDash25[Profit])/[@Sales]", "pct", [f(x) for x in st.Margin]),
                "Attainment": ("=[@Sales]/XLOOKUP([@Store_Name],tStores[Store_Name],tStores[Annual_Target_2025])", "pct", [f(x) for x in st.Attain]),
                "z_Sales": ("=STANDARDIZE([@Sales],AVERAGE([Sales]),STDEV.S([Sales]))", "dec2", [f(x) for x in st.zS]),
                "z_Margin": ("=STANDARDIZE([@Margin],AVERAGE([Margin]),STDEV.S([Margin]))", "dec2", [f(x) for x in st.zM]),
                "z_Attain": ("=STANDARDIZE([@Attainment],AVERAGE([Attainment]),STDEV.S([Attainment]))", "dec2", [f(x) for x in st.zA]),
                "Composite": ("=AVERAGE([@z_Sales],[@z_Margin],[@z_Attain])", "dec2", [f(x) for x in st.Comp])}
        mm.h("tStoreZ: FY2025 store measures and their z-scores")
        T["z"] = mm.s.table(df, "tStoreZ", calc=calc, widths={"Store_Name": 24, "Sales": 14, "Margin": 9, "Attainment": 11, "z_Sales": 9,
                                                               "z_Margin": 9, "z_Attain": 9, "Composite": 10})
        mm.s.r += 1
        mm.s.chart("bar", [{"name": "z (sales)", "categories": T["z"].ref("Store_Name"), "values": T["z"].ref("z_Sales")}],
                   f"Only {st.Store_Name[st.zS.idxmax()]} is more than 2 SD above the average store", "", "z-score", legend=False, rows=16, size=(640, 290))

    top = st.loc[st.zS.idxmax()]
    tp(m, "4.2.6",
       [("Average store sales", "=AVERAGE(tStoreZ[Sales])", f(st.Sales.mean()), N, ""),
        ("SD of store sales", "=STDEV.S(tStoreZ[Sales])", f(st.Sales.std(ddof=1)), N, ""),
        (f"z of {top.Store_Name}", "=MAX(tStoreZ[z_Sales])", f(top.zS), "dec2", "(sales - mean) / SD: 2.6 SDs above the average store.")],
       [("n", "What is the sales z-score of Abuja - Wuse? (2 decimals)", '=XLOOKUP("Abuja - Wuse",tStoreZ[Store_Name],tStoreZ[z_Sales])',
         f(st.set_index("Store_Name").zS["Abuja - Wuse"]), "dec2", "z of Wuse"),
        ("n", "How many stores have a sales z-score above +1 or below -1?", '=COUNTIF(tStoreZ[z_Sales],">1")+COUNTIF(tStoreZ[z_Sales],"<-1")',
         int((st.zS.abs() > 1).sum()), "int", "Stores beyond 1 SD")],
       extra=ztable)

    best = st.loc[st.Comp.idxmax()]
    tp(m, "4.2.7",
       [("Best composite score", "=MAX(tStoreZ[Composite])", f(best.Comp), "dec2", "Average of three z-scores: one scale for three measures."),
        ("Best store overall", "=INDEX(tStoreZ[Store_Name],MATCH(MAX(tStoreZ[Composite]),tStoreZ[Composite],0))", best.Store_Name, "gen", "")],
       [("t", "Which store has the lowest composite score?", st.loc[st.Comp.idxmin()].Store_Name,
         "=INDEX(tStoreZ[Store_Name],MATCH(MIN(tStoreZ[Composite]),tStoreZ[Composite],0))", "Lowest composite"),
        ("n", "What is Jos - Rayfield's composite score? (2 decimals)", '=XLOOKUP("Jos - Rayfield",tStoreZ[Store_Name],tStoreZ[Composite])',
         f(st.set_index("Store_Name").Comp["Jos - Rayfield"]), "dec2", "Rayfield composite")])

    C = D.C23
    oc = C.Occupancy_Pct
    om, osd = oc.mean(), oc.std(ddof=1)
    hiN, loN = int((oc > om + 2 * osd).sum()), int((oc < om - 2 * osd).sum())
    O_C = "tCensus23[Occupancy_Pct]"
    tp(m, "4.2.8",
       [("Mean occupancy", f"=AVERAGE({O_C})", f(om), "pct", ""), ("SD of occupancy", f"=STDEV.S({O_C})", f(osd), "pct", ""),
        ("Days with z below -2", f'=COUNTIF({O_C},"<"&(AVERAGE({O_C})-2*STDEV.S({O_C})))', loN, "int", "Unusually empty."),
        ("Lowest day", f"=INDEX(tCensus23[Census_Date],MATCH(MIN({O_C}),{O_C},0))", f(int((C.Census_Date[oc.idxmin()] - pd.Timestamp('1899-12-30')).days)),
         "date", f"{oc.min():.1%} occupied.")],
       [("n", "How many days had occupancy more than 2 SD above the mean?", f'=COUNTIF({O_C},">"&(AVERAGE({O_C})+2*STDEV.S({O_C})))', hiN, "int",
         "High anomaly days"),
        ("n", "What is the z-score of the lowest day? (2 decimals)", f"=STANDARDIZE(MIN({O_C}),AVERAGE({O_C}),STDEV.S({O_C}))", f((oc.min() - om) / osd), "dec2",
         "z of the lowest day")])

    O = D.O.total_price
    tp(m, "4.2.9",
       [("Salary: mean ÷ median", f"=AVERAGE({S_})/MEDIAN({S_})", f(S.mean() / S.median()), "dec2", "Near 1: mean and SD are fine."),
        ("Order value: mean ÷ median", f"=AVERAGE({O_})/MEDIAN({O_})", f(O.mean() / O.median()), "dec2", "Far above 1: use median and IQR.")],
       [("t", "Length of stay has skewness 2.65. Which centre should you report: mean or median?", "median"),
        ("t", "Store sales, margin and attainment are in different units. Which score puts them on one scale?", "z-score", None, None,
         ("z-scores", "z", "z score", "standardise", "standardize"))])
    m.finish()


# =====================================================================================================
def m43(b, D):
    CV, H, S = D.CV, D.H, D.S
    y = CV.Final_Score
    m = ModuleSheet(b, "4.3 Correlation", "4.3  Correlation Analysis",
                    "Tutor walk-through after Module 4.3: CORREL, Spearman, the matrix, scatterplots, confounders and Simpson's paradox", "4.3.1 - 4.3.10")
    m.dataset([("tCV", "Sheet Data_CV (A1)", "Kingsway: exam score against attendance, study hours, previous score, tutoring, sleep and exercise."),
               ("tHotel", "Sheet Data_Hotel (A1)", "Atlantic Pearl: Days_Ahead and Is_Cancelled (1 = cancelled), for the causation and Simpson examples."),
               ("tSal", "Sheet Data_Sal (A1)", "Salary by gender and experience: the confounder example."),
               ("tStoreZ", "Sheet 4.2 Distributions", "12 stores: the Spearman rank example.")])
    DR = [("Lecture_Attendance_Pct", "Attendance %"), ("Study_Hours_Wk", "Hours studied"), ("Last_Sem_Score", "Previous score"),
          ("Tutorials_Attended", "Tutoring"), ("Exercise_Hrs_Wk", "Physical activity"), ("Sleep_Hrs_Night", "Sleep hours")]
    r = {c: float(np.corrcoef(CV[c], y)[0, 1]) for c, _ in DR}
    tp(m, "4.3.1",
       [(f"r: {lab}", f"=CORREL(tCV[{c}],tCV[Final_Score])", r[c], "dec3", "") for c, lab in DR[:4]],
       [("n", "What is r between physical activity (Exercise_Hrs_Wk) and exam score? (3 decimals)", "=CORREL(tCV[Exercise_Hrs_Wk],tCV[Final_Score])",
         r["Exercise_Hrs_Wk"], "dec3", "r exercise", 0.0006),
        ("t", "Is that link strong, moderate, weak or none?", "none", None, None, ("no link", "no relationship"))],
       note="Rough guide: |r| below 0.1 none; 0.1-0.3 weak; 0.3-0.5 moderate; above 0.5 strong.")

    tp(m, "4.3.2",
       [("CORREL: attendance and score", "=CORREL(tCV[Lecture_Attendance_Pct],tCV[Final_Score])", r["Lecture_Attendance_Pct"], "dec3", ""),
        ("PEARSON gives the same", "=PEARSON(tCV[Lecture_Attendance_Pct],tCV[Final_Score])", r["Lecture_Attendance_Pct"], "dec3", ""),
        ("r squared (shared variation)", "=RSQ(tCV[Final_Score],tCV[Lecture_Attendance_Pct])", r["Lecture_Attendance_Pct"] ** 2, "pct",
         "About a fifth of the variation in scores.")],
       [("n", "What is r squared for study hours and exam score? (as a %)", "=RSQ(tCV[Final_Score],tCV[Study_Hours_Wk])", r["Study_Hours_Wk"] ** 2, "pct", "r² hours"),
        ("n", "What is r between sleep and exam score? (3 decimals)", "=CORREL(tCV[Sleep_Hrs_Night],tCV[Final_Score])", r["Sleep_Hrs_Night"], "dec3", "r sleep", 0.0006)])

    st = D.STO.copy()
    g = D.D25.groupby("Store_Name").Sales.sum()
    st["Sales"] = st.Store_Name.map(g)
    rs = float(st.Sales.rank(ascending=False).corr(st.Floor_Area_sqm.rank(ascending=False)))
    rp = float(np.corrcoef(st.Sales, st.Floor_Area_sqm)[0, 1])

    def ranks(mm):
        df = pd.DataFrame({"Store_Name": st.Store_Name})
        calc = {"Sales": ("=XLOOKUP([@Store_Name],tStoreZ[Store_Name],tStoreZ[Sales])", N, [f(x) for x in st.Sales]),
                "Floor_Area": ("=XLOOKUP([@Store_Name],tStores[Store_Name],tStores[Floor_Area_sqm])", "int", [int(x) for x in st.Floor_Area_sqm]),
                "Rank_Sales": ("=RANK.AVG([@Sales],[Sales])", "dec1", [f(x) for x in st.Sales.rank(ascending=False)]),
                "Rank_Area": ("=RANK.AVG([@Floor_Area],[Floor_Area])", "dec1", [f(x) for x in st.Floor_Area_sqm.rank(ascending=False)])}
        mm.h("tRanks: store sales and floor area, with their ranks")
        mm.s.table(df, "tRanks", calc=calc, widths={"Store_Name": 24, "Sales": 14, "Floor_Area": 11, "Rank_Sales": 11, "Rank_Area": 11})
        mm.s.r += 1

    tp(m, "4.3.3",
       [("Pearson r: sales and floor area", "=CORREL(tRanks[Sales],tRanks[Floor_Area])", rp, "dec3", "Uses the values."),
        ("Spearman: CORREL of the ranks", "=CORREL(tRanks[Rank_Sales],tRanks[Rank_Area])", rs, "dec3", "Uses only the order: robust to outliers.")],
       [("n", "What rank (by sales) does Lagos - Lekki have?", '=XLOOKUP("Lagos - Lekki",tRanks[Store_Name],tRanks[Rank_Sales])',
         f(st.set_index("Store_Name").Sales.rank(ascending=False)["Lagos - Lekki"]), "dec1", "Lekki sales rank"),
        ("t", "Spearman is the Pearson correlation of what?", "ranks", None, None, ("the ranks", "rank"))],
       steps=["Add a rank column for each variable with RANK.AVG (ties get the average rank).", "CORREL the two rank columns: that is Spearman's rho."],
       note="On tCV (5,840 students) Spearman for attendance is 0.579 against Pearson 0.470: the link is strong but curved.", extra=ranks)

    M = ["Lecture_Attendance_Pct", "Study_Hours_Wk", "Last_Sem_Score"]
    pairs = [("Final_Score", c) for c in M] + [(M[0], M[1]), (M[0], M[2]), (M[1], M[2])]
    tp(m, "4.3.4",
       [(f"{a} with {c}", f"=CORREL(tCV[{a}],tCV[{c}])", float(np.corrcoef(CV[a], CV[c])[0, 1]), "dec3", "") for a, c in pairs],
       [("t", "Which driver is most strongly correlated with Final_Score?", "Lecture_Attendance_Pct", None, None, ("attendance", "lecture attendance")),
        ("n", "What is r between attendance and study hours? (3 decimals)", "=CORREL(tCV[Lecture_Attendance_Pct],tCV[Study_Hours_Wk])",
         float(np.corrcoef(CV[M[0]], CV[M[1]])[0, 1]), "dec3", "r attendance-hours", 0.0006)],
       note="The drivers are nearly uncorrelated with each other, so each adds its own information to a regression (4.4.10).")

    tp(m, "4.3.5", None, [],
       steps=["Data › Data Analysis › Correlation.", "Input range: the adjacent columns (with headers); tick Labels in first row.",
              "Output range: an empty area; then add a colour scale (Home › Conditional Formatting)."],
       do=[("Run Data Analysis › Correlation on Data_CV columns D:J (Study_Hours_Wk to Final_Score, with headers).",
            "A 7 x 7 lower-triangle matrix; the Final_Score row matches the CORREL values in 4.3.1."),
           ("Add a 3-colour scale to the matrix.", "Attendance and hours stand out; the rest are pale.")])

    def scat(mm):
        mm.s.chart("scatter", [{"name": "Students", "categories": dref("Data_CV", CV, "Study_Hours_Wk"), "values": dref("Data_CV", CV, "Final_Score"),
                                "trendline": {"type": "linear"}}],
                   f"More study hours go with higher scores, but loosely (r = {r['Study_Hours_Wk']:.2f})", "Hours studied a week", "Final score",
                   legend=False, rows=16, size=(640, 290),
                   extra=lambda ch: ch.set_y_axis({"name": "Final score", "min": 30, "max": 100, "num_font": {"size": 8}}))

    tp(m, "4.3.6",
       [("r, all students", "=CORREL(tCV[Study_Hours_Wk],tCV[Final_Score])", r["Study_Hours_Wk"], "dec3", "Read the scatter below before trusting it."),
        ("Scores of 80 or more", '=COUNTIF(tCV[Final_Score],">=80")', int((y >= 80).sum()), "int", "A few high scorers sit above the cloud."),
        ("Fewest / most hours a week", '=MIN(tCV[Study_Hours_Wk])&" to "&MAX(tCV[Study_Hours_Wk])', f"{int(CV.Study_Hours_Wk.min())} to {int(CV.Study_Hours_Wk.max())}",
         "gen", "")],
       [("n", "How many students study more than 30 hours a week?", '=COUNTIF(tCV[Study_Hours_Wk],">30")', int((CV.Study_Hours_Wk > 30).sum()), "int",
         "Students over 30 hours"),
        ("o", "Describe the scatter in one sentence: direction, strength and anything odd.",
         "Upward and loose: more hours go with higher scores on average, with wide scatter at every level and a few high scorers above the cloud.")],
       extra=scat)

    LB = [(0, 7, "0-7"), (8, 30, "8-30"), (31, 90, "31-90"), (91, 180, "91-180"), (181, 9999, "181+")]
    rate = {lab: float(H.Is_Cancelled[(H.Days_Ahead >= a) & (H.Days_Ahead <= b_)].mean()) for a, b_, lab in LB}
    nn = {lab: int(((H.Days_Ahead >= a) & (H.Days_Ahead <= b_)).sum()) for a, b_, lab in LB}
    rl = float(np.corrcoef(H.Days_Ahead, H.Is_Cancelled)[0, 1])

    def lead(mm):
        df = pd.DataFrame({"Days_Ahead": [x[2] for x in LB], "From": [x[0] for x in LB], "To": [x[1] for x in LB]})
        calc = {"Bookings": ("=COUNTIFS(tHotel[Days_Ahead],\">=\"&[@From],tHotel[Days_Ahead],\"<=\"&[@To])", "int", [nn[x[2]] for x in LB]),
                "Cancel_Rate": ("=AVERAGEIFS(tHotel[Is_Cancelled],tHotel[Days_Ahead],\">=\"&[@From],tHotel[Days_Ahead],\"<=\"&[@To])", "pct",
                                [rate[x[2]] for x in LB])}
        mm.h("tLead: cancellation rate by days booked ahead")
        t = mm.s.table(df, "tLead", fmts={"From": "int", "To": "int"}, calc=calc, widths={"Days_Ahead": 11, "Bookings": 10, "Cancel_Rate": 12})
        mm.s.r += 1
        mm.s.chart("column", [{"name": "Cancel rate", "categories": t.ref("Days_Ahead"), "values": t.ref("Cancel_Rate"), "labels": True, "label_fmt": "0%"}],
                   f"Bookings made 6+ months ahead cancel {rate['181+']:.0%} of the time", "Days booked ahead", "", legend=False, y_fmt="0%", rows=15,
                   size=(560, 260))

    tp(m, "4.3.7",
       [("r: days ahead and cancelled", "=CORREL(tHotel[Days_Ahead],tHotel[Is_Cancelled])", rl, "dec3", "Moderate positive link."),
        ("Cancel rate, booked 0-7 days ahead", "=XLOOKUP(\"0-7\",tLead[Days_Ahead],tLead[Cancel_Rate])", rate["0-7"], "pct", "")],
       [("n", "What is the cancellation rate for bookings made 181+ days ahead?", '=AVERAGEIFS(tHotel[Is_Cancelled],tHotel[Days_Ahead],">=181")',
         rate["181+"], "pct", "Cancel rate 181+"),
        ("o", "Does booking early cause cancellation? Give one other explanation.",
         "No: early bookers simply have more time for plans to change, and OTA bookings (which cancel more) are made earlier.")],
       extra=lead)

    gF = S[S.Gender == "Female"].Annual_Salary_NGN.mean() - S[S.Gender == "Male"].Annual_Salary_NGN.mean()
    gX = S[S.Gender == "Female"].Years_of_Experience.mean() - S[S.Gender == "Male"].Years_of_Experience.mean()
    band = S[(S.Years_of_Experience >= 6) & (S.Years_of_Experience <= 10)]
    gB = band[band.Gender == "Female"].Annual_Salary_NGN.mean() - band[band.Gender == "Male"].Annual_Salary_NGN.mean()
    AV = "AVERAGEIFS(tSal[{c}],tSal[Gender],\"{g}\"{x})"
    xb = ',tSal[Years_of_Experience],">=6",tSal[Years_of_Experience],"<=10"'
    tp(m, "4.3.8",
       [("Raw gap: female minus male average salary", "=" + AV.format(c="Annual_Salary_NGN", g="Female", x="") + "-" + AV.format(c="Annual_Salary_NGN", g="Male", x=""),
         f(gF), N, "The headline gap."),
        ("Experience gap (years)", "=" + AV.format(c="Years_of_Experience", g="Female", x="") + "-" + AV.format(c="Years_of_Experience", g="Male", x=""),
         f(gX), "dec2", "Women in the survey have less experience: a confounder."),
        ("Gap at the same experience (6-10 years)", "=" + AV.format(c="Annual_Salary_NGN", g="Female", x=xb) + "-" + AV.format(c="Annual_Salary_NGN", g="Male", x=xb),
         f(gB), N, "Compare like with like.")],
       [("n", "What is the female minus male gap for 0-2 years of experience?",
         "=" + AV.format(c="Annual_Salary_NGN", g="Female", x=xb.replace(">=6", ">=0").replace("<=10", "<=2")) + "-"
         + AV.format(c="Annual_Salary_NGN", g="Male", x=xb.replace(">=6", ">=0").replace("<=10", "<=2")),
         f(S[(S.Years_of_Experience <= 2) & (S.Gender == "Female")].Annual_Salary_NGN.mean() - S[(S.Years_of_Experience <= 2) & (S.Gender == "Male")].Annual_Salary_NGN.mean()),
         N, "Gap at 0-2 years")])

    def rate2(y_=None, room=None):
        h = H
        if y_:
            h = h[h.Check_In_Year == y_]
        if room:
            h = h[h.Room_Category == room]
        return float(h.Is_Cancelled.mean())
    SP = "AVERAGEIFS(tHotel[Is_Cancelled],tHotel[Room_Category],\"{r}\"{y})"
    tp(m, "4.3.9",
       [("Deluxe, 2023 arrivals", "=" + SP.format(r="Deluxe", y=",tHotel[Check_In_Year],2023"), rate2(2023, "Deluxe"), "pct", "Deluxe cancels less..."),
        ("Standard Twin, 2023 arrivals", "=" + SP.format(r="Standard Twin", y=",tHotel[Check_In_Year],2023"), rate2(2023, "Standard Twin"), "pct", ""),
        ("Deluxe, both years", "=" + SP.format(r="Deluxe", y=""), rate2(None, "Deluxe"), "pct", "...but more overall."),
        ("Standard Twin, both years", "=" + SP.format(r="Standard Twin", y=""), rate2(None, "Standard Twin"), "pct", "")],
       [("n", "What is the Standard Twin cancellation rate for 2024 arrivals?", "=" + SP.format(r="Standard Twin", y=",tHotel[Check_In_Year],2024"),
         rate2(2024, "Standard Twin"), "pct", "Twin 2024"),
        ("n", "What share of Deluxe bookings are 2024 arrivals?", '=COUNTIFS(tHotel[Room_Category],"Deluxe",tHotel[Check_In_Year],2024)/COUNTIF(tHotel[Room_Category],"Deluxe")',
         float((H[H.Room_Category == "Deluxe"].Check_In_Year == 2024).mean()), "pct", "Deluxe share in 2024"),
        ("o", "Explain the reversal in one sentence.",
         "Most Deluxe bookings fall in 2024, when every room type cancelled more, so the combined Deluxe rate is pulled up by the year mix.")])

    tp(m, "4.3.10",
       [("r, attendance and score", "=ROUND(CORREL(tCV[Lecture_Attendance_Pct],tCV[Final_Score]),2)", round(r["Lecture_Attendance_Pct"], 2), "dec2", "Report 2 decimals."),
        ("Students (n)", "=COUNT(tCV[Final_Score])", len(y), "int", "Always give n.")],
       [("o", "Write one responsible sentence about study hours and exam score.",
         f"Across {len(y):,} Kingsway students in 2024/25, study hours and exam score are moderately associated (r = {r['Study_Hours_Wk']:.2f}); this shows association, not cause.")])
    m.finish()


# =====================================================================================================
def ols(y, X):
    Xm = np.column_stack([X, np.ones(len(y))])
    beta, *_ = np.linalg.lstsq(Xm, y, rcond=None)
    res = y - Xm @ beta
    n, k = Xm.shape
    df = n - k
    sse = float((res ** 2).sum())
    se = np.sqrt(np.diag(sse / df * np.linalg.inv(Xm.T @ Xm)))
    r2 = 1 - sse / float(((y - y.mean()) ** 2).sum())
    return dict(b=beta, se=se, df=df, r2=r2, sey=math.sqrt(sse / df), res=res)


def m44(b, D):
    S = D.S
    y = S.Annual_Salary_NGN.values.astype(float)
    x = S.Years_of_Experience.values.astype(float)
    m = ModuleSheet(b, "4.4 Regression", "4.4  Regression Analysis",
                    "Tutor walk-through after Module 4.4: SLOPE, INTERCEPT, RSQ, LINEST, p-values, residuals and prediction", "4.4.1 - 4.4.12")
    m.dataset([("tSal", "Sheet Data_Sal (A1)", "Salary (y) against Years_of_Experience, Age, Is_Masters and Is_PhD (x). The four inputs sit side by side "
                "(columns G:J), so LINEST can take tSal[[Age]:[Is_PhD]].")])
    sl, ic = np.polyfit(x, y, 1)
    r2s = float(np.corrcoef(x, y)[0, 1] ** 2)
    C4 = ["Age", "Years_of_Experience", "Is_Masters", "Is_PhD"]
    m4 = ols(y, S[C4].values.astype(float))
    m3 = ols(y, S[C4[1:]].values.astype(float))
    L4 = "LINEST(tSal[Annual_Salary_NGN],tSal[[Age]:[Is_PhD]],TRUE,TRUE)"
    L3 = "LINEST(tSal[Annual_Salary_NGN],tSal[[Years_of_Experience]:[Is_PhD]],TRUE,TRUE)"
    col = {"Is_PhD": 1, "Is_Masters": 2, "Years_of_Experience": 3, "Age": 4, "Intercept": 5}   # LINEST is right-to-left
    idx = {"Age": 0, "Years_of_Experience": 1, "Is_Masters": 2, "Is_PhD": 3, "Intercept": 4}

    def co(k):
        return f"INDEX({L4},1,{col[k]})"

    def pv(k):
        return f"T.DIST.2T(ABS(INDEX({L4},1,{col[k]})/INDEX({L4},2,{col[k]})),INDEX({L4},4,2))"

    def pval(k):
        i = idx[k]
        return float(2 * stats.t.sf(abs(m4["b"][i] / m4["se"][i]), m4["df"]))

    SY = "tSal[Annual_Salary_NGN],tSal[Years_of_Experience]"
    tp(m, "4.4.1",
       [("Extra salary per year of experience (slope)", f"=SLOPE({SY})", f(sl), N, "The answer to the question, on average.")],
       [("n", "What salary does the line give at 0 years (the intercept)?", f"=INTERCEPT({SY})", f(ic), N, "Intercept")],
       steps=["Name the outcome (y): Annual_Salary_NGN.", "Name the input (x): Years_of_Experience.", "Fit the straight line y = a + b x and read b."])

    def chart(mm):
        mm.s.chart("scatter", [{"name": "Respondents", "categories": dref("Data_Sal", S, "Years_of_Experience"), "values": dref("Data_Sal", S, "Annual_Salary_NGN"),
                                "trendline": {"type": "linear", "display_equation": True, "display_r_squared": True}}],
                   f"Each extra year goes with about ₦{sl / 1e3:,.0f}k more salary (R² {r2s:.2f})", "Years of experience", "Annual salary (₦)",
                   legend=False, y_fmt='₦#,##0.0,,"m"', rows=16, size=(640, 300))

    tp(m, "4.4.2",
       [("Slope", f"=SLOPE({SY})", f(sl), N, ""), ("Intercept", f"=INTERCEPT({SY})", f(ic), N, ""),
        ("R-squared", f"=RSQ({SY})", r2s, "dec3", "Share of salary variation explained by experience.")],
       [("n", "Predict the salary for 5 years of experience with FORECAST.LINEAR.", f"=FORECAST.LINEAR(5,{SY})", f(ic + sl * 5), N, "Prediction at 5 years"),
        ("n", "What is the standard error of the simple model (STEYX)?", f"=STEYX({SY})", f(math.sqrt(((y - ic - sl * x) ** 2).sum() / (len(y) - 2))), N, "STEYX")],
       do=[("Right-click any point on the scatter below › Add Trendline › Linear; tick Display Equation and Display R-squared.",
            f"y = {sl:,.0f}x + {ic:,.0f}; R² = {r2s:.4f} (the same as SLOPE, INTERCEPT and RSQ).")], extra=chart)

    tp(m, "4.4.3",
       [(f"Coefficient: {k.replace('_', ' ')}", f"={co(k)}", f(m4["b"][idx[k]]), N, "") for k in ["Intercept", "Years_of_Experience", "Age", "Is_Masters", "Is_PhD"]],
       [("n", "What is the p-value of Age? (3 decimals)", f"={pv('Age')}", pval("Age"), "dec3", "p-value of Age", 0.0006),
        ("n", "What extra salary goes with a PhD (Is_PhD coefficient)?", f"={co('Is_PhD')}", f(m4["b"][3]), N, "PhD coefficient")],
       steps=["Put the input columns side by side (here Age to Is_PhD, columns G:J).", "=LINEST(y, x-columns, TRUE, TRUE) returns a 5-row block.",
              "Read it with INDEX(LINEST(...), row, column). Columns run right to left: the last input first, the intercept last."],
       note="Row 1 = coefficients, row 2 = standard errors, row 3 = R-squared and SE of the estimate, row 4 = F and residual degrees of freedom.")

    prof = dict(Age=35, Years_of_Experience=10, Is_Masters=1, Is_PhD=0)
    pred = m4["b"][4] + sum(m4["b"][idx[k]] * v for k, v in prof.items())
    PF = f"={co('Intercept')}+{co('Years_of_Experience')}*10+{co('Age')}*35+{co('Is_Masters')}*1+{co('Is_PhD')}*0"
    tp(m, "4.4.4",
       [("Simple: salary at 10 years", f"=INTERCEPT({SY})+SLOPE({SY})*10", f(ic + sl * 10), N, "a + b x"),
        ("Multiple: 10 years, age 35, Master's", PF, f(pred), N, "Intercept + each coefficient x its value.")],
       [("n", "Multiple model: predict for 20 years, age 45, PhD (Is_Masters 0, Is_PhD 1).",
         f"={co('Intercept')}+{co('Years_of_Experience')}*20+{co('Age')}*45+{co('Is_PhD')}*1",
         f(m4["b"][4] + m4["b"][1] * 20 + m4["b"][0] * 45 + m4["b"][3]), N, "Prediction 20y PhD")])

    tp(m, "4.4.5",
       [("R-squared, experience only", f"=RSQ({SY})", r2s, "pct", ""), ("R-squared, four inputs", f"=INDEX({L4},3,1)", m4["r2"], "pct", "")],
       [("n", "What share of salary variation does the four-input model leave unexplained?", f"=1-INDEX({L4},3,1)", 1 - m4["r2"], "pct", "Unexplained share")])

    n_ = len(y)
    adj = lambda r2, k: 1 - (1 - r2) * (n_ - 1) / (n_ - k - 1)  # noqa: E731
    tp(m, "4.4.6",
       [("R-squared, 3 inputs (no Age)", f"=INDEX({L3},3,1)", m3["r2"], "dec4", ""),
        ("Adjusted R-squared, 3 inputs", f"=1-(1-INDEX({L3},3,1))*(COUNT(tSal[Annual_Salary_NGN])-1)/(COUNT(tSal[Annual_Salary_NGN])-3-1)", adj(m3["r2"], 3), "dec4", ""),
        ("Adjusted R-squared, 4 inputs", f"=1-(1-INDEX({L4},3,1))*(COUNT(tSal[Annual_Salary_NGN])-1)/(COUNT(tSal[Annual_Salary_NGN])-4-1)", adj(m4["r2"], 4), "dec4",
         "Barely higher: Age adds almost nothing.")],
       [("n", "By how much does adding Age raise the adjusted R-squared? (4 decimals)",
         f"=(1-(1-INDEX({L4},3,1))*(COUNT(tSal[Annual_Salary_NGN])-1)/(COUNT(tSal[Annual_Salary_NGN])-5))-(1-(1-INDEX({L3},3,1))*(COUNT(tSal[Annual_Salary_NGN])-1)/(COUNT(tSal[Annual_Salary_NGN])-4))",
         adj(m4["r2"], 4) - adj(m3["r2"], 3), "dec4", "Gain from Age", 0.00006)],
       note="Adjusted R-squared = 1 - (1 - R²) x (n - 1) / (n - k - 1), where k is the number of inputs.")

    tp(m, "4.4.7",
       [(f"p-value: {k.replace('_', ' ')}", f"={pv(k)}", pval(k), "dec4", "Below 0.05: significant." if pval(k) < .05 else "Above 0.05: not significant.")
        for k in ["Years_of_Experience", "Age", "Is_Masters"]],
       [("n", "What is the p-value of Is_PhD? (4 decimals)", f"={pv('Is_PhD')}", pval("Is_PhD"), "dec4", "p-value of PhD", 0.00006),
        ("t", "Is Age significant at the 5% level? (Yes or No)", "No")],
       note="p = T.DIST.2T(|coefficient / standard error|, residual degrees of freedom).")

    RB = [(0, 2, "0-2"), (3, 5, "3-5"), (6, 10, "6-10"), (11, 15, "11-15"), (16, 20, "16-20"), (21, 99, "21+")]
    res = y - (ic + sl * x)
    rb = {lab: float(res[(x >= a) & (x <= b_)].mean()) for a, b_, lab in RB}

    def resid(mm):
        df = pd.DataFrame({"Years": [t[2] for t in RB], "From": [t[0] for t in RB], "To": [t[1] for t in RB]})
        fx = (f"=SUMPRODUCT((tSal[Years_of_Experience]>=[@From])*(tSal[Years_of_Experience]<=[@To])*(tSal[Annual_Salary_NGN]-INTERCEPT({SY})-SLOPE({SY})*tSal[Years_of_Experience]))"
              "/COUNTIFS(tSal[Years_of_Experience],\">=\"&[@From],tSal[Years_of_Experience],\"<=\"&[@To])")
        mm.h("tResid: average residual (actual - predicted) by experience band, simple model")
        t = mm.s.table(df, "tResid", fmts={"From": "int", "To": "int"}, calc={"Avg_Residual": (fx, N, [rb[t_[2]] for t_ in RB])},
                       widths={"Years": 8, "Avg_Residual": 14})
        mm.s.r += 1
        mm.s.chart("column", [{"name": "Average residual", "categories": t.ref("Years"), "values": t.ref("Avg_Residual")}],
                   "Residuals curve: the line over-predicts at both ends of experience", "Years of experience", "Avg residual (₦)", legend=False,
                   y_fmt='₦#,##0.0,,"m"', rows=15, size=(560, 260))

    tp(m, "4.4.8",
       [("Average residual, 0-2 years", "=INDEX(tResid[Avg_Residual],1)", rb["0-2"], N, "Negative: the line predicts too high."),
        ("Average residual, all respondents", f"=SUMPRODUCT(tSal[Annual_Salary_NGN]-INTERCEPT({SY})-SLOPE({SY})*tSal[Years_of_Experience])/COUNT(tSal[Annual_Salary_NGN])",
         0.0, N, "Always 0 overall: look at the pattern by band.")],
       [("n", "What is the average residual for 21+ years?", "=INDEX(tResid[Avg_Residual],6)", rb["21+"], N, "Residual 21+"),
        ("t", "Is the residual pattern random or curved?", "curved", None, None, ("curve", "pattern", "not random"))],
       extra=resid)

    pi = 1.96 * m4["sey"]
    tp(m, "4.4.9",
       [("Point prediction (4.4.4 profile)", PF, f(pred), N, "10 years, age 35, Master's."),
        ("Standard error of the estimate", f"=INDEX({L4},3,2)", m4["sey"], N, "Typical size of a residual."),
        ("Lower end of the rough 95% range", PF + f"-1.96*INDEX({L4},3,2)", f(pred - pi), N, ""),
        ("Upper end", PF + f"+1.96*INDEX({L4},3,2)", f(pred + pi), N, "Individual salaries vary widely around the line.")],
       [("n", "How wide is the rough 95% range (upper minus lower)?", f"=2*1.96*INDEX({L4},3,2)", f(2 * pi), N, "Width of the range")])

    rae = float(np.corrcoef(S.Age, S.Years_of_Experience)[0, 1])
    tp(m, "4.4.10",
       [("r (Age, Years_of_Experience)", "=CORREL(tSal[Age],tSal[Years_of_Experience])", rae, "dec3", "The two inputs move almost together."),
        ("VIF = 1 / (1 - r squared)", "=1/(1-CORREL(tSal[Age],tSal[Years_of_Experience])^2)", 1 / (1 - rae ** 2), "dec1", "Above about 5: a warning.")],
       [("n", "What is the standard error of the Age coefficient?", f"=INDEX({L4},2,4)", f(m4["se"][0]), N, "SE of Age"),
        ("t", "Which input would you drop from the four-input model?", "Age")])

    tp(m, "4.4.11",
       [("Fewest years in the data", "=MIN(tSal[Years_of_Experience])", f(x.min()), "dec1", ""), ("Most years", "=MAX(tSal[Years_of_Experience])", f(x.max()), "dec1",
                                                                                                    "Do not predict beyond this range.")],
       [("t", "Can the model predict a salary for 45 years of experience? (Yes or No)", "No"),
        ("o", "Name one input missing from the model that probably affects salary.", "Job family, city or employer size: none is in the model.")])

    tp(m, "4.4.12",
       [("Years coefficient", f"={co('Years_of_Experience')}", f(m4["b"][1]), N, ""), ("PhD coefficient", f"={co('Is_PhD')}", f(m4["b"][3]), N, ""),
        ("Adjusted R-squared", f"=1-(1-INDEX({L4},3,1))*(COUNT(tSal[Annual_Salary_NGN])-1)/(COUNT(tSal[Annual_Salary_NGN])-5)", adj(m4["r2"], 4), "pct", "")],
       [("o", "Write the three-line summary (finding, fit, limits) for HR.",
         f"Each year of experience goes with about ₦{m4['b'][1] / 1e6:.2f}m more salary and a PhD with about ₦{m4['b'][3] / 1e6:.2f}m. "
         f"The model explains {adj(m4['r2'], 4):.0%} of the variation. Limits: {n_:,} survey respondents, association only, curved at the extremes.")])
    m.finish()
