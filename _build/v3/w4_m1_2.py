"""Week 4 edition-3 slide copy: Modules 4.1 (descriptive statistics, Analysis ToolPak) and 4.2 (distributions, Pareto, percentiles,
z-scores). Data: tSal (salary survey), tOrders (SuperMart Online), tCV (Kingsway), tLOS / tCensus23 (Mercyland 2023), tDash25 +
tStores (SuperMart FY2025). Every number is computed here from w4_data."""
import numpy as np
import pandas as pd
from scipy.stats import norm

from v3kit import T
from w4_fx import (C23, CV, D25, EX, I, LOS, LS, N, Nm, OV, ORD, P, SAL, STO, Y, F, F2, kurt, pct_rank_inc, q, skew)

# ---------------------------------------------------------------- 4.1 figures
YM, YMED, YSD = Y.mean(), Y.median(), Y.std()
YMODE = float(Y.mode().iloc[0])
YMODE_N = int((Y == YMODE).sum())
YQ1, YQ3 = q(Y, .25), q(Y, .75)
OM, OMED, OSD = OV.mean(), OV.median(), OV.std()
TOP10 = OV.drop(OV.nlargest(10).index)
OQ1, OQ3 = q(OV, .25), q(OV, .75)
OFENCE = OQ3 + 1.5 * (OQ3 - OQ1)
O_OUT = int((OV > OFENCE).sum())
O_OUT_SHARE = OV[OV > OFENCE].sum() / OV.sum()
QMODE = int(ORD.quantity.mode().iloc[0])
QMODE_N = int((ORD.quantity == QMODE).sum())
SEN = SAL.Seniority.value_counts()
LMODE = int(LS.mode().iloc[0])
CVS = [("Online order value", OM, OSD, "₦"), ("Length of stay 2023", LS.mean(), LS.std(), "days"),
       ("Annual salary", YM, YSD, "₦"), ("Exam score", EX.mean(), EX.std(), "pts")]
CVS.sort(key=lambda r: -r[2] / r[1])
PCTS = [(p, q(Y, p / 100)) for p in (10, 25, 50, 75, 90)]


def _m(v, unit):
    return N(v) if unit == "₦" else (f"{v:.1f} days" if unit == "days" else f"{v:.1f}")


# ---------------------------------------------------------------- 4.2 figures
EM, ESD = EX.mean(), EX.std()
W1 = ((EX - EM).abs() <= ESD).mean()
W2 = ((EX - EM).abs() <= 2 * ESD).mean()
SW1 = ((Y - YM).abs() <= YSD).mean()
EBINS = pd.cut(EX, [0, 55, 60, 65, 70, 75, 80, 101], right=False,
               labels=["Below 55", "55-59", "60-64", "65-69", "70-74", "75-79", "80+"]).value_counts().sort_index()
EPEAK = EBINS.idxmax()
NT = ORD[ORD.order_status.isin(["Delivered", "Completed"])]
CUST = NT.groupby("customer_id").total_price.sum().sort_values(ascending=False)
CUM = CUST.cumsum() / CUST.sum()
K80 = int((CUM < 0.8).sum() + 1)
TOP20 = CUST.head(int(len(CUST) * 0.2)).sum() / CUST.sum()
ST = (D25.groupby(["Store_ID", "Store_Name"]).agg(Sales=("Sales", "sum"), Profit=("Profit", "sum")).reset_index()
      .merge(STO[["Store_ID", "Annual_Target_2025"]], on="Store_ID"))
ST["Margin"] = ST.Profit / ST.Sales
ST["Attain"] = ST.Sales / ST.Annual_Target_2025
for _c in ("Sales", "Margin", "Attain"):
    ST["z_" + _c] = (ST[_c] - ST[_c].mean()) / ST[_c].std()
ST["Comp"] = ST[["z_Sales", "z_Margin", "z_Attain"]].mean(axis=1)
STZ = ST.sort_values("Comp", ascending=False).reset_index(drop=True)
OCC = C23.Occupancy_Pct
OZ = (OCC - OCC.mean()) / OCC.std()
ANOM = C23[OZ.abs() > 2]
_nlo, _nhi = int((OZ < -2).sum()), int((OZ > 2).sum())
ANOM_TXT = ("all unusually low" if not _nhi else "all unusually high" if not _nlo else f"{_nlo} low, {_nhi} high")


def SKR(v):
    return "Strong right tail" if v > 1 else "Moderate right tail" if v > 0.5 else "Left tail" if v < -0.5 else "Roughly balanced"


def KUR(v):
    return "Very heavy tails" if v > 3 else "Heavier tails than normal" if v > 1 else "Flatter than normal" if v < -1 else "Close to normal"

LOWDAY = C23.loc[OZ.idxmin()]
BAND = pd.cut(Y, [0, YQ1, YMED, YQ3, 1e12], labels=["Band 1 (lowest 25%)", "Band 2", "Band 3", "Band 4 (top 25%)"]).value_counts().sort_index()


def sname(s):
    return s.split(" - ")[1] if " - " in s else s


M41 = dict(
    id="4.1", title="Descriptive Statistics and the Analysis ToolPak",
    subtitle="Centre, spread and shape of a column, the ToolPak summary, and how to report the figures honestly.",
    objectives=["Pick the right centre (mean, median, mode) for the data's shape.",
                "Measure spread with range, IQR, standard deviation and the coefficient of variation.",
                "Read skewness and kurtosis, percentiles and quartiles.",
                "Flag outliers with the 1.5 x IQR rule.",
                "Run the Analysis ToolPak's Descriptive Statistics and report the results honestly."],
    topics=[
        T("4.1.1", "What Descriptive Statistics Do",
          defn="Descriptive statistics summarise a column in a few numbers: its centre (a typical value), its spread (how far values "
               "vary) and its shape (balanced or lopsided).",
          simple="They describe what is in the data; they do not explain why.",
          grid=dict(label=f"Annual salary, ETEH Talent survey (tSal, {I(len(Y))} people)", widths=[1, 1.6, 2], rows=[
              ["Family", "Measure", "Value"],
              ["Centre", "Median", N(YMED)],
              ["Spread", "Middle 50% (Q1 to Q3)", f"{N(YQ1)} to {N(YQ3)}"],
              ["Shape", "Skewness", f"{skew(Y):.2f} (balanced)"],
              ["Size", "Count", f"{I(len(Y))} respondents"]]),
          watch="One average on its own is not an analysis; always give centre, spread and shape together."),

        T("4.1.2", "Mean",
          defn="The mean (AVERAGE) is the sum of the values divided by their count.",
          simple="It uses every value, so a few very large values pull it up. Mean x count = total, which is why finance uses it.",
          grid=dict(label="Mean and median side by side", widths=[2.2, 1.3, 1.3], rows=[
              ["Data", "Mean", "Median"],
              [f"Salary (tSal, {I(len(Y))})", N(YM), N(YMED)],
              [f"Order value (tOrders, {I(len(OV))})", N(OM), N(OMED)],
              ["Order value without the 10 largest", N(TOP10.mean()), N(TOP10.median())]]),
          fx=[("=AVERAGE(tSal[Annual_Salary_NGN])", N(YM), "Mean salary")], fx_widths=[2.6, 1.1, 1.1],
          watch=f"For order values the mean is {OM / OMED:.1f} times the median; quoting it as 'typical' misleads."),

        T("4.1.3", "Median",
          defn="The median (MEDIAN) is the middle value after sorting: half the values are below it and half above.",
          simple="Extreme values do not move it, so use it for lopsided data such as prices, incomes and waiting times.",
          grid=dict(label="A typical SuperMart Online order (tOrders)", widths=[2.4, 1.4], rows=[
              ["Measure", "Order value"],
              ["Median (typical order)", N(OMED)],
              ["Mean", N(OM)],
              ["Mean / median", f"{OM / OMED:.2f}"]]),
          fx=[("=MEDIAN(tOrders[total_price])", N(OMED), "Middle order value")], fx_widths=[2.4, 1, 1.2],
          watch="Rule of thumb: when mean / median is above about 1.2, report the median as the typical value."),

        T("4.1.4", "Mode",
          defn="The mode is the most common value: MODE.SNGL for numbers, a count per category for text.",
          simple="It is the only centre that works for categories, such as the commonest seniority or basket size.",
          grid=dict(label="Most common value in four columns", widths=[2.2, 1.4, 1.6], rows=[
              ["Column", "Mode", "How often"],
              ["Salary (tSal)", N(YMODE), f"{YMODE_N} respondents"],
              ["Seniority (tSal)", SEN.idxmax(), f"{I(SEN.max())} of {I(len(SAL))}"],
              ["Order quantity (tOrders)", f"{QMODE} item", f"{I(QMODE_N)} of {I(len(ORD))}"],
              ["Length of stay 2023 (tLOS)", f"{LMODE} days", f"{I((LS == LMODE).sum())} admissions"]]),
          watch="A salary mode on a round number (₦12m) shows people round their answers; it is not a 'typical' pay."),

        T("4.1.5", "Range and Interquartile Range",
          defn="The range is maximum minus minimum. The interquartile range (IQR) is Q3 minus Q1: the spread of the middle 50%.",
          simple="One extreme value changes the range; the IQR ignores extremes, so report it with the median.",
          grid=dict(label="Salary spread (tSal)", widths=[2, 2.4], rows=[
              ["Measure", "Salary"],
              ["Minimum / maximum", f"{N(Y.min())} / {N(Y.max())}"],
              ["Range", N(Y.max() - Y.min())],
              ["Q1 / Q3", f"{N(YQ1)} / {N(YQ3)}"],
              ["IQR (middle 50%)", N(YQ3 - YQ1)]]),
          fx=[("=QUARTILE.INC(r,3)-QUARTILE.INC(r,1)", N(YQ3 - YQ1), "IQR")], fx_widths=[2.6, 1.1, 0.8],
          watch="The range describes the two most extreme people, not the typical spread."),

        T("4.1.6", "Variance and Standard Deviation",
          defn="The standard deviation (SD) is the typical distance of values from the mean, in the data's own units; the variance is "
               "SD squared.",
          simple="Use STDEV.S for a sample (most business data) and STDEV.P only when you have the whole population.",
          grid=dict(label="Spread around the mean", widths=[2.4, 1.4, 1.4], rows=[
              ["Column", "Mean", "SD (STDEV.S)"],
              ["Salary (tSal)", N(YM), N(YSD)],
              ["Exam score (tCV)", F(EX.mean()), F(EX.std())],
              ["Length of stay 2023 (tLOS)", f"{LS.mean():.1f} days", f"{LS.std():.1f} days"]]),
          watch="Variance is in squared units (₦ squared); report the SD to managers, not the variance."),

        T("4.1.7", "Coefficient of Variation",
          defn="The coefficient of variation (CV) is the standard deviation divided by the mean, shown as a percentage.",
          simple="It removes the units, so the spread of naira, days and exam points can be compared fairly.",
          grid=dict(label="Which measure varies most? (CV = SD / mean)", widths=[2, 1.3, 1.3, 0.9], size=11.5, rows=[
              ["Measure", "Mean", "SD", "CV"]] + [[n, _m(m_, u), _m(s_, u), P(s_ / m_)] for n, m_, s_, u in CVS]),
          watch="Comparing raw SDs across different units (₦ against days) is meaningless; compare CVs."),

        T("4.1.8", "Skewness",
          defn="Skewness (SKEW) measures lopsidedness: positive means a long tail of large values, negative a long tail of small "
               "values, near zero balanced.",
          simple="Above about +1 the mean is pulled well above the median; report the median.",
          grid=dict(label="Shape of four columns", widths=[2.4, 1, 2], rows=[
              ["Column", "SKEW", "Reading"],
              ["Order value (tOrders)", F2(skew(OV)), SKR(skew(OV))],
              ["Exam score (tCV)", F2(skew(EX)), SKR(skew(EX))],
              ["Salary (tSal)", F2(skew(Y)), SKR(skew(Y))],
              ["Length of stay 2023 (tLOS)", F2(skew(LS)), SKR(skew(LS))]]),
          watch="Skewness is only a summary; look at a histogram before deciding what it means."),

        T("4.1.9", "Kurtosis",
          defn="Kurtosis (KURT) compares the weight of a distribution's tails with a normal curve: positive means heavier tails and "
               "more extreme values, negative means flatter.",
          simple="High kurtosis warns that extreme values will keep appearing; plan for them.",
          grid=dict(label="Tails compared with a normal curve", widths=[2.4, 1, 2], rows=[
              ["Column", "KURT", "Reading"],
              ["Order value (tOrders)", F2(kurt(OV)), KUR(kurt(OV))],
              ["Exam score (tCV)", F2(kurt(EX)), KUR(kurt(EX))],
              ["Salary (tSal)", F2(kurt(Y)), KUR(kurt(Y))],
              ["Length of stay 2023 (tLOS)", F2(kurt(LS)), KUR(kurt(LS))]]),
          watch="Excel's KURT reports excess kurtosis: 0 is a normal curve, not 3."),

        T("4.1.10", "Percentiles and Quartiles",
          defn="The kth percentile is the value below which k% of the data falls; quartiles are the 25th, 50th and 75th percentiles.",
          simple="Use PERCENTILE.INC(range, 0.9) or QUARTILE.INC(range, 1); HR uses them to set pay bands.",
          grid=dict(label="Salary percentiles (tSal)", widths=[1.2, 1.8, 2.4], rows=[
              ["Percentile", "Salary", "Meaning"]] + [
              [f"P{p}", N(v), {10: "Entry level", 25: "Q1", 50: "Median", 75: "Q3", 90: "Top 10% start here"}[p]] for p, v in PCTS],
              note=f"=PERCENTILE.INC(tSal[Annual_Salary_NGN],0.9) returns {N(PCTS[-1][1])}"),
          watch="PERCENTILE.INC and PERCENTILE.EXC give different answers on small data; state which one you used."),

        T("4.1.11", "Outliers and the 1.5 x IQR Rule",
          defn="An outlier is a value far from the rest. The 1.5 x IQR rule flags values above Q3 + 1.5 x IQR or below "
               "Q1 - 1.5 x IQR.",
          simple="Flag outliers, then investigate them; never delete them just to tidy a chart.",
          grid=dict(label="SuperMart Online order values (tOrders)", widths=[2.6, 1.6], rows=[
              ["Step", "Value"],
              ["Q1 / Q3", f"{N(OQ1)} / {N(OQ3)}"],
              ["Upper fence = Q3 + 1.5 x IQR", N(OFENCE)],
              ["Orders above the fence", f"{I(O_OUT)} of {I(len(OV))}"],
              ["Their share of order value", P(O_OUT_SHARE)]]),
          watch=f"Here the 'outliers' are {P(O_OUT_SHARE, 0)} of the money: large electronics orders are the business, not errors."),

        T("4.1.12", "Enabling the Analysis ToolPak",
          defn="The Analysis ToolPak is a free Excel add-in that adds Data › Data Analysis, with tools for descriptive statistics, "
               "histograms, correlation and regression.",
          simple="Turn it on once per computer; the button then appears at the right of the Data tab.",
          steps=["File › Options › Add-ins", "Manage: Excel Add-ins › Go", "Tick Analysis ToolPak › OK",
                 "Data tab › Data Analysis now appears on the right"],
          points=["Output is static: it does not update when the data changes.", "Excel for the web does not have it; use formulas.",
                  "Mac: Tools › Excel Add-ins."],
          watch="ToolPak results are pasted values; rerun the tool after any change to the data."),

        T("4.1.13", "Descriptive Statistics Output Explained",
          defn="Data › Data Analysis › Descriptive Statistics, with Summary statistics ticked, writes a table of mean, standard error, "
               "median, mode, SD, variance, kurtosis, skewness, range, minimum, maximum, sum and count.",
          simple="Every row equals an Excel function, so you can check it with formulas.",
          grid=dict(label="ToolPak rows for salary, with the matching function", widths=[1.7, 1.6, 2.2], size=11.5, rows=[
              ["ToolPak row", "Value", "Same as"],
              ["Mean", N(YM), "AVERAGE"],
              ["Standard Error", N(YSD / np.sqrt(len(Y))), "STDEV.S / SQRT(COUNT)"],
              ["Median / Mode", f"{Nm(YMED)} / {Nm(YMODE)}", "MEDIAN / MODE.SNGL"],
              ["Standard Deviation", N(YSD), "STDEV.S"],
              ["Kurtosis / Skewness", f"{kurt(Y):.2f} / {skew(Y):.2f}", "KURT / SKEW"],
              ["Count", I(len(Y)), "COUNT"]]),
          watch="'Standard Error' is the uncertainty of the mean, not the spread of the data; do not report it as the SD."),

        T("4.1.14", "Reporting Descriptive Statistics Honestly",
          defn="An honest summary states the count, the right centre for the shape, a matching spread measure, and any outliers or "
               "missing values.",
          simple="Median with IQR for lopsided data; mean with SD for balanced data. Round to what the reader needs.",
          grid=dict(label="Two versions of the same finding", widths=[1.2, 3.6], size=11.5, rows=[
              ["Version", "Wording"],
              ["Misleading", f"The average SuperMart Online order is {N(OM)}."],
              ["Honest", f"Half of {I(len(OV))} orders are under {N(OMED)} (IQR {N(OQ1)} to {N(OQ3)}); {I(O_OUT)} large orders lift "
                         f"the mean to {N(OM)}."]]),
          watch="Too many decimals suggest false precision: ₦126,792.15 should be ₦126,800 in a report."),
    ],
    takeaways=["Centre, spread and shape: report all three.",
               "Median and IQR for lopsided data; mean and SD for balanced data.",
               "CV compares spread across different units.",
               "Flag outliers with 1.5 x IQR, then investigate them.",
               "ToolPak output is static: every row matches a function you can check."],
)


M42 = dict(
    id="4.2", title="Distributions, Pareto, Percentiles and Z-Scores",
    subtitle="Read a distribution's shape, apply the 80/20 rule, rank and standardise values, and decide what counts as abnormal.",
    objectives=["Describe a distribution and compare it with the normal curve.",
                "Build a histogram and choose its bins.",
                "Run a Pareto analysis and state the real split.",
                "Rank and band values with percentile ranks.",
                "Standardise with z-scores, combine measures and flag anomalies."],
    topics=[
        T("4.2.1", "What a Distribution Is",
          defn="A distribution shows how often each value, or range of values, occurs in a column.",
          simple="Count the values in each range; the counts show where most values sit and how far they spread.",
          grid=dict(label=f"Kingsway exam scores (tCV, {I(len(EX))} students)", widths=[1.3, 1.2, 1.2], rows=[
              ["Score", "Students", "Share"]] + [[k, I(v), P(v / len(EX))] for k, v in EBINS.items()]),
          watch="A single average hides the shape; two very different distributions can share the same mean."),

        T("4.2.2", "Normal Distribution",
          defn="The normal distribution is the symmetric bell curve: about 68% of values lie within 1 SD of the mean and 95% within "
               "2 SD.",
          simple="Check real data against the 68 / 95 rule before using methods that assume a bell curve.",
          grid=dict(label="How close to normal? Share within 1 and 2 SD of the mean", widths=[2, 1.3, 1.3], rows=[
              ["Column", "Within 1 SD", "Within 2 SD"],
              ["Normal curve", "68.3%", "95.4%"],
              ["Exam score (tCV)", P(W1), P(W2)],
              ["Salary (tSal)", P(SW1), P(((Y - YM).abs() <= 2 * YSD).mean())]]),
          fx=[("=NORM.DIST(75,AVERAGE(tCV[Final_Score]),STDEV.S(tCV[Final_Score]),TRUE)", P(norm.cdf(75, EM, ESD)),
               "Normal share below 75")], fx_widths=[3.6, 0.9, 1.4],
          watch=(f"Exam scores are more peaked than normal ({P(W1, 0)} within 1 SD): normal-curve shares would mislead." if W1 > 0.71 else
                 f"Exam scores are close to the 68 / 95 rule ({P(W1, 0)} and {P(W2, 0)}), so normal-curve shares are a fair guide.")),

        T("4.2.3", "Histograms for Distribution Analysis",
          defn="A histogram is a column chart of counts per bin (a range of values), with no gaps between the columns.",
          simple="Insert › Statistic Chart › Histogram, or COUNTIFS per bin; pick a bin width that shows the shape.",
          grid=dict(label=f"The same {I(len(EX))} exam scores, three bin widths", widths=[1.1, 0.8, 2.6], rows=[
              ["Bin width", "Bins", "What the chart shows"],
              ["2 points", str(int(np.ceil((EX.max() - EX.min() + 1) / 2))), "Spiky columns: noise"],
              ["5 points", str(int(np.ceil((EX.max() - EX.min() + 1) / 5))), f"Peak at {EPEAK} ({P(EBINS[EPEAK] / len(EX), 0)}), tails both sides"],
              ["20 points", "3", "Almost one block: shape hidden"]]),
          steps=["Select tCV[Final_Score]", "Insert › Statistic Chart › Histogram",
                 "Format the horizontal axis › Bin width: 5", "Title the finding, not the variable"],
          split=0.55,
          watch="Very wide bins hide the shape and very narrow ones show noise; try two widths before choosing."),

        T("4.2.4", "Pareto Analysis (the 80/20 Rule)",
          defn="Pareto analysis ranks items by size and finds how few of them make up most of the total, often stated as 80/20.",
          simple="Sort largest first, add a running total as a share, and find where it crosses 80%.",
          grid=dict(label=f"SuperMart Online net revenue by customer ({I(len(CUST))} buyers)", widths=[2.8, 1.6], rows=[
              ["Measure", "Value"],
              ["Customers needed for 80% of revenue", f"{I(K80)} ({P(K80 / len(CUST), 0)})"],
              ["Revenue from the top 20% of customers", P(TOP20)],
              ["Net revenue (Delivered + Completed)", N(CUST.sum())]]),
          watch=f"Here it is 80/{round(K80 / len(CUST) * 100)}, not 80/20: always report the split the data shows."),

        T("4.2.5", "Percentile Rank and Banding",
          defn="A percentile rank (PERCENTRANK.INC) gives the share of values below a given value; banding groups values into ranges "
               "such as quartile bands.",
          simple="'Your salary is at the 75th percentile' is clearer than a raw figure.",
          grid=dict(label="Salary quartile bands (tSal)", widths=[2.2, 2, 1.1], rows=[
              ["Band", "Salary range", "People"],
              ["Band 1 (lowest 25%)", f"up to {N(YQ1)}", I(BAND.iloc[0])],
              ["Band 2", f"{Nm(YQ1)} to {Nm(YMED)}", I(BAND.iloc[1])],
              ["Band 3", f"{Nm(YMED)} to {Nm(YQ3)}", I(BAND.iloc[2])],
              ["Band 4 (top 25%)", f"over {N(YQ3)}", I(BAND.iloc[3])]],
              note=f"=PERCENTRANK.INC(tSal[Annual_Salary_NGN],16000000) returns {P(pct_rank_inc(Y, 16e6))}: the share of salaries below ₦16m"),
          watch="Bands are uneven when many people share a value (₦7m, ₦16m): count each band, do not assume 25%."),

        T("4.2.6", "Z-Scores and Standardisation",
          defn="A z-score is how many standard deviations a value is from the mean: z = (value - mean) / SD, or STANDARDIZE.",
          simple="z = 0 is average, +2 is unusually high, -2 unusually low, whatever the units.",
          grid=dict(label="FY2025 sales by store: z-scores (tDash25)", widths=[2.2, 1.4, 1], rows=[
              ["Store", "Sales", "z"]] + [[r.Store_Name, N(r.Sales), f"{r.z_Sales:+.2f}"]
                                          for r in ST.sort_values("Sales", ascending=False).iloc[[0, 1, -1]].itertuples()]),
          fx=[("=STANDARDIZE(x,AVERAGE(r),STDEV.S(r))", "z-score of x: same as (x - AVERAGE(r)) / STDEV.S(r)")], fx_widths=[3, 3],
          watch="z-scores assume a roughly symmetric spread; with heavy skew, use percentile ranks instead."),

        T("4.2.7", "Standardising Multi-Measure Comparisons",
          defn="Standardising turns measures with different units into z-scores so they can be averaged into one composite score.",
          simple="Average the z-scores of sales, margin and target attainment to rank stores on all three at once.",
          grid=dict(label="SuperMart stores FY2025: composite of three z-scores", widths=[2, 0.8, 0.8, 0.9, 1], size=11, rows=[
              ["Store", "Sales z", "Margin z", "Attain. z", "Composite"]] + [
              [sname(r.Store_Name), f"{r.z_Sales:+.2f}", f"{r.z_Margin:+.2f}", f"{r.z_Attain:+.2f}", f"{r.Comp:+.2f}"]
              for r in STZ.iloc[[0, 1, 2, -1]].itertuples()]),
          watch="Equal weights are a choice; agree the weights with the decision-maker and state them."),

        T("4.2.8", "Detecting Anomalies Systematically",
          defn="Systematic anomaly detection applies one written rule, such as |z| > 2 or the 1.5 x IQR fences, to every row.",
          simple="The rule finds candidates; a person then checks each one against what happened.",
          grid=dict(label="Mercyland daily bed occupancy 2023 (tCensus23), rule |z| > 2", widths=[2.6, 1.8], rows=[
              ["Measure", "Value"],
              ["Mean / SD occupancy", f"{P(OCC.mean())} / {P(OCC.std())}"],
              ["Days flagged", f"{len(ANOM)} of 365, {ANOM_TXT}"],
              ["Lowest day", f"{LOWDAY.Census_Date:%d %b %Y}: {P(LOWDAY.Occupancy_Pct)} (z {OZ.min():.2f})"]]),
          watch="For a hospital, low occupancy is not the risk: also apply a fixed operational limit, such as 95% full."),

        T("4.2.9", "Choosing the Right Distribution Summary",
          defn="The right summary depends on the shape and the question: mean and SD for symmetric data, median and percentiles for "
               "skewed data, counts and shares for categories.",
          simple="Look at the histogram first, then pick the summary.",
          grid=dict(label="Summary by situation", widths=[2.2, 2.4], rows=[
              ["Situation", "Use"],
              ["Symmetric (salary)", "Mean and SD"],
              ["Skewed (order value)", "Median, IQR, percentiles"],
              ["Few items carry the total", "Pareto table"],
              ["Different units", "z-scores or CV"],
              ["Rare extremes", "Percentiles P90 to P99"]]),
          watch="Using the mean and SD on skewed data makes ordinary values look extreme and extreme values look normal."),
    ],
    takeaways=["Look at the distribution before summarising it.",
               "Check the 68 / 95 rule before assuming a bell curve.",
               "Pareto: report the real split, not 80/20 by habit.",
               "Percentile ranks and z-scores put values on a common scale.",
               "Anomaly rules find candidates; people confirm them."],
)

for _m in (M41, M42):
    for _t in _m["topics"]:
        if _t.get("fx"):
            _t.setdefault("stack", True)      # formula tables run full width

MODULES = [M41, M42]
