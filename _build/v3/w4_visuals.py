"""Week 4 edition-3.2 visuals for Modules 4.1-4.5: a native PowerPoint chart on nearly every topic, and worked tables that
show how each answer is reached with actual values from the week's datasets (tOrders, tSal, tCV, tCensus23, tDash25,
tHotel, tMonthly). Every number is computed here from the same figures the slide copy uses.

attach(modules) overrides the listed keys of each topic: a value of None removes that key (e.g. an fx block folded into a
table note, so each slide keeps two body blocks: the worked table and the chart)."""
import numpy as np
import pandas as pd
from scipy.stats import norm

import w4_m1_2 as A
import w4_m3_4 as B
import w4_m5_6 as C
from slidekit import GREEN, NAVY, ORANGE, RED
from w4_fx import CV, EX, I, LS, MON, N, Nm, OV, ORD, P, SAL, Y, mae, mape

GREY = "9AA5B1"
FM = '₦#,##0,,"m"'
FM1 = '₦#,##0.0,,"m"'
FK = '₦#,##0,"k"'


def _m(x, d=1):
    return f"₦{x / 1e6:,.{d}f}m"


def _u(v, unit):
    return N(v) if unit == "₦" else (f"{v:.1f} days" if unit == "days" else f"{v:.1f}")


def _cols(vals, hi, base=NAVY, mark=ORANGE):
    return [mark if i in hi else base for i in range(len(vals))]


# ---------------------------------------------------------------- 4.1 worked samples (actual rows)
# seven consecutive SuperMart Online order lines (tOrders rows 19-25): one large order among small ones
S7 = ORD.iloc[17:24][["order_id", "total_price"]].reset_index(drop=True)
S7V = S7.total_price.astype(float).tolist()
S7_SUM, S7_MEAN, S7_MED = sum(S7V), float(np.mean(S7V)), float(np.median(S7V))
S7_SORT = sorted(zip(S7V, S7.order_id))
S8 = ORD.iloc[24]            # an 8th order, for the even-count rule
S8_SORT = sorted(S7V + [float(S8.total_price)])
S8_MED = float(np.median(S8_SORT))
assert S7_SUM % 7 == 0 and S7_MEAN > 4 * S7_MED

QTY = ORD.quantity.value_counts().sort_index()
YMIN, YMAX = float(Y.min()), float(Y.max())
SB = {s: SAL[SAL.Seniority == s].Annual_Salary_NGN for s in ["Junior", "Mid", "Senior", "Executive"]}

# exam-score histogram (5-point bins), coloured inside / outside one SD of the mean
_eb = list(range(int(EX.min() // 5 * 5), int(EX.max()) + 6, 5))
EXB = [(a, int(((EX >= a) & (EX < a + 5)).sum())) for a in _eb[:-1]]
EXB = [(a, n) for a, n in EXB if n > 0]
_lo1, _hi1 = A.EM - A.ESD, A.EM + A.ESD

# salary histogram (₦2m bins)
_sb = list(range(2_000_000, 26_000_001, 2_000_000))
SALB = [(a, int(((Y >= a) & (Y < a + 2_000_000)).sum())) for a in _sb[:-1]]

# percentile curve
PC = [(p, float(np.percentile(Y, p))) for p in range(10, 100, 10)]

# outliers: share of orders and of value inside / above the upper fence
_out = OV > A.OFENCE

# ---------------------------------------------------------------- 4.2
ST = A.ST.sort_values("Sales", ascending=False).reset_index(drop=True)
SMEAN, SSD = float(A.ST.Sales.mean()), float(A.ST.Sales.std())
STZ = A.STZ.sort_values("Comp", ascending=False)
OCC = A.OCC
OLO, OHI = OCC.mean() - 2 * OCC.std(), OCC.mean() + 2 * OCC.std()
RATIO = [("Exam score (tCV)", EX.mean() / EX.median()), ("Salary (tSal)", Y.mean() / Y.median()),
         ("Length of stay (tLOS)", LS.mean() / LS.median()), ("Order value (tOrders)", OV.mean() / OV.median())]

# ---------------------------------------------------------------- 4.3
R = B.R
ORDR = sorted(B.DRIVERS, key=lambda c: -R[c])
SAMP = CV.sample(800, random_state=4)          # 800 of the 5,840 students, so each dot can be seen
R_AH = CV.Lecture_Attendance_Pct.corr(CV.Study_Hours_Wk)
ABAND = pd.cut(CV.Lecture_Attendance_Pct, [0, 60, 70, 80, 90, 101], right=False,
               labels=["under 60%", "60-69%", "70-79%", "80-89%", "90%+"])
AB = CV.groupby(ABAND, observed=True).Final_Score.mean()

# ---------------------------------------------------------------- 4.4
S1, A1, B1 = B.S1, B.A1, B.B1
SALS = SAL.sample(700, random_state=4)
XB = pd.cut(SAL.Years_of_Experience, [-0.1, 2, 5, 10, 15, 20, 60], labels=["0-2", "3-5", "6-10", "11-15", "16-20", "21+"])
XMID = {"0-2": 1, "3-5": 4, "6-10": 8, "11-15": 13, "16-20": 18, "21+": 25}
XAVG = SAL.groupby(XB, observed=True).Annual_Salary_NGN.mean()
YEARS = list(range(0, int(B.XMAX) + 1, 2))
SEY = S1["sey"]

# ---------------------------------------------------------------- 4.5
S, YR = C.S, C.YR
MLAB = [pd.Timestamp(m).strftime("%b %y") for m in MON.Month]
MA12S = S.rolling(12).mean()
MA3S = S.rolling(3).mean()
FC26 = float(C.HWF["fc"].sum())
_LAB = {"Holt-Winters (as FORECAST.ETS)": "FORECAST.ETS", "Same month last year": "Same month last year",
        "Straight-line trend": "Straight line (FORECAST.LINEAR)"}          # only the methods taught in 4.5
BACKM = sorted([(_LAB[n], mape(C.ACT, f)) for n, f in C.BACK if n in _LAB], key=lambda x: x[1])
TE_LAB = [pd.Timestamp(m).strftime("%b") for m in C.TE.Month]


def _nn(x):
    return None if pd.isna(x) else float(x)


VIS = {
    # ============================================================ 4.1
    "4.1.1": dict(chart=dict(kind="histogram", cats=[f"{a // 1_000_000}-{a // 1_000_000 + 2}m" for a, _ in SALB],
                             series=[("People", [n for _, n in SALB])], fmt="#,##0", legend=False, x_title="Annual salary (₦)",
                             title=f"Centre near {_m(A.YMED, 0)}, most between {_m(A.YQ1, 0)} and {_m(A.YQ3, 0)}"),
                  split=0.48),

    "4.1.2": dict(
        fx=None, grid_note=f"=AVERAGE(tSal[Annual_Salary_NGN]) = {N(A.YM)}",
        chart=dict(kind="column", cats=[f"#{i}" for i in range(1, 8)] + ["Mean"], series=[("Order value", S7V + [S7_MEAN])],
                   fmt=FK, labels=True, label_fmt=FK, legend=False, point_colors=_cols(S7V + [0], {7}),
                   title=f"7 real orders: one large order lifts the mean to {N(S7_MEAN)}"),
        split=0.5, stack=False),

    "4.1.3": dict(order=["grid", "chart"], split=0.42),

    "4.1.4": dict(
        grid=dict(label="Worked example: count each order quantity (tOrders)", widths=[1.3, 1.4, 1.5], rows=[
            ["Quantity", "Order lines", ""]] + [[str(k), I(v), "← most common = mode" if k == A.QMODE else ""]
                                               for k, v in QTY.items()] + [["Total", I(QTY.sum()), ""]],
            note=f"=MODE.SNGL(tOrders[quantity]) = {A.QMODE}"),
        chart=dict(kind="column", cats=[str(k) for k in QTY.index], series=[("Order lines", QTY.tolist())], fmt="#,##0",
                   labels=True, legend=False, x_title="Items in the order line", point_colors=_cols(QTY, {0}),
                   title=f"Quantity {A.QMODE} is the mode: {I(A.QMODE_N)} of {I(len(ORD))} lines"),
        split=0.5),

    "4.1.5": dict(
        grid=dict(label="Worked example: salary spread (tSal, all 1,781)", widths=[1.3, 3.0], size=11, rows=[
            ["Measure", "Working"],
            ["Range", f"Max − Min = {_m(YMAX)} − {_m(YMIN)} = {_m(YMAX - YMIN)}"],
            ["Q1", f"=QUARTILE.INC(r,1) = {_m(A.YQ1)}"],
            ["Q3", f"=QUARTILE.INC(r,3) = {_m(A.YQ3)}"],
            ["IQR", f"Q3 − Q1 = {_m(A.YQ3)} − {_m(A.YQ1)} = {_m(A.YQ3 - A.YQ1)}"]]),
        fx=None,
        chart=dict(kind="box", box=[(s, float(v.min()), float(v.quantile(.25)), float(v.median()), float(v.quantile(.75)),
                                     float(v.max())) for s, v in SB.items()], fmt=FM, y_title="Annual salary",
                   title="Box = IQR (Q1 to Q3, line at the median); whiskers = range"),
        split=0.5, stack=False),

    "4.1.6": dict(
        chart=dict(kind="histogram", cats=[str(a) for a, _ in EXB], series=[("Students", [n for _, n in EXB])], fmt="#,##0",
                   legend=False, x_title="Exam score (bins of 5)",
                   point_colors=[NAVY if _lo1 - 2.5 <= a + 2.5 <= _hi1 + 2.5 else GREY for a, _ in EXB],
                   title=f"Dark bars: within 1 SD ({A.ESD:.1f}) of the mean ({A.EM:.1f})"),
        split=0.52),

    "4.1.7": dict(
        grid=dict(label="Worked example: CV = SD ÷ mean", widths=[1.8, 2.6], size=11, rows=[
            ["Measure", "Working"]] + [[n, f"{_u(s_, u)} ÷ {_u(m_, u)} = {P(s_ / m_)}"] for n, m_, s_, u in A.CVS]),
        chart=dict(kind="bar", cats=[n for n, *_ in sorted(A.CVS, key=lambda c: -c[2] / c[1])],
                   series=[("CV", [s_ / m_ for _, m_, s_, _u in sorted(A.CVS, key=lambda c: -c[2] / c[1])])],
                   fmt="0%", labels=True, legend=False, title="Order value varies most relative to its mean"),
        split=0.52),

    "4.1.8": dict(chart=dict(kind="column", cats=[f"P{p}" for p, _ in PC], series=[("Salary", [v for _, v in PC])], fmt=FM,
                             labels=True, label_fmt=FM, legend=False, point_colors=_cols(PC, {4}),
                             title=f"Salary at each percentile: P50 (median) = {_m(A.YMED, 0)}"),
                  split=0.5),

    "4.1.9": dict(
        grid=dict(label="Worked example: the upper fence (tOrders)", widths=[1.6, 2.8], size=11, rows=[
            ["Step", "Working"],
            ["Q1 / Q3", f"{N(A.OQ1)} / {N(A.OQ3)}"],
            ["IQR", f"Q3 − Q1 = {N(A.OQ3 - A.OQ1)}"],
            ["Upper fence", f"Q3 + 1.5 × IQR = {N(A.OQ3)} + 1.5 × {N(A.OQ3 - A.OQ1)} = {N(A.OFENCE)}"],
            ["Above the fence", f"{I(A.O_OUT)} of {I(len(OV))} orders"],
            ["Their value share", P(A.O_OUT_SHARE)]]),
        chart=dict(kind="column", cats=["Share of orders", "Share of order value"],
                   series=[("Inside the fences", [1 - _out.mean(), 1 - A.O_OUT_SHARE]), ("Above the fence", [_out.mean(), A.O_OUT_SHARE])],
                   fmt="0%", labels=True, label_fmt="0%", colors=[NAVY, ORANGE],
                   title=f"{P(_out.mean(), 0)} of orders are 'outliers' but carry {P(A.O_OUT_SHARE, 0)} of the value"),
        split=0.52),

    "4.1.11": dict(chart=dict(kind="column", cats=["Mean", "Median", "Mode"], series=[("Salary", [A.YM, A.YMED, A.YMODE])],
                              fmt=FM1, labels=True, label_fmt=FM1, legend=False, colors=[NAVY],
                              title="Three ToolPak centre rows for salary, side by side"),
                   split=0.56),

    "4.1.12": dict(chart=dict(kind="box", box=[("Order value", max(float(OV.min()), A.OQ1 - 1.5 * (A.OQ3 - A.OQ1)), A.OQ1, A.OMED,
                                                A.OQ3, A.OFENCE)], fmt=FK, y_title="Order value",
                              title=f"Honest picture: half the orders are under {N(A.OMED)}",
                              note=f"Whiskers stop at the 1.5 x IQR fences; {I(A.O_OUT)} larger orders sit above."),
                   split=0.58),

    # ============================================================ 4.2
    "4.2.1": dict(chart=dict(kind="histogram", cats=list(A.EBINS.index.astype(str)), series=[("Share of students", (A.EBINS / A.EBINS.sum()).tolist())],
                             fmt="0%", labels=True, label_fmt="0%", legend=False, x_title="Exam score",
                             title=f"Most students score {A.EPEAK}; fewer at each end"),
                  split=0.45),

    "4.2.2": dict(
        grid=dict(label="Share within 1 and 2 SD of the mean", widths=[1.7, 1.1, 1.1], rows=[
            ["Column", "Within 1 SD", "Within 2 SD"],
            ["Normal curve", "68.3%", "95.4%"],
            ["Exam score (tCV)", P(A.W1), P(A.W2)],
            ["Salary (tSal)", P(A.SW1), P(((Y - A.YM).abs() <= 2 * A.YSD).mean())]],
            note=f"=NORM.DIST(75,mean,SD,TRUE) = {P(norm.cdf(75, A.EM, A.ESD))}: the normal share below 75"),
        fx=None,
        chart=dict(kind="column", cats=["Within 1 SD", "Within 2 SD"],
                   series=[("Normal curve", [0.683, 0.954]), ("Exam score", [A.W1, A.W2]),
                           ("Salary", [A.SW1, ((Y - A.YM).abs() <= 2 * A.YSD).mean()])],
                   fmt="0%", labels=True, label_fmt="0%", y_min=0, y_max=1, colors=[GREY, NAVY, ORANGE],
                   title="Real columns against the 68 / 95 rule"),
        split=0.47, stack=False),

    "4.2.5": dict(chart=dict(kind="column", cats=["Band 1", "Band 2", "Band 3", "Band 4"], series=[("People", A.BAND.tolist())],
                             fmt="#,##0", labels=True, legend=False,
                             title="People per quartile band: not exactly 25% each"),
                  split=0.55),

    "4.2.6": dict(
        grid=dict(label="Worked example: z = (sales − mean) ÷ SD (tDash25)", widths=[1.5, 2.6, 0.7], size=11, rows=[
            ["Store", "Working", "z"]] + [
            [A.sname(r.Store_Name), f"({_m(r.Sales)} − {_m(SMEAN)}) ÷ {_m(SSD)}", f"{r.z_Sales:+.2f}"]
            for r in ST.iloc[[0, 1, -1]].itertuples()],
            note=f"Mean {_m(SMEAN)}, SD {_m(SSD)}; =STANDARDIZE(x,AVERAGE(r),STDEV.S(r))"),
        fx=None,
        chart=dict(kind="bar", cats=[A.sname(n) for n in ST.Store_Name], series=[("z", ST.z_Sales.tolist())], fmt="+0.0;-0.0;0",
                   legend=False, point_colors=[NAVY if z >= 0 else ORANGE for z in ST.z_Sales],
                   title="Store sales as z-scores: above 0 is above average"),
        split=0.5, stack=False),

    "4.2.7": dict(chart=dict(kind="bar", cats=[A.sname(n) for n in STZ.Store_Name], series=[("Composite", STZ.Comp.tolist())],
                             fmt="+0.0;-0.0;0", legend=False, point_colors=[NAVY if c >= 0 else ORANGE for c in STZ.Comp],
                             title="Composite score: the average of three z-scores"),
                  split=0.55),

    "4.2.8": dict(chart=dict(kind="line", cats=[f"{d:%d %b}" for d in A.C23.Census_Date],
                             series=[("Occupancy", OCC.tolist()), ("Mean − 2 SD", [OLO] * len(OCC)), ("Mean + 2 SD", [OHI] * len(OCC))],
                             fmt="0%", colors=[NAVY, ORANGE, ORANGE], markers=False,
                             title=f"{len(A.ANOM)} days fall outside the ±2 SD lines"),
                  split=0.48),

    "4.2.9": dict(chart=dict(kind="bar", cats=[n for n, _ in RATIO], series=[("Mean ÷ median", [r for _, r in RATIO])],
                             fmt="0.00", labels=True, legend=False, point_colors=[NAVY if r < 1.2 else ORANGE for _, r in RATIO],
                             title="Mean ÷ median above 1.2: lopsided, so use the median"),
                  split=0.5),

    # ============================================================ 4.3
    "4.3.1": dict(chart=dict(kind="bar", cats=[B.SHORT[c] for c in ORDR], series=[("r", [R[c] for c in ORDR])],
                             fmt="+0.00;-0.00;0.00", labels=True, legend=False,
                             point_colors=[NAVY if abs(R[c]) >= 0.3 else GREY for c in ORDR],
                             title="r with exam score: attendance links most strongly"),
                  split=0.48),

    "4.3.2": dict(chart=dict(kind="scatter", points=[("Students", SAMP.Lecture_Attendance_Pct.tolist(), SAMP.Final_Score.tolist())],
                             trend=True, marker=3, x_title="Lecture attendance %", y_title="Exam score",
                             title=f"Attendance and score rise together: r = {R['Lecture_Attendance_Pct']:.2f}",
                             note="800 of the 5,840 students shown, so each point can be seen."),
                  order=["fx", "chart"], stack=False, split=0.62, max_scale=1.0),

    "4.3.3": dict(chart=dict(kind="scatter", points=[("Students", SAMP.Lecture_Attendance_Pct.tolist(), SAMP.Study_Hours_Wk.tolist())],
                             marker=3, x_title="Lecture attendance %", y_title="Hours studied a week",
                             title=f"Two drivers that do not overlap: r = {abs(R_AH) if abs(R_AH) < 0.005 else R_AH:.2f}",
                             note="A cloud with no direction is what r near 0 looks like."),
                  split=0.5),

    "4.3.6": dict(chart=dict(kind="column", cats=list(B.LEAD.index.astype(str)), series=[("Cancelled", B.LEAD["mean"].tolist())],
                             fmt="0%", labels=True, label_fmt="0%", legend=False, x_title="Days booked ahead",
                             title="The earlier the booking, the more cancellations"),
                  split=0.48),

    "4.3.7": dict(chart=dict(kind="column", cats=list(B.GAPX.index.astype(str)),
                             series=[("Female", B.GAPX.Female.tolist()), ("Male", B.GAPX.Male.tolist())], fmt=FM, colors=[ORANGE, NAVY],
                             x_title="Years of experience", title="Like with like: average salary within each experience band"),
                  split=0.5),

    "4.3.9": dict(chart=dict(kind="column", cats=list(AB.index.astype(str)), series=[("Average exam score", AB.tolist())],
                             fmt="0", labels=True, label_fmt="0.0", legend=False, x_title="Lecture attendance",
                             title="What the reader sees: average score by attendance band"),
                  split=0.5),

    # ============================================================ 4.4
    "4.4.1": dict(chart=dict(kind="column", cats=list(XAVG.index.astype(str)), series=[("Average salary", XAVG.tolist())], fmt=FM,
                             labels=True, label_fmt=FM1, legend=False, x_title="Years of experience",
                             title="Salary rises with experience: the pattern regression measures"),
                  split=0.45),

    "4.4.2": dict(chart=dict(kind="scatter", points=[("Respondents", SALS.Years_of_Experience.tolist(), SALS.Annual_Salary_NGN.tolist())],
                             trend=True, marker=3, fmt=FM, x_title="Years of experience", y_title="Annual salary",
                             title=f"Trendline: Salary = {_m(A1, 2)} + {_m(B1, 2)} x Years (R² {S1['r2']:.2f})",
                             note="700 of the 1,781 respondents shown."),
                  order=["fx", "chart"], stack=False, split=0.64, max_scale=1.0, fx_widths=[4.4, 1.1, 1.4]),

    "4.4.3": dict(chart=dict(kind="line", cats=[str(x) for x in YEARS], series=[("Predicted salary", [A1 + B1 * x for x in YEARS])],
                             fmt=FM, labels=[YEARS.index(x) for x in (0, 10, 20) if x in YEARS], label_fmt=FM1, legend=False,
                             x_title="Years of experience", title=f"Each year adds {_m(B1, 2)}: a straight line"),
                  split=0.52),

    "4.4.4": dict(chart=dict(kind="doughnut", cats=["Explained by experience", "Left to other things"],
                             series=[("Share of variation", [S1["r2"], 1 - S1["r2"]])], colors=[NAVY, GREY],
                             title=f"R² = {S1['r2']:.2f}: experience explains {P(S1['r2'], 0)}"),
                  split=0.5),

    "4.4.6": dict(chart=dict(kind="line", cats=[str(x) for x in YEARS],
                             series=[("Prediction", [A1 + B1 * x for x in YEARS]), ("Upper (+2 SE)", [A1 + B1 * x + 1.96 * SEY for x in YEARS]),
                                     ("Lower (−2 SE)", [A1 + B1 * x - 1.96 * SEY for x in YEARS])],
                             fmt=FM, colors=[NAVY, GREY, GREY], x_title="Years of experience",
                             title=f"At 10 years: {_m(B.PRED)} (about {_m(B.PRED - B.PI)} to {_m(B.PRED + B.PI)})"),
                  split=0.48),

    "4.4.7": dict(chart=dict(kind="line", cats=list(XAVG.index.astype(str)),
                             series=[("Actual average", XAVG.tolist()), ("Straight line", [A1 + B1 * XMID[k] for k in XAVG.index.astype(str)])],
                             fmt=FM, colors=[NAVY, ORANGE], x_title="Years of experience",
                             title="The real pattern bends; the straight line does not"),
                  split=0.5),

    "4.4.8": dict(chart=dict(kind="column", cats=["2 years", "5 years", "10 years", "15 years"],
                             series=[("Predicted salary", [A1 + B1 * x for x in (2, 5, 10, 15)])], fmt=FM, labels=True, label_fmt=FM1,
                             legend=False, title="What HR reads: predicted salary by experience"),
                  split=0.55),

    # ============================================================ 4.5
    "4.5.1": dict(chart=dict(kind="line", cats=MLAB, series=[("Sales", S.tolist())], fmt=FM, legend=False,
                             markers=False, title="48 months in time order, one value per month"),
                  split=0.42),

    "4.5.2": dict(chart=dict(kind="line", cats=MLAB, series=[("Monthly sales", S.tolist()), ("12-month average", [_nn(x) for x in MA12S])],
                             fmt=FM, colors=[GREY, NAVY], markers=False, title="The 12-month average shows the trend: a dip in 2023, then growth"),
                  split=0.45),

    "4.5.4": dict(chart=dict(kind="line", cats=MLAB[-24:], series=[("Monthly sales", S.tolist()[-24:]),
                                                                   ("3-month average", [_nn(x) for x in MA3S.tolist()[-24:]])],
                             fmt=FM, colors=[GREY, NAVY], title="The 3-month average smooths the monthly ups and downs"),
                  order=["fx", "chart"], stack=False, split=0.62, max_scale=1.0),

    "4.5.5": dict(chart=dict(kind="column", cats=[str(y) for y in YR.index] + ["2026 forecast"], series=[("Sales", YR.tolist() + [FC26])],
                             fmt=FM, labels=True, label_fmt=FM, legend=False, point_colors=_cols(list(YR) + [0], {len(YR)}),
                             title=f"FORECAST.ETS: 2026 at {_m(FC26, 0)}"),
                  split=0.5),

    "4.5.6": dict(chart=dict(kind="line", cats=TE_LAB, series=[("Actual 2025", list(C.ACT)), ("Forecast from 2022-24", list(C.HWB["fc"]))],
                             fmt=FM, colors=[NAVY, ORANGE], title=f"2025: forecast against actual (MAPE {P(mape(C.ACT, C.HWB['fc']))})"),
                  split=0.5),

    "4.5.8": dict(chart=dict(kind="column", cats=[str(y) for y in C.YOY.index[1:]], series=[("Growth", C.YOY.tolist()[1:])], fmt="0%",
                             labels=True, label_fmt="+0.0%;-0.0%", legend=False,
                             point_colors=[ORANGE if g < 0 else NAVY for g in C.YOY.tolist()[1:]],
                             title="Growth changes year to year: the forecast assumes the recent pattern holds"),
                  split=0.52),

    "4.5.9": dict(chart=dict(kind="bar", cats=[n for n, _ in BACKM], series=[("MAPE on 2025", [m for _, m in BACKM])],
                             fmt="0%", labels=True, label_fmt="0.0%", legend=False,
                             point_colors=[GREEN if n == BACKM[0][0] else GREY for n, _ in BACKM],
                             title=f"Smallest error on 2025: {BACKM[0][0]}"),
                  split=0.6),
}


# ---------------------------------------------------------------- worked-example slides for Mean and Median
WX = "WORKED EXAMPLE"
SAL7 = [2, 3, 3, 4, 5, 6, 26]                     # the seven-salary example already used in edition 3.1
M7, MD7 = float(np.mean(SAL7)), float(np.median(SAL7))

MORE = {
    "4.1.2": [dict(
        title="Worked Example: Calculating the Mean", pill=WX, defn_label="THE QUESTION",
        defn="Seven SuperMart Online order lines are copied from tOrders (rows 19 to 25). What is their mean order value?",
        grid=dict(label="Seven actual orders (tOrders)", widths=[0.5, 2.2, 1.6], size=12, rows=[
            ["#", "order_id", "total_price"]] + [[str(i), o, N(v)] for i, (o, v) in enumerate(zip(S7.order_id, S7V), 1)] + [
            ["", "Total (add all 7)", N(S7_SUM)],
            ["", "Mean = total ÷ 7", N(S7_MEAN)]]),
        steps=[f"Add the seven values: {' + '.join(f'{v:,.0f}' for v in S7V)} = {N(S7_SUM)}",
               "Count the values: n = 7",
               f"Divide: {N(S7_SUM)} ÷ 7 = {N(S7_MEAN)}",
               f"Excel: =AVERAGE(C2:C8) returns {N(S7_MEAN)}",
               f"Check: {sum(v < S7_MEAN for v in S7V)} of the 7 orders are below the mean, because order #2 "
               f"({N(max(S7V))}) pulls it up"],
        order=["grid", "steps"], split=0.47, steps_label="STEP BY STEP",
        watch="The mean uses every value, so one very large order drags it up. 4.1.3 finds the median of the same orders.")],

    "4.1.3": [dict(
        title="Worked Example: Finding the Median", pill=WX, defn_label="THE QUESTION",
        defn=f"Take the same seven orders (mean {N(S7_MEAN)}). What is the median order value?",
        grid=dict(label="Step 1: sort from smallest to largest", widths=[1.0, 2.2, 1.5, 1.4], size=12, rows=[
            ["Position", "order_id", "total_price", ""]] + [
            [str(i), o, N(v), "← middle" if i == 4 else ""] for i, (v, o) in enumerate(S7_SORT, 1)]),
        steps=["Sort the values from smallest to largest",
               "Count them: n = 7 (an odd number)",
               "Middle position = (7 + 1) ÷ 2 = 4, so take the 4th value",
               f"Median = {N(S7_MED)}. Excel: =MEDIAN(C2:C8)",
               f"Even count? Add an 8th order of {N(S8.total_price)}: the middle two are {N(S8_SORT[3])} and "
               f"{N(S8_SORT[4])}, so the median = ({S8_SORT[3]:,.0f} + {S8_SORT[4]:,.0f}) ÷ 2 = {N(S8_MED)}"],
        order=["grid", "steps"], split=0.5, steps_label="STEP BY STEP",
        watch=f"The median ({N(S7_MED)}) is far below the mean ({N(S7_MEAN)}): a typical order is small. "
              f"Use the median for lopsided data."),
        dict(
        title="Worked Example: Mean or Median?", pill=WX, defn_label="THE QUESTION",
        defn=f"Seven staff at a small Lagos firm earn (₦ million a year): {', '.join(map(str, SAL7))}. What is a typical salary?",
        grid=dict(label="Mean and median, worked by hand", widths=[1.1, 1.3, 2.4], size=12, rows=[
            ["Staff", "Salary (₦m)", ""]] + [[f"Staff {i}", str(v), "← 4th of 7 = median" if i == 4 else ""]
                                            for i, v in enumerate(SAL7, 1)] + [
            ["Mean", f"{M7:.0f}", f"Total {sum(SAL7)} ÷ 7 = {M7:.0f}"],
            ["Median", f"{MD7:.0f}", "Middle (4th) value of the sorted list"]]),
        chart=dict(kind="column", cats=[f"S{i}" for i in range(1, 8)] + ["Mean", "Median"], series=[("₦m", SAL7 + [M7, MD7])],
                   fmt="0", labels=True, legend=False, point_colors=[GREY] * 7 + [ORANGE, GREEN], label="CHART OF THE SEVEN SALARIES",
                   title=f"{sum(v < M7 for v in SAL7)} of 7 staff earn less than the mean of ₦{M7:.0f}m"),
        split=0.5,
        watch=f"Without the ₦26m director the mean falls to ₦{np.mean(SAL7[:-1]):.1f}m but the median moves only to "
              f"₦{np.median(SAL7[:-1]):.1f}m. The median shows what a typical person earns.")],
}


def attach(modules):
    n = 0
    for m in modules:
        for t in m["topics"]:
            for k, v in VIS.get(t["tag"], {}).items():
                if k == "grid_note":
                    t["grid"] = dict(t["grid"], note=v)
                elif v is None:
                    t.pop(k, None)
                else:
                    t[k] = v
            if t["tag"] in MORE:
                # the edition-3.1 'Mean or Median?' steps slide on 4.1.3 is replaced by the table versions above
                t["more"] = [x for x in t.get("more", []) if "Mean or Median" not in x.get("title", "")] + MORE[t["tag"]]
            if t.get("chart"):
                n += 1
    return n
