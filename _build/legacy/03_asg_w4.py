"""Week 4 assignment - statistics, forecasting, what-if, storytelling, audit, AI verification, ethics, reconciliation (modules 4.1-4.14)."""
from __future__ import annotations

import numpy as np
import pandas as pd
from scipy import stats

from asg_common import Asg
from wk_common import DATASETS_TEXT, L, hotel, sm_sales

# ------------------------------------------------------------------ data + truth
sal_raw = L("sal_dirty").copy()
sal = L("sal_clean").copy()
EDU = {"Bachelor's Degree": "Bachelor's", "Bachelor's": "Bachelor's", "Master's Degree": "Master's", "Master's": "Master's", "PhD": "PhD", "phD": "PhD",
       "High School": "High School"}
n_dedup = len(sal_raw.drop_duplicates())
_nb = sal_raw.drop_duplicates().dropna().copy()
n_noblank = len(_nb)
_nb["Education_Level"] = _nb.Education_Level.map(EDU)
_nb = _nb.drop_duplicates()
n_std = len(_nb)
n_low = int((_nb.Annual_Salary_NGN < 1_000_000).sum())
_ok = _nb[_nb.Annual_Salary_NGN >= 1_000_000]
n_exp = int((_ok.Years_of_Experience > _ok.Age - 16).sum())
assert n_std - n_low - n_exp == len(sal), (n_std, n_low, n_exp, len(sal))
S = sal.Annual_Salary_NGN
mu, sd = S.mean(), S.std(ddof=1)
q1_, q3_ = S.quantile(.25), S.quantile(.75)
fence = q3_ + 1.5 * (q3_ - q1_)
phd, bach = sal[sal.Education_Level == "PhD"].Annual_Salary_NGN.mean(), sal[sal.Education_Level == "Bachelor's"].Annual_Salary_NGN.mean()

cv = L("cv_perf")
STU = ["Matric_No", "Faculty", "Level", "Sleep_Hrs_Night", "Last_Sem_Score", "Tutorials_Attended", "Study_Hours_Wk", "Lecture_Attendance_Pct", "Final_Score"]
stu = cv[STU].copy()
y = stu.Final_Score.astype(float)
xa, xh = stu.Lecture_Attendance_Pct.astype(float), stu.Study_Hours_Wk.astype(float)
r_att, r_hrs, r_slp = np.corrcoef(xa, y)[0, 1], np.corrcoef(xh, y)[0, 1], np.corrcoef(stu.Sleep_Hrs_Night, y)[0, 1]
lr = stats.linregress(xa, y)
X = np.column_stack([xh, xa, np.ones(len(y))])
beta, *_ = np.linalg.lstsq(X, y, rcond=None)
yhat = X @ beta
r2_multi = 1 - ((y - yhat) ** 2).sum() / ((y - y.mean()) ** 2).sum()
b_hours, b_att = beta[0], beta[1]

h = hotel()
hon = h[(h.Reservation_Status == "Stayed")]
hm = hon.groupby(hon.Check_In.dt.to_period("M")).agg(Bookings=("Reservation_ID", "size"), Room_Nights=("Nights", "sum"),
                                                         Booking_Value=("Reservation_Value_NGN", "sum")).reset_index()
hm["Month"] = hm.Check_In.dt.to_timestamp()
hm["Month_No"] = range(1, len(hm) + 1)
hm = hm[["Month", "Month_No", "Bookings", "Room_Nights", "Booking_Value"]]
NTR = 12   # Jul-2023 .. Jun-2024 training window (one full year)
tr, te = hm.iloc[:NTR], hm.iloc[NTR:]
fit = stats.linregress(tr.Month_No, tr.Room_Nights)
f_lin = fit.intercept + fit.slope * te.Month_No
f_naive = np.repeat(tr.Room_Nights.iloc[-1], len(te))
ape_l = (te.Room_Nights - f_lin).abs() / te.Room_Nights
ape_n = (te.Room_Nights - f_naive).abs() / te.Room_Nights
fit_all = stats.linregress(hm.Month_No, hm.Room_Nights)
f_jan25 = fit_all.intercept + fit_all.slope * (len(hm) + 1)
q4_23 = hm[(hm.Month >= "2023-10-01") & (hm.Month < "2024-01-01")].Room_Nights.sum()
q4_24 = hm[(hm.Month >= "2024-10-01")].Room_Nights.sum()
NM = len(hm)                       # 18 months: Jul 2023 - Dec 2024
q4g = q4_24 / q4_23 - 1
dec_lift = float(hm[hm.Month.dt.month == 12].Room_Nights.mean() / hm[hm.Month.dt.month.isin([10, 11])].Room_Nights.mean() - 1)
low_m = hm.loc[hm.Room_Nights.idxmin(), "Month"]

sm = sm_sales()
fur = sm[(sm.Order_Date.dt.year == 2025) & (sm.Category == "Furniture")]
F_S, F_P = float(fur.Sales.sum()), float(fur.Profit.sum())
F_L = float((fur.Sales / (1 - fur.Discount)).sum())
F_C = F_S - F_P
d_now = 1 - F_S / F_L
d_goal = 1 - F_C / (0.92 * F_L)
marg = lambda d: (F_L * (1 - d) - F_C) / (F_L * (1 - d))

# reconciliation: 2024 booking value vs finance ledger with planted differences
rec = hm[hm.Month >= "2024-01-01"][["Month", "Booking_Value"]].reset_index(drop=True)
ledger = (rec.Booking_Value / 1000).round() * 1000
ledger.iloc[2] -= 12_450_000      # Mar: late posting moved to April (timing)
ledger.iloc[3] += 12_450_000      # Apr
ledger.iloc[8] += 8_730_500       # Sep: no-show charges billed by Finance, not in booking system (permanent)
rec["Finance_Ledger"] = ledger
rec_diff = rec.Finance_Ledger - rec.Booking_Value


def build(key: bool):
    A = Asg(4, "ETEH Week 4 Assignment - Statistics, Forecasting & Responsible Reporting", key,
            "Week 4 · statistics, forecasting, what-if, storytelling, audit, AI and ethics")
    b = A.b

    # ---------------------------------------------------------------- function guide (self-contained help)
    # Every function the tasks need is taught in the Week 1-4 examples workbooks; the guide shows each pattern on this workbook's own tables,
    # answering a DIFFERENT question from the tasks, so it teaches without giving answers away.
    g = b.sheet("Function Guide", "Function guide: the worksheet patterns this assignment uses, on its own tables", A.subtitle, "Reference")
    g.scenario("Every formula below is live and works on this workbook's tables (tSalary, tStudents, tHotelM on the D_ sheets). Each one answers a "
               "different question from the tasks: learn the pattern, then apply it. Full worked examples are in "
               "01_Section_Examples/ETEH_Week4_Section_Examples.xlsx (modules 4.1-4.14).")
    g.h2("Shape and spread: MEDIAN, STDEV.S, PERCENTILE.INC, QUARTILE.INC   [Tasks 1-2]")
    g.calc([("90th-percentile salary", "=PERCENTILE.INC(tSalary[Annual_Salary_NGN],0.9)", float(S.quantile(.9)), "naira", "90% of salaries are at or below this."),
            ("Coefficient of variation", "=STDEV.S(tSalary[Annual_Salary_NGN])/AVERAGE(tSalary[Annual_Salary_NGN])", float(sd / mu), "pct",
             "Spread relative to the mean - compares variability across measures with different units."),
            ("Mean salary, Master's holders", '=AVERAGEIFS(tSalary[Annual_Salary_NGN],tSalary[Education_Level],"Master\'s")',
             float(sal[sal.Education_Level == "Master's"].Annual_Salary_NGN.mean()), "naira", "Text with an apostrophe goes inside the quotes as typed.")])
    g.h2("Where a value sits: STANDARDIZE and NORM.DIST(x, mean, sd, TRUE)   [Task 2]")
    g.calc([("z-score of ₦8,000,000", "=STANDARDIZE(8000000,AVERAGE(tSalary[Annual_Salary_NGN]),STDEV.S(tSalary[Annual_Salary_NGN]))", float((8e6 - mu) / sd), "dec2",
             "Negative = below the mean; |z| > 2 is unusual."),
            ("Normal-model share below ₦8m", "=NORM.DIST(8000000,AVERAGE(tSalary[Annual_Salary_NGN]),STDEV.S(tSalary[Annual_Salary_NGN]),TRUE)",
             float(stats.norm.cdf(8e6, mu, sd)), "pct", "TRUE = cumulative: the area to the LEFT. For 'above', use 1 − NORM.DIST(…).")])
    g.h2("Relationships: CORREL, RSQ, SLOPE, INTERCEPT, FORECAST.LINEAR   [Tasks 3-4]")
    ls = stats.linregress(stu.Last_Sem_Score.astype(float), y)
    g.calc([("r: last-semester score vs final score", "=CORREL(tStudents[Last_Sem_Score],tStudents[Final_Score])", float(ls.rvalue), "dec3", "−1 to +1; sign = direction."),
            ("R² of the same pair", "=RSQ(tStudents[Final_Score],tStudents[Last_Sem_Score])", float(ls.rvalue ** 2), "pct", "y first, then x - share of variation explained."),
            ("Slope (final points per last-semester point)", "=SLOPE(tStudents[Final_Score],tStudents[Last_Sem_Score])", float(ls.slope), "dec3", "y first, then x."),
            ("Predicted final score at last-semester 60", "=FORECAST.LINEAR(60,tStudents[Final_Score],tStudents[Last_Sem_Score])", float(ls.intercept + 60 * ls.slope),
             "dec2", "= INTERCEPT + SLOPE × 60.")])
    g.h2("Several predictors at once: INDEX(LINEST(y, adjacent x columns, TRUE, TRUE), row, col)   [Task 4]")
    X2 = np.column_stack([stu.Sleep_Hrs_Night.astype(float), stu.Last_Sem_Score.astype(float), np.ones(len(y))])
    b2, *_ = np.linalg.lstsq(X2, y, rcond=None)
    r2_2 = 1 - ((y - X2 @ b2) ** 2).sum() / ((y - y.mean()) ** 2).sum()
    g.calc([("R² with sleep + last-semester score", "=INDEX(LINEST(tStudents[Final_Score],tStudents[[Sleep_Hrs_Night]:[Last_Sem_Score]],TRUE,TRUE),3,1)", float(r2_2), "dec3",
             "Row 3, column 1 of the LINEST stats block is R². The x columns must sit side by side."),
            ("Coefficient of Last_Sem_Score", "=INDEX(LINEST(tStudents[Final_Score],tStudents[[Sleep_Hrs_Night]:[Last_Sem_Score]],TRUE,TRUE),1,1)", float(b2[1]), "dec3",
             "LINEST returns coefficients in REVERSE column order: column 1 = the LAST x column.")])
    g.h2("Forecast and test: FORECAST.LINEAR on a time index, APE = |actual − forecast| ÷ actual   [Task 5]")
    fb = stats.linregress(hm.Month_No, hm.Bookings)
    g.calc([("Bookings forecast for Month_No 20 (Feb 2025)", "=FORECAST.LINEAR(20,tHotelM[Bookings],tHotelM[Month_No])", float(fb.intercept + 20 * fb.slope), "int",
             "Month_No is the x; lock training ranges with $ when you copy down."),
            ("APE if Dec 2024 bookings had been forecast at 1,500", f"=ABS(INDEX(tHotelM[Bookings],{NM})-1500)/INDEX(tHotelM[Bookings],{NM})",
             float(abs(hm.Bookings.iloc[-1] - 1500) / hm.Bookings.iloc[-1]), "pct", "MAPE = the AVERAGE of the APE column.")])
    g.h2("Totals for a date window: SUMIFS(sum, date, \">=\"&DATE(…), date, \"<\"&DATE(…))   [Tasks 9, 11]")
    g.calc([("Booking value, Q3 2024", '=SUMIFS(tHotelM[Booking_Value],tHotelM[Month],">="&DATE(2024,7,1),tHotelM[Month],"<"&DATE(2024,10,1))',
             float(hm[(hm.Month >= "2024-07-01") & (hm.Month < "2024-10-01")].Booking_Value.sum()), "naira", "A ratio of two such windows minus 1 is a growth rate.")])

    # ---------------------------------------------------------------- T1 descriptive
    s = A.task("T1 Descriptive Stats", "Task 1 - Clean and describe the ETEH Talent Salary Survey", "4.1",
               "D_Salary_Raw (tSalRaw) is the survey export as received - the form was submitted many times by the same people. "
               "D_Salary (tSalary) is the de-duplicated, complete version. Answer Q1.2-1.3 by cleaning a copy of the raw data; use tSalary for the rest.")
    A.q(s, "1.1", "Rows in the raw export.", "=ROWS(tSalRaw)", len(sal_raw), "int", 1, header=True)
    A.q(s, "1.2", "Rows left after Data > Remove Duplicates (all columns).", n_dedup, None, "int", 1, f"{len(sal_raw) - n_dedup:,} duplicates - 73% of the export.")
    A.q(s, "1.3", "Rows left after also removing rows with ANY blank field.", n_noblank, None, "int", 1)
    A.q(s, "1.4", "Standardise Education_Level ('Bachelor's Degree' → 'Bachelor's', 'Master's Degree' → 'Master's', 'phD' → 'PhD') and remove duplicates again. Rows left?",
        n_std, None, "int", 1, "Label variants hid one more duplicate.")
    A.q(s, "1.5", "How many of those rows have an annual salary below ₦1,000,000 (keying errors such as ₦35,000)?", n_low, None, "int", 1,
        f"Removing them and {n_exp} record with experience > age − 16 gives the {len(sal):,} rows in tSalary.")
    A.q(s, "1.6", "Mean annual salary (₦) - tSalary.", "=AVERAGE(tSalary[Annual_Salary_NGN])", mu, "naira", 1)
    A.q(s, "1.7", "Median annual salary (₦).", "=MEDIAN(tSalary[Annual_Salary_NGN])", S.median(), "naira", 1)
    A.q(s, "1.8", "Sample standard deviation (₦).", "=STDEV.S(tSalary[Annual_Salary_NGN])", sd, "naira", 1)
    A.q(s, "1.9", "Interquartile range (₦).", "=QUARTILE.INC(tSalary[Annual_Salary_NGN],3)-QUARTILE.INC(tSalary[Annual_Salary_NGN],1)", q3_ - q1_, "naira", 1)
    A.q(s, "1.10", "Mean vs median: what shape does that suggest?", "Slight right (positive) skew", marks=1,
        choices=["Slight right (positive) skew", "Left (negative) skew", "Perfectly symmetric", "Bimodal"], guide="Mean ₦11.34m > median ₦11.0m.")
    A.q(s, "1.11", "Mean salary of PhD holders minus Bachelor's holders (₦).",
        '=AVERAGEIFS(tSalary[Annual_Salary_NGN],tSalary[Education_Level],"PhD")-AVERAGEIFS(tSalary[Annual_Salary_NGN],tSalary[Education_Level],"Bachelor\'s")',
        phd - bach, "naira", 2, "The apostrophe in Bachelor's must be inside the quotes.")
    A.end_task(s)

    # ---------------------------------------------------------------- T2 distributions
    s = A.task("T2 Distributions & Z", "Task 2 - Where does a salary sit in the distribution?", "4.2",
               "A candidate has been offered ₦25,000,000. HR asks how unusual that is and whether any survey salaries are outliers.")
    A.q(s, "2.1", "z-score of ₦25,000,000.", "=STANDARDIZE(25000000,AVERAGE(tSalary[Annual_Salary_NGN]),STDEV.S(tSalary[Annual_Salary_NGN]))",
        (25e6 - mu) / sd, "dec2", 2, "≈ 2.7 SD above the mean - unusual but inside the data's range.", header=True)
    within = ((S >= mu - sd) & (S <= mu + sd)).mean()
    A.q(s, "2.2", "Share of salaries within ±1 SD of the mean.",
        '=COUNTIFS(tSalary[Annual_Salary_NGN],">="&(AVERAGE(tSalary[Annual_Salary_NGN])-STDEV.S(tSalary[Annual_Salary_NGN])),'
        'tSalary[Annual_Salary_NGN],"<="&(AVERAGE(tSalary[Annual_Salary_NGN])+STDEV.S(tSalary[Annual_Salary_NGN])))/COUNT(tSalary[Annual_Salary_NGN])',
        within, "pct", 2, "Normal curve predicts ≈ 68%.")
    A.q(s, "2.3", "IQR upper fence (Q3 + 1.5 × IQR).",
        "=QUARTILE.INC(tSalary[Annual_Salary_NGN],3)+1.5*(QUARTILE.INC(tSalary[Annual_Salary_NGN],3)-QUARTILE.INC(tSalary[Annual_Salary_NGN],1))",
        fence, "naira", 2)
    A.q(s, "2.4", "How many survey salaries are above that fence?", f'=COUNTIF(tSalary[Annual_Salary_NGN],">"&{s.addr(s.r - 1, 8)})',
        int((S > fence).sum()), "int", 1, "None - the survey was capped at ₦25m.")
    A.q(s, "2.5", "Under a normal model, probability a salary exceeds ₦20m.",
        "=1-NORM.DIST(20000000,AVERAGE(tSalary[Annual_Salary_NGN]),STDEV.S(tSalary[Annual_Salary_NGN]),TRUE)", 1 - stats.norm.cdf(20e6, mu, sd), "pct", 1,
        f"Actual share above ₦20m: {(S > 20e6).mean():.1%}.")
    A.end_task(s)

    # ---------------------------------------------------------------- T3 correlation
    s = A.task("T3 Correlation", "Task 3 - What moves exam scores at Kingsway?", "4.3",
               "D_Students (tStudents) holds 5,840 Kingsway students. The Dean wants to know which habits are most associated with exam performance.")
    A.q(s, "3.1", "Correlation between Lecture_Attendance_Pct and Final_Score.", "=CORREL(tStudents[Lecture_Attendance_Pct],tStudents[Final_Score])", r_att, "dec3", 2, header=True)
    A.q(s, "3.2", "Correlation between Study_Hours_Wk and Final_Score.", "=CORREL(tStudents[Study_Hours_Wk],tStudents[Final_Score])", r_hrs, "dec3", 2)
    A.q(s, "3.3", "Which of these has the WEAKEST relationship with Final_Score?", "Sleep_Hrs_Night", marks=1,
        choices=["Sleep_Hrs_Night", "Lecture_Attendance_Pct", "Study_Hours_Wk", "Last_Sem_Score"], guide=f"r = {r_slp:.3f}.")
    A.q(s, "3.4", "Share of exam-score variance explained by attendance alone (R²).", "=RSQ(tStudents[Final_Score],tStudents[Lecture_Attendance_Pct])", r_att ** 2, "pct", 1)
    A.long(s, "3.5", "The Dean says: 'Attendance causes higher scores, so we will make lectures compulsory.' Give one reason to be cautious.",
           "Correlation is not causation: motivated students both attend more AND study more (confounder - Motivation, Guardian_Support). "
           "Compulsory attendance may raise attendance without raising motivation, so scores may not rise as much as r suggests. "
           "A pilot with a comparison group would test the effect.", marks=2, rows=3, guide="Names confounding/reverse causation 1; concrete example 1.")
    A.end_task(s)

    # ---------------------------------------------------------------- T4 regression
    s = A.task("T4 Regression", "Task 4 - Predict exam score from attendance and study hours", "4.4",
               "Fit a simple regression (Final_Score on Lecture_Attendance_Pct) and a multiple regression (on Study_Hours_Wk AND Lecture_Attendance_Pct - adjacent columns in tStudents).")
    A.q(s, "4.1", "Slope: exam points per attendance point.", "=SLOPE(tStudents[Final_Score],tStudents[Lecture_Attendance_Pct])", lr.slope, "dec3", 2, header=True)
    A.q(s, "4.2", "Intercept.", "=INTERCEPT(tStudents[Final_Score],tStudents[Lecture_Attendance_Pct])", lr.intercept, "dec2", 1)
    A.q(s, "4.3", "Predicted exam score at 80% attendance (simple model).", "=FORECAST.LINEAR(80,tStudents[Final_Score],tStudents[Lecture_Attendance_Pct])",
        lr.intercept + lr.slope * 80, "dec1", 2)
    xs = "tStudents[[Study_Hours_Wk]:[Lecture_Attendance_Pct]]"
    A.q(s, "4.4", "R² of the multiple model (LINEST with stats).", f"=INDEX(LINEST(tStudents[Final_Score],{xs},TRUE,TRUE),3,1)", r2_multi, "dec3", 2,
        "Data Analysis ToolPak > Regression gives the same.")
    A.q(s, "4.5", "Coefficient of Study_Hours_Wk in the multiple model.", f"=INDEX(LINEST(tStudents[Final_Score],{xs},TRUE,TRUE),1,2)", b_hours, "dec3", 2,
        "LINEST returns coefficients in REVERSE column order: row 1 = {Attendance, Hours, Intercept}.")
    bh = f"{b_hours:.2f}"
    A.q(s, "4.6", "Best interpretation of that coefficient:", f"+1 study hour/week ≈ +{bh} points, attendance held constant", marks=1,
        choices=[f"+1 study hour/week ≈ +{bh} points, attendance held constant", f"Studying causes {b_hours:.0%} higher scores", f"Hours explain {b_hours:.0%} of variance",
                 f"Each hour adds {bh} to attendance"])
    A.end_task(s)

    # ---------------------------------------------------------------- T5 forecasting
    s = A.task("T5 Forecasting", "Task 5 - Forecast Atlantic Pearl room-nights and test the method", "4.5",
               f"D_Hotel_Monthly (tHotelM) holds stayed bookings by arrival month, Jul 2023 - Dec 2024 (Month_No 1-{NM}). Fit on the first {NTR} months "
               "(Jul 2023 - Jun 2024), test on Jul-Dec 2024, then forecast January 2025.")
    A.q(s, "5.1", f"December room-nights are about {dec_lift:.0%} above October-November in both years. What is this?", "A seasonal peak (festive travel)", marks=1,
        choices=["A seasonal peak (festive travel)", "A data-entry error to delete", "Random noise", "Proof that the trend is rising"], header=True,
        guide="It repeats every year at the same time, so it is seasonality - a straight line cannot capture it.")
    s.band("5.2", "Holdout test - fill the yellow columns (train = rows 1-11 of tHotelM)")
    r0, c0 = 4, 1      # header row / first col of tHotelM on D_Hotel_Monthly (asserted below)
    e1, e2 = r0 + 2, r0 + 1 + NTR   # Excel rows of the training window
    TY, TX = f"'D_Hotel_Monthly'!$E${e1}:$E${e2}", f"'D_Hotel_Monthly'!$C${e1}:$C${e2}"
    NAIVE = f"'D_Hotel_Monthly'!$E${e2}"
    hold = pd.DataFrame({"Month": te.Month.values, "Month_No": te.Month_No.values, "Actual": te.Room_Nights.values})
    top = s.r + 1
    A.fill(s, hold, {"Linear": (f"=FORECAST.LINEAR(C{{r}},{TY},{TX})", f_lin.tolist(), "int"),
                     "Naive": (f"={NAIVE}", f_naive.tolist(), "int"),
                     "APE_Linear": ("=ABS(D{r}-E{r})/D{r}", ape_l.tolist(), "pct"),
                     "APE_Naive": ("=ABS(D{r}-F{r})/D{r}", ape_n.tolist(), "pct")},
           fmts={"Month": "mon", "Actual": "int"}, marks=4, guide="Linear with $-locked training ranges 2; naive = last training month (Jun 2024) 1; APE 1.")
    h1, h2 = top + 1, top + len(hold)
    A.q(s, "5.3", "MAPE of the linear-trend forecast.", f"=AVERAGE(G{h1}:G{h2})", ape_l.mean(), "pct", 2, header=True)
    A.q(s, "5.4", "MAPE of the naive forecast.", f"=AVERAGE(H{h1}:H{h2})", ape_n.mean(), "pct", 1)
    best = "Naive" if ape_n.mean() < ape_l.mean() else "Linear trend"
    A.q(s, "5.5", "Which method was more accurate on the holdout?", best, marks=1, choices=["Linear trend", "Naive"],
        guide=(f"Naive MAPE {ape_n.mean():.1%} vs linear {ape_l.mean():.1%}: the line fitted on one year chases the Dec 2023 peak and Jan-Feb 2024 dip."
               if best == "Naive" else f"Linear MAPE {ape_l.mean():.1%} vs naive {ape_n.mean():.1%}."))
    A.q(s, "5.6", f"Linear-trend forecast for January 2025 (Month_No {NM + 1}) using ALL {NM} months.", f"=FORECAST.LINEAR({NM + 1},tHotelM[Room_Nights],tHotelM[Month_No])",
        f_jan25, "int", 2)
    A.long(s, "5.7", "Give one reason not to trust the January 2025 trend forecast without adjustment.",
           f"The series is short ({NM} months) and strongly seasonal: every December peaks and the following January drops ({low_m:%B %Y} was the lowest month). "
           f"A straight line ignores that, so it forecasts January at an average level ({f_jan25:,.0f}) when last January was {int(hm.Room_Nights[hm.Month == '2024-01-01'].iloc[0]):,}. "
           "Use a seasonal method (FORECAST.ETS or same-month-last-year) for monthly planning.", marks=1, rows=2)
    A.end_task(s)

    # ---------------------------------------------------------------- T6 what-if
    s = A.task("T6 What-If Analysis", "Task 6 - What discount cap rescues Furniture margin?", "4.6",
               "SuperMart Furniture made a 1.4% margin in 2025. Model: Revenue = List value × (1 − discount); Cost is fixed at its 2025 level. "
               "The inputs below come from the 2025 sales lines.")
    inp = s.inputs([("Furniture list value before discount, 2025 (₦)", round(F_L), "naira", "Σ Sales ÷ (1 − Discount) over 2025 Furniture lines"),
                    ("Furniture cost of goods, 2025 (₦)", round(F_C), "naira", "Sales − Profit"),
                    ("Furniture actual sales, 2025 (₦)", round(F_S), "naira", "After discounts"),
                    ("Discount to test", 0.10, "pct", "Change this cell; Goal Seek changes it too")])
    LV, CO, SA, DI = (inp[k] for k in ["Furniture list value before discount, 2025 (₦)", "Furniture cost of goods, 2025 (₦)", "Furniture actual sales, 2025 (₦)", "Discount to test"])
    Lr, Cr = round(F_L), round(F_C)
    mfx = lambda d: (Lr * (1 - d) - Cr) / (Lr * (1 - d))
    A.q(s, "6.1", "Effective average discount in 2025.", f"=1-{SA}/{LV}", 1 - round(F_S) / Lr, "pct", 2, header=True)
    A.q(s, "6.2", "Margin at the discount in the input cell (10%).", f"=({LV}*(1-{DI})-{CO})/({LV}*(1-{DI}))", mfx(0.10), "pct", 2)
    A.q(s, "6.3", "Goal Seek: discount that gives an 8% margin (check algebraically).", f"=1-{CO}/(0.92*{LV})", 1 - Cr / (0.92 * Lr), "pct", 3,
        "Goal Seek: Set 6.2 cell To 0.08 By changing the discount input.")
    s.band("6.4", "One-variable data table: margin at each discount level")
    dt = pd.DataFrame({"Discount": [0, 0.05, 0.10, 0.15, 0.20, 0.25]})
    A.fill(s, dt, {"Margin": (f"=({LV}*(1-B{{r}})-{CO})/({LV}*(1-B{{r}}))", [mfx(d) for d in dt.Discount], "pct")}, fmts={"Discount": "pct0"}, marks=2,
           guide="Formula copied down with $-locked inputs (or a What-If Data Table).")
    A.q(s, "6.5", "Main weakness of this model?", "It assumes volume does not change with discount", marks=1,
        choices=["It assumes volume does not change with discount", "It uses too many inputs", "Margins cannot be negative", "Goal Seek is inaccurate"], header=True)
    A.end_task(s)

    # ---------------------------------------------------------------- T7 storytelling
    s = A.task("T7 Story, Report & Slides", "Task 7 - Communicate the Atlantic Pearl forecast", "4.7-4.9",
               "You must brief the Atlantic Pearl General Manager on Task 5 in one slide and a short written summary.")
    HL = (f"Q4 room-nights fell {-q4g:.0%} year on year; plan 2025 on a seasonal baseline" if q4g < 0 else
          f"Q4 room-nights rose {q4g:.0%} year on year; plan 2025 on a seasonal baseline")
    A.q(s, "7.1", "Which is the best slide headline?", HL, marks=2,
        choices=[HL, "Monthly room-nights chart", "Forecast analysis 2023-2025",
                 "Data shows interesting trends"], header=True, guide="A headline states the finding with a number.")
    A.q(s, "7.2", "What comes FIRST in an executive report?", "Key finding and recommendation", marks=1,
        choices=["Key finding and recommendation", "Methodology", "Data sources", "Appendix tables"])
    A.q(s, "7.3", "How many key messages should one slide carry?", "One", marks=1, choices=["One", "Three", "As many as fit", "Five"])
    A.q(s, "7.4", "Best visual for actual vs forecast room-nights with the holdout period marked?", "Line chart with forecast shaded", marks=2,
        choices=["Line chart with forecast shaded", "Pie chart", "3-D column chart", "Table only"])
    A.long(s, "7.5", "Write a 3-sentence executive summary for the General Manager (finding, evidence, recommendation).",
           f"Stayed room-nights {'fell' if q4g < 0 else 'rose'} {abs(q4g):.0%} year on year in Q4 (Q4 2024 {q4_24:,} vs {q4_23:,} in Q4 2023), with a December peak every year. "
           f"On the July-December holdout a simple 'same as last month' forecast {'beat' if ape_n.mean() < ape_l.mean() else 'trailed'} a straight-line trend "
           f"(MAPE {ape_n.mean():.1%} vs {ape_l.mean():.1%}), so a straight line is not a safe planning basis. Plan 2025 staffing and rates on a seasonal baseline "
           "(peak December, low January-February) and revisit the forecast quarterly.",
           marks=4, rows=4, guide="Finding 1, evidence with numbers 2, actionable recommendation 1.")
    A.end_task(s)

    # ---------------------------------------------------------------- T8 audit
    s = A.task("T8 Audit & Protect", "Task 8 - Find the four planted errors before this goes to the board", "4.10",
               "A colleague built the 2025 zone summary below. It contains FOUR errors. For each, give the cell address and the fix. "
               "Then say how you would protect the sheet.")
    zz = sm[sm.Order_Date.dt.year == 2025].groupby("Zone").Sales.sum().reindex(["South-West", "South-South", "North-Central", "South-East"])
    z24 = sm[sm.Order_Date.dt.year == 2024].groupby("Zone").Sales.sum().reindex(zz.index)
    ws = s.ws
    s.h2("FY2025 sales by zone (as received)")
    hdr = s.r
    for j, t in enumerate(["Zone", "FY2024 (₦)", "FY2025 (₦)", "Growth", "Share of FY2025"]):
        ws.write(hdr, 1 + j, t, A.F["hdr"])
    for i, z in enumerate(zz.index):
        r = hdr + 1 + i
        R = r + 1
        ws.write(r, 1, z, A.F["cell"])
        ws.write_number(r, 2, float(z24[z]), b.nf("naira"))
        if i == 2:
            ws.write_string(r, 3, f"₦{zz[z]:,.0f}", b.nf("naira"))                                # ERROR 1: number stored as text
        else:
            ws.write_number(r, 3, float(zz[z]), b.nf("naira"))
        if i == 3:
            ws.write_formula(r, 4, "=0.18", b.nf("pct"), 0.18)                                     # ERROR 2: hard-coded growth
        elif i == 2:
            ws.write_formula(r, 4, f"=IFERROR(D{R}/C{R}-1,0)", b.nf("pct"), 0)
        else:
            ws.write_formula(r, 4, f"=D{R}/C{R}-1", b.nf("pct"), float(zz[z] / z24[z] - 1))
        tot_R = hdr + 6
        if i == 0:
            ws.write_formula(r, 5, f"=D{R}/D${tot_R}", b.nf("pct"), float(zz[z] / (zz.sum() - zz.iloc[2])))
        else:
            ws.write_formula(r, 5, f"=D{R}/D{tot_R + i}", b.nf("pct"), "#VALUE!" if i == 2 else "#DIV/0!")                          # ERROR 3: relative denominator drifts
    tr_ = hdr + 5
    ws.write(tr_, 1, "Total", A.F["lab"])
    ws.write_formula(tr_, 2, f"=SUM(C{hdr + 2}:C{hdr + 4})", b.nf("naira", style="result"), float(z24.iloc[:3].sum()))   # ERROR 4: range misses last zone
    ws.write_formula(tr_, 3, f"=SUM(D{hdr + 2}:D{hdr + 5})", b.nf("naira", style="result"), float(zz.sum() - zz.iloc[2]))
    s.r = tr_ + 5     # keep the rows the drifting denominators point at empty
    T = lambda i, c: f"{c}{hdr + 2 + i}"
    A.q(s, "8.1", "Error A - a value that SUM ignores. Cell?", T(2, "D"), marks=2, guide="North-Central FY2025 typed as text '₦…' - re-enter as a number (VALUE/SUBSTITUTE).", header=True)
    A.q(s, "8.2", "Error B - a hard-coded result. Cell?", T(3, "E"), marks=2, guide="South-East growth typed as 0.18 - replace with =D/C-1.")
    A.q(s, "8.3", "Error C - a copy-down reference error. First wrong cell?", T(1, "F"), marks=2, guide=f"Share denominator drifts (D{tr_ + 2} …) - lock it: =D/D${tr_ + 1}.")
    A.q(s, "8.4", "Error D - a total that misses a row. Cell?", f"C{tr_ + 1}", marks=1, guide=f"=SUM(C{hdr + 2}:C{hdr + 4}) omits South-East - extend to row {hdr + 5}.")
    A.q(s, "8.5", "Which feature stops users typing over formulas but lets them change inputs?", "Unlock input cells, then Protect Sheet", marks=1,
        choices=["Unlock input cells, then Protect Sheet", "Mark as Final", "Hide the sheet", "Save as PDF"])
    A.end_task(s)

    # ---------------------------------------------------------------- T9 AI verification
    s = A.task("T9 Verify AI Claims", "Task 9 - An AI assistant summarised the data. Check every claim.", "4.11-4.12",
               "For each claim enter TRUE or FALSE, then the correct figure as a live formula. Never paste an AI number into a report unchecked.")
    claims = [
        ("'The median salary in the survey is ₦11.3m.'", "FALSE", "=MEDIAN(tSalary[Annual_Salary_NGN])", S.median(), "naira", "₦11.3m is the MEAN; median is ₦11.0m."),
        ("'Attendance explains about a fifth of the variation in Kingsway exam scores.'", "TRUE" if 0.17 <= r_att ** 2 <= 0.25 else "FALSE",
         "=RSQ(tStudents[Final_Score],tStudents[Lecture_Attendance_Pct])", r_att ** 2, "pct", f"R² = {r_att ** 2:.1%}."),
        ("'PhD holders earn about 75% more than Bachelor's holders on average.'", "TRUE" if 1.70 <= phd / bach <= 1.80 else "FALSE",
         '=AVERAGEIFS(tSalary[Annual_Salary_NGN],tSalary[Education_Level],"PhD")/AVERAGEIFS(tSalary[Annual_Salary_NGN],tSalary[Education_Level],"Bachelor\'s")',
         phd / bach, "dec2", f"Ratio {phd / bach:.2f}x."),
        ("'Atlantic Pearl room-nights grew by more than 50% from Q4 2023 to Q4 2024.'", "TRUE" if q4_24 / q4_23 > 1.5 else "FALSE",
         '=SUMIFS(tHotelM[Room_Nights],tHotelM[Month],">="&DATE(2024,10,1))/SUMIFS(tHotelM[Room_Nights],tHotelM[Month],">="&DATE(2023,10,1),tHotelM[Month],"<"&DATE(2024,1,1))-1',
         q4g, "pct", f"{q4g:+.1%}: room-nights {'fell' if q4g < 0 else 'rose'}."),
    ]
    for i, (t, tf, fx, v, k, gd) in enumerate(claims, 1):
        A.q(s, f"9.{i}a", f"Claim: {t}  TRUE or FALSE?", tf, marks=1, choices=["TRUE", "FALSE"], header=(i == 1))
        A.q(s, f"9.{i}b", "Correct figure (live formula).", fx, v, k, 1, gd)
    A.end_task(s)

    # ---------------------------------------------------------------- T10 ethics
    s = A.task("T10 Ethics & Privacy", "Task 10 - Classify the fields before sharing data", "4.13",
               "Mercyland wants to share its admissions table with a university researcher. Classify each field under the Nigeria Data Protection Act 2023 categories used in class.")
    opts = ["Direct identifier (PII)", "Sensitive health data (PHI)", "Quasi-identifier", "Not personal data"]
    for i, (f, a) in enumerate([("Patient_Name", opts[0]), ("Diagnosis_Group", opts[1]), ("Age + Gender + Admit_Date together", opts[2]),
                                ("Ward staffed-bed count", opts[3])], 1):
        A.q(s, f"10.{i}", f"Field: {f}", a, marks=1, choices=opts, header=(i == 1))
    A.end_task(s)

    # ---------------------------------------------------------------- T11 reconciliation
    s = A.task("T11 Reconcile & Benchmark", "Task 11 - Reconcile booking value with the Finance ledger", "4.14",
               "Finance's ledger of 2024 room revenue does not match the booking system. Compute the differences, find the months that break and explain them.")
    top = s.r + 1
    A.fill(s, rec, {"Difference": ("=D{r}-C{r}", rec_diff.tolist(), "naira"),
                    "Flag": ('=IF(ABS(E{r})>1000,"Investigate","OK")', ["Investigate" if abs(x) > 1000 else "OK" for x in rec_diff], "gen")},
           fmts={"Month": "mon", "Booking_Value": "naira", "Finance_Ledger": "naira"}, marks=3,
           guide="Difference = ledger − system 2; flag with ±₦1,000 rounding tolerance 1.")
    s.ws.set_column(2, 4, 16)
    r1, r2 = top + 1, top + len(rec)
    big = rec_diff.abs().idxmax()
    A.q(s, "11.1", "Which month has the largest absolute difference?", rec.Month[big].strftime("%b %Y"), marks=1, header=True,
        guide="Mar and Apr are equal and opposite (₦12.45m timing); either accepted.")
    A.q(s, "11.2", "Net difference for the whole year (₦).", f"=SUM(E{r1}:E{r2})", rec_diff.sum(), "naira", 2, "≈ ₦8.73m + rounding: the Sep no-show charges.")
    A.long(s, "11.3", "Classify the March/April and September differences and say what action each needs.",
           "March −₦12.45m / April +₦12.45m: TIMING difference - revenue posted in the wrong month; nets to zero over the year. Action: correct the posting period.\n"
           "September +₦8.73m: PERMANENT difference - no-show charges billed by Finance but not recorded in the booking system. Action: record them in the system "
           "(or document as a reconciling item) so both sources agree. Remaining small differences are ledger rounding to ₦1,000.",
           marks=4, rows=4, guide="Timing vs permanent 2; action for each 2.")
    A.end_task(s)

    # ---------------------------------------------------------------- data
    b.data_sheet("D_Salary_Raw", sal_raw, "tSalRaw", fmts={"Annual_Salary_NGN": "naira"}, title="ETEH Talent Salary Survey - raw form export (table tSalRaw)",
                 widths={"Job_Title": 28})
    b.data_sheet("D_Salary", sal, "tSalary", fmts={"Annual_Salary_NGN": "naira"}, title="ETEH Talent Salary Survey - clean (table tSalary)",
                 widths={"Job_Title": 28, "Job_Family": 22})
    b.data_sheet("D_Students", stu, "tStudents", title="Kingsway University students, 2024/25 semester 1 (table tStudents)", widths={"Matric_No": 18, "Faculty": 18})
    t = b.data_sheet("D_Hotel_Monthly", hm, "tHotelM", fmts={"Month": "mon", "Booking_Value": "naira", "Room_Nights": "int"},
                     title="Atlantic Pearl stayed bookings by arrival month (table tHotelM)")
    assert (t.r0, t.c0) == (r0, c0) and t.cols[1] == "Month_No" and t.cols[3] == "Room_Nights"
    return A.finish("Week 4 Assignment - Statistics, Forecasting & Responsible Reporting",
                    ["Scenario: ETEH's consulting team has four client questions this week - HR pay benchmarking (Talent Salary Survey), exam drivers for Kingsway's Dean, "
                     "a demand forecast for Atlantic Pearl and a margin rescue plan for SuperMart Furniture - and every number must survive audit, AI-checking and a privacy review.",
                     "Tasks 1-11 follow the Week 4 modules in order. Time guide: 8 hours."],
                    [("Salary Survey", DATASETS_TEXT["Salary Survey"]), ("Kingsway", DATASETS_TEXT["Kingsway"]), ("Atlantic Pearl", DATASETS_TEXT["Atlantic Pearl"]),
                     ("SuperMart Stores", DATASETS_TEXT["SuperMart Stores"] + " - Task 6 uses the 2025 Furniture lines; Task 8 the 2024-25 zone totals.")],
                    rubric=[["Live formulas", "Typed numbers score zero where a formula is required."],
                            ["Tolerance", "Statistics to the decimals shown; money ±₦1; percentages ±0.1 pt."],
                            ["Written answers", "Marked on the finding, the evidence (a number) and the action - not on length."]])


if __name__ == "__main__":
    for k in (False, True):
        print("Week 4", "KEY" if k else "student", "->", build(k))
