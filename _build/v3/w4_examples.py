"""Week 4 worked examples (edition 3.1): one extra 'WORKED EXAMPLE' slide after the basic topics, on numbers small enough
to check by hand or with a calculator. Every value below is computed here (numpy, Excel's .S / .INC conventions), so the
slide and the formula agree. Attached as continuation slides, so the topic numbering does not change."""
import numpy as np

from w4_fx import I, N

WX = "WORKED EXAMPLE"


def _f(x, d=2):
    return f"{x:,.{d}f}"


# ---------------------------------------------------------------- 4.1 mean and median: seven salaries (₦ million)
SAL7 = [2, 3, 3, 4, 5, 6, 26]
MEAN7, MED7 = float(np.mean(SAL7)), float(np.median(SAL7))
MEAN6 = float(np.mean(SAL7[:-1]))

# ---------------------------------------------------------------- 4.1 range, quartiles, IQR, fences: nine delivery times (days)
DAYS = [1, 2, 2, 3, 4, 5, 6, 8, 9]
Q1, Q2, Q3 = (float(np.percentile(DAYS, p)) for p in (25, 50, 75))   # = QUARTILE.INC
IQR = Q3 - Q1
LO, HI = Q1 - 1.5 * IQR, Q3 + 1.5 * IQR

# ---------------------------------------------------------------- 4.1 variance and SD: five daily orders
ORD5 = [4, 6, 8, 10, 12]
M5 = float(np.mean(ORD5))
DEV = [x - M5 for x in ORD5]
SS = float(sum(d * d for d in DEV))
VAR5, SD5 = SS / (len(ORD5) - 1), float(np.std(ORD5, ddof=1))

# ---------------------------------------------------------------- 4.2 z-scores: one exam, class mean 60, SD 10
ZM, ZSD = 60, 10
ZS = [("Ada", 75), ("Bola", 55), ("Chidi", 60), ("Dayo", 82)]

# ---------------------------------------------------------------- 4.3 / 4.4 study hours and exam score: six students
HRS = np.array([2, 4, 5, 6, 8, 10])
SCO = np.array([48, 55, 52, 63, 66, 74])
R = float(np.corrcoef(HRS, SCO)[0, 1])
B, A = (float(v) for v in np.polyfit(HRS, SCO, 1))
X7 = 7
P7 = A + B * X7

# ---------------------------------------------------------------- 4.5 moving average and accuracy: six months (₦ million)
MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
SALES6 = [10, 12, 14, 13, 15, 17]
MA3 = [None, None] + [sum(SALES6[i - 2:i + 1]) / 3 for i in range(2, 6)]
ACT4, FC4 = [100, 120, 110, 130], [90, 125, 100, 135]
ERR4 = [a - f for a, f in zip(ACT4, FC4)]
MAE4 = sum(abs(e) for e in ERR4) / len(ERR4)
MAPE4 = sum(abs(e) / a for e, a in zip(ERR4, ACT4)) / len(ERR4)

# ---------------------------------------------------------------- 4.6 Goal Seek and a data table: a small product launch
PRICE, COST, FIXED = 2_500, 1_500, 1_000_000
MARGIN = PRICE - COST
BE = FIXED / MARGIN
TARGET = 500_000
UNITS_T = (TARGET + FIXED) / MARGIN


def profit(u):
    return u * MARGIN - FIXED


UNITS_LIST = [800, 1_000, 1_200, 1_500, 2_000]

MORE = {
    "4.1.3": [dict(
        title="Worked Example: Mean or Median?", pill=WX, defn_label="THE QUESTION",
        defn=f"Seven staff at a small Lagos firm earn (₦ million a year): {', '.join(map(str, SAL7))}. What is a typical salary?",
        steps=[f"Mean: add them up = {sum(SAL7)}; divide by 7 = {_f(MEAN7, 1)}",
               f"Median: already sorted; the 4th of 7 values = {_f(MED7, 0)}",
               f"Drop the ₦26m director: mean falls to {_f(MEAN6, 1)}, median moves only to {_f(float(np.median(SAL7[:-1])), 1)}",
               f"Six of the seven earn less than the mean of {_f(MEAN7, 0)}"],
        fx=[("=AVERAGE(B2:B8)", _f(MEAN7, 1), "Mean"), ("=MEDIAN(B2:B8)", _f(MED7, 0), "Median, the typical salary")],
        fx_widths=[2.2, 0.9, 2.2], order=["steps", "fx"], split=0.55,
        watch="One very large value pulls the mean; the median shows what a typical person earns.")],

    "4.1.5": [dict(
        title="Worked Example: Range, Quartiles and IQR", pill=WX, defn_label="THE QUESTION",
        defn=f"Nine deliveries took (days): {', '.join(map(str, DAYS))}. How spread out are the delivery times?",
        steps=[f"Range = largest − smallest = {DAYS[-1]} − {DAYS[0]} = {DAYS[-1] - DAYS[0]}",
               f"Median (5th of 9) = {_f(Q2, 0)}; Q1 (3rd value) = {_f(Q1, 0)}; Q3 (7th value) = {_f(Q3, 0)}",
               f"IQR = Q3 − Q1 = {_f(Q3, 0)} − {_f(Q1, 0)} = {_f(IQR, 0)}: the middle half spans {_f(IQR, 0)} days",
               f"Outlier fences: Q1 − 1.5 × IQR = {_f(LO, 0)}; Q3 + 1.5 × IQR = {_f(HI, 0)}. No value is outside, "
               f"but a 15-day delivery would be."],
        fx=[("=MAX(B2:B10)-MIN(B2:B10)", I(DAYS[-1] - DAYS[0]), "Range"),
            ("=QUARTILE.INC(B2:B10,1)", _f(Q1, 0), "Q1, in D2"),
            ("=QUARTILE.INC(B2:B10,3)", _f(Q3, 0), "Q3, in D3"),
            ("=D3-D2", _f(IQR, 0), "IQR")],
        fx_widths=[2.6, 0.7, 1.3], order=["steps", "fx"], split=0.5,
        watch="The range uses only the two most extreme values; the IQR ignores them, so one odd delivery does not change it.")],

    "4.1.6": [dict(
        title="Worked Example: Standard Deviation Step by Step", pill=WX, defn_label="THE QUESTION",
        defn=f"A kiosk took {', '.join(map(str, ORD5))} orders on five days. How far does a day typically sit from the average?",
        grid=dict(label="Five days, worked by hand", widths=[0.9, 1.0, 1.4, 1.5], size=12, rows=[
            ["Day", "Orders", "Minus mean", "Squared"]] + [
            [str(i), str(x), (f"{d:+.0f}" if d else "0"), f"{d * d:.0f}"] for i, (x, d) in enumerate(zip(ORD5, DEV), 1)] + [
            ["Total", str(sum(ORD5)), "0", f"{SS:.0f}"]]),
        steps=[f"Mean = {sum(ORD5)} ÷ 5 = {_f(M5, 0)}",
               f"Variance = sum of squares ÷ (n − 1) = {SS:.0f} ÷ 4 = {_f(VAR5, 0)}",
               f"SD = square root of {_f(VAR5, 0)} = {_f(SD5)} orders",
               f"Excel: =STDEV.S(B2:B6) = {_f(SD5)}"],
        order=["grid", "steps"], split=0.52,
        watch="Divide by n − 1 for a sample (STDEV.S). The deviations always add to zero, which is why we square them.")],

    "4.2.6": [dict(
        title="Worked Example: z-Scores on One Exam", pill=WX, defn_label="THE QUESTION",
        defn=f"In one exam the class mean is {ZM} and the SD is {ZSD}. How unusual is each student's score?",
        grid=dict(label="z = (score − mean) ÷ SD", widths=[1.0, 0.9, 1.9, 0.8, 1.8], size=12, rows=[
            ["Student", "Score", "Working", "z", "Reading"]] + [
            [n, str(x), f"({x} − {ZM}) ÷ {ZSD}", (f"{(x - ZM) / ZSD:+.1f}" if x != ZM else "0.0"),
             "Unusually high" if (x - ZM) / ZSD > 2 else ("Above average" if x > ZM else ("Exactly average" if x == ZM else "Below average"))]
            for n, x in ZS]),
        fx=[(f"=STANDARDIZE(75,{ZM},{ZSD})", "1.5", "Ada"), (f"=(82-{ZM})/{ZSD}", "2.2", "Dayo")], fx_widths=[2.3, 0.9, 0.9],
        order=["grid", "fx"], stack=True,
        watch="Beyond ±2 is unusual on a roughly bell-shaped spread. A z of 1.5 means 1.5 SDs above the mean.")],

    "4.3.2": [dict(
        title="Worked Example: Hours Studied and Exam Score", pill=WX, defn_label="THE QUESTION",
        defn="Six students record weekly study hours and exam score. Do the two move together?",
        grid=dict(label="Six students", widths=[1.0, 1.2, 1.2], size=12, rows=[
            ["Student", "Hours / wk", "Score"]] + [[str(i), str(h), str(s)] for i, (h, s) in enumerate(zip(HRS, SCO), 1)]),
        steps=["Draw the scatter first: the points rise from left to right",
               f"=CORREL(B2:B7,C2:C7) = {_f(R)}",
               f"{_f(R)} is close to +1: a strong positive relationship",
               "More hours go with higher scores; it does not prove the hours caused them (4.3.6)"],
        order=["grid", "steps"], split=0.42,
        watch="r runs from −1 to +1. Near 0 means no straight-line relationship; the sign gives the direction.")],

    "4.4.2": [dict(
        title="Worked Example: Fit a Line and Predict", pill=WX, defn_label="THE QUESTION",
        defn=f"Using the same six students (hours: {', '.join(map(str, HRS))}; scores: {', '.join(map(str, SCO))}), what score would "
             f"we expect at {X7} hours a week?",
        fx=[("=SLOPE(C2:C7,B2:B7)", _f(B), "b: extra points per hour"),
            ("=INTERCEPT(C2:C7,B2:B7)", _f(A), "a: score at 0 hours"),
            ("=RSQ(C2:C7,B2:B7)", _f(R * R), "share of variation explained"),
            (f"=FORECAST.LINEAR({X7},C2:C7,B2:B7)", _f(P7, 1), f"predicted at {X7} hours")],
        fx_widths=[3.0, 0.8, 2.0],
        steps=[f"Equation: Score = {_f(A)} + {_f(B)} × Hours",
               f"At {X7} hours: {_f(A)} + {_f(B)} × {X7} = {_f(P7, 1)}",
               f"R² = {_f(R * R)}: hours explain about {R * R:.0%} of the score differences"],
        order=["fx", "steps"], stack=True,
        watch="Predict only inside the data (2 to 10 hours). At 30 hours the line would promise a score above 100.")],

    "4.5.4": [dict(
        title="Worked Example: A 3-Month Moving Average", pill=WX, defn_label="THE QUESTION",
        defn="A shop sold (₦ million) " + ", ".join(f"{m} {s}" for m, s in zip(MON, SALES6)) + ". What is the underlying trend?",
        grid=dict(label="Average of each month and the two before it", widths=[0.8, 1.0, 2.3, 1.2], size=12, rows=[
            ["Month", "Sales", "Working", "3-month MA"]] + [
            [m, str(s), "needs 3 months" if ma is None else f"({' + '.join(map(str, SALES6[i - 2:i + 1]))}) ÷ 3",
             "" if ma is None else _f(ma, 0)] for i, (m, s, ma) in enumerate(zip(MON, SALES6, MA3))]),
        fx=[("=AVERAGE(B2:B4)", "12", "first average, in row 4; fill down")], fx_widths=[1.9, 0.5, 2.4],
        order=["grid", "fx"], stack=True,
        watch="The dip in April (13) disappears in the average, which rises steadily 12, 13, 14, 15: the trend is up.")],

    "4.5.6": [dict(
        title="Worked Example: How Good Was the Forecast?", pill=WX, defn_label="THE QUESTION",
        defn="Four months of actual sales against what was forecast. On average, how far off was the forecast?",
        grid=dict(label="Errors month by month", widths=[0.8, 0.9, 1.0, 0.9, 1.0, 1.2], size=12, rows=[
            ["Month", "Actual", "Forecast", "Error", "Size of error", "% of actual"]] + [
            [m, str(a), str(f), f"{e:+d}", str(abs(e)), f"{abs(e) / a:.1%}"] for m, a, f, e in zip(MON, ACT4, FC4, ERR4)] + [
            ["Average", "", "", "", _f(MAE4, 1), f"{MAPE4:.1%}"]]),
        steps=[f"MAE = average size of error = ({' + '.join(str(abs(e)) for e in ERR4)}) ÷ 4 = {_f(MAE4, 1)}",
               f"MAPE = average % error = {MAPE4:.1%}",
               "Ignore the sign when averaging, or +10 and −10 cancel to a 'perfect' 0"],
        order=["grid", "steps"], stack=True,
        watch=f"Say it plainly: 'the forecast is typically about {_f(MAE4, 1)} (about {MAPE4:.0%}) off'.")],

    "4.6.2": [dict(
        title="Worked Example: Goal Seek for a Profit Target", pill=WX, defn_label="THE QUESTION",
        defn=f"A product sells at {N(PRICE)} and costs {N(COST)} to make; fixed costs are {N(FIXED)}. How many units give a profit "
             f"of {N(TARGET)}?",
        grid=dict(label="The model on the sheet", widths=[1.5, 2.6, 1.3], size=12, rows=[
            ["Cell", "Formula", "Value"],
            ["B1 Price", "input", N(PRICE)], ["B2 Cost", "input", N(COST)], ["B3 Fixed", "input", N(FIXED)],
            ["B4 Units", "input (Goal Seek changes it)", I(UNITS_T)],
            ["B5 Profit", "=B4*(B1-B2)-B3", N(TARGET)]]),
        steps=["Data › What-If Analysis › Goal Seek", f"Set cell B5 · To value {TARGET} · By changing cell B4",
               f"Goal Seek returns {I(UNITS_T)} units", f"Check by hand: {I(UNITS_T)} × {N(MARGIN)} − {N(FIXED)} = {N(TARGET)}",
               f"Break-even (profit 0): {N(FIXED)} ÷ {N(MARGIN)} = {I(BE)} units"],
        order=["grid", "steps"], split=0.5,
        watch="Goal Seek changes one input to hit one target. Note the answer, then put the original input back if needed.")],

    "4.6.4": [dict(
        title="Worked Example: Profit at Different Sales Volumes", pill=WX, defn_label="THE QUESTION",
        defn="Using the same product model, what is the profit at 800, 1,000, 1,200, 1,500 and 2,000 units?",
        grid=dict(label="One-variable data table on B4 (Units)", widths=[1.4, 1.8, 1.6], size=12, rows=[
            ["Units", "Profit", "Reading"]] + [
            [I(u), N(profit(u)), "Loss" if profit(u) < 0 else ("Break-even" if profit(u) == 0 else
                                                              ("Target met" if profit(u) >= TARGET else "Profit"))]
            for u in UNITS_LIST]),
        steps=["Type the units down D2:D6; in E1 type =B5", "Select D1:E6 › What-If Analysis › Data Table",
               "Column input cell: B4 › OK", "Excel fills the profits; change B1 and the table updates"],
        order=["grid", "steps"], split=0.5,
        watch="Each extra 100 units adds ₦100,000 profit here, because each unit earns a ₦1,000 margin.")],
}


def attach(modules):
    n = 0
    for m in modules:
        for t in m["topics"]:
            if t["tag"] in MORE:
                t["more"] = t.get("more", []) + MORE[t["tag"]]
                n += 1
    return n
