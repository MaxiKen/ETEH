"""Week 4 edition-3 slide copy: Modules 4.5 (time series and forecasting) and 4.6 (what-if analysis and Solver).
Data: tMonthly (Lagos Retail monthly sales 2022-2025) and tHotel (Atlantic Pearl bookings). Every number is computed here."""
import calendar

import numpy as np
import pandas as pd
from scipy.optimize import linprog

from v3kit import T
from w4_fx import HOT, I, MON, N, Nk, Nm, P, ROOMS, hw_fit, mae, mape

# ---------------------------------------------------------------- 4.5 figures
S = MON.Sales
YR = MON.groupby("Year").Sales.sum()
YOY = YR.pct_change()
QS = MON.assign(Q=(MON.Month_No - 1) // 3 + 1).groupby("Q").Sales.sum() / S.sum()
TR, TE = MON[MON.Year <= 2024], MON[MON.Year == 2025]
SI = TR.groupby("Month_No").Sales.mean() / TR.Sales.mean()
_ds = TR.Sales / TR.Month_No.map(SI)
_b, _a = np.polyfit(TR.t, _ds, 1)
CLASS = ((_a + _b * TE.t) * TE.Month_No.map(SI)).values
NAIVE = TR[TR.Year == 2024].Sales.values
_b2, _a2 = np.polyfit(TR.t, TR.Sales, 1)
LIN = (_a2 + _b2 * TE.t).values
HWB = hw_fit(TR.Sales.values)
HWF = hw_fit(S.values)
ACT = TE.Sales.values
MA3 = S.tail(3).mean()
MA12 = S.tail(12).mean()
_l = S.iloc[0]
for _v in S:
    _l = 0.3 * _v + 0.7 * _l
SES = _l
AC12 = S.autocorr(12)
AC6 = S.autocorr(6)
MSHORT = [calendar.month_abbr[m] for m in range(1, 13)]
BACK = [("Holt-Winters (as FORECAST.ETS)", HWB["fc"]), ("Classical decomposition", CLASS),
        ("Same month last year", NAIVE), ("Straight-line trend", LIN)]

# ---------------------------------------------------------------- 4.6 figures
H = HOT.assign(C=(HOT.Reservation_Status == "Cancelled").astype(int))
H24 = H[(H.Check_In_Year == 2024)]
HON = H24[H24.Reservation_Status != "Cancelled"]
RN = HON.Nights.sum()
REV = HON.Reservation_Value_NGN.sum()
ADR24 = REV / RN
OCC24 = RN / (ROOMS * 366)
VC, FC, DAYS = 28_000, 3_600_000_000, 365
OCC0, ADR0 = round(OCC24, 2), round(ADR24, -2)


def profit(occ, adr, fixed=FC):
    return ROOMS * DAYS * occ * (adr - VC) - fixed


def revenue(occ, adr):
    return ROOMS * DAYS * occ * adr


ADR_1BN = (1e9 + FC) / (ROOMS * DAYS * OCC0) + VC
OCC_T = {t: (t + FC) / (ROOMS * DAYS * (ADR0 - VC)) for t in (1e9, 3e9, 6e9)}
OCCS = [0.40, 0.45, 0.50, 0.55, 0.60, 0.65, 0.70]
ADRS = [80_000, 100_000, 120_000]
BE_ADR = FC / (ROOMS * DAYS * OCC0) + VC
SCEN = {"Pessimistic": (OCC0 - 0.06, ADR0 * 0.92, FC * 1.05), "Base": (OCC0, ADR0, FC),
        "Optimistic": (OCC0 + 0.06, ADR0 * 1.08, FC)}
SCP = {k: profit(o, a, f) for k, (o, a, f) in SCEN.items()}

COMM = {"Corporate": 0.0, "Airline Crew": 0.0, "Online Travel Agent": 0.18, "Walk-in / Travel Agent": 0.08}
SEGS = list(COMM)
SG = H24[H24.Sales_Channel.isin(SEGS)].groupby("Sales_Channel").agg(adr=("Room_Rate_NGN", "mean"), cr=("C", "mean"))
SG["net"] = SG.adr * (1 - SG.cr) * (1 - pd.Series(COMM))
LO = {"Corporate": 40, "Airline Crew": 10, "Online Travel Agent": 0, "Walk-in / Travel Agent": 0}
HI = {"Corporate": 70, "Airline Crew": 10, "Online Travel Agent": 220, "Walk-in / Travel Agent": 110}
_c = [-SG.net[s] for s in SEGS]
_A = [[1] * 4, [0, 0, 1, 0]]
_bnd = [(LO[s], HI[s]) for s in SEGS]
LP = linprog(_c, A_ub=_A, b_ub=[ROOMS, ROOMS // 2], bounds=_bnd, method="highs", integrality=[1] * 4)
ALLOC = dict(zip(SEGS, np.round(LP.x).astype(int)))
LPV = sum(ALLOC[s] * SG.net[s] for s in SEGS)
SHORTSEG = {"Corporate": "Corporate", "Airline Crew": "Airline Crew", "Online Travel Agent": "Online Travel Agent",
            "Walk-in / Travel Agent": "Walk-in / TA"}


M45 = dict(
    id="4.5", title="Time Series and Forecasting",
    subtitle="Find the trend and the seasonal pattern, smooth with moving averages, forecast with FORECAST.ETS and the "
             "Forecast Sheet, and check how far off the forecast is.",
    objectives=["Describe a time series by its trend, seasonality and noise.",
                "Calculate a moving average to see the trend.",
                "Forecast with FORECAST.ETS or Data › Forecast Sheet.",
                "Measure accuracy with MAE and MAPE.",
                "Chart a forecast with its band and choose a simple method honestly."],
    topics=[
        T("4.5.1", "What a Time Series Is",
          defn="A time series is one measure recorded at regular intervals (daily, monthly, yearly) in time order, with no gaps.",
          simple="Forecasting tools need one row per period, sorted oldest first, with dates in one column and values in another.",
          grid=dict(label=f"Lagos Retail monthly sales (tMonthly, {I(len(MON))} months)", widths=[1.2, 1.6, 1.3], rows=[
              ["Year", "Sales", "Change"]] + [[str(y), Nm(v), "" if np.isnan(YOY[y]) else f"{YOY[y]:+.1%}"] for y, v in YR.items()]),
          watch="A missing month or an unsorted row breaks FORECAST.ETS; check the timeline first."),

        T("4.5.2", "Trend Identification",
          defn="The trend is the long-run direction of a series once the ups and downs within each year are averaged out.",
          simple="A 12-month moving average or a straight-line fit (FORECAST.LINEAR, a chart trendline) shows it.",
          grid=dict(label="Trend in tMonthly", widths=[2.6, 1.8], rows=[
              ["Measure", "Value"],
              ["Straight-line slope, 2022-24", f"{Nm(_b2, 2)} a month"],
              ["Annual sales 2022 to 2025", f"{Nm(YR.iloc[0])} to {Nm(YR.iloc[-1])}"],
              ["12-month average, end of 2025", Nm(MA12)]]),
          watch="2023 dipped before two years of growth: one straight line hides that turn."),

        T("4.5.3", "Seasonality",
          defn="Seasonality is a pattern that repeats every year (or week); a seasonal index is each month's average divided by the "
               "overall monthly average.",
          simple="Index 1.00 is a normal month; 1.85 is 85% above normal.",
          grid=dict(label="Seasonal index, 2022-24 (tMonthly)", widths=[0.62] * 6, size=11, rows=[
              MSHORT[:6], [f"{SI[m]:.2f}" for m in range(1, 7)], MSHORT[6:], [f"{SI[m]:.2f}" for m in range(7, 13)]]),
          watch=f"Quarter 4 brings {P(QS[4])} of sales and Quarter 1 only {P(QS[1])}: plan stock and staff by season."),

        T("4.5.4", "Moving Averages",
          defn="A moving average replaces each value with the average of the last n periods, smoothing out noise.",
          simple="Short windows react fast; a 12-month window removes the season entirely.",
          fx=[("=AVERAGE(E47:E49)", Nm(MA3), "3-month: last 3 months of 2025"),
              ("=AVERAGE(E38:E49)", Nm(MA12), "12-month: the whole of 2025")],
          fx_widths=[2.6, 1.0, 2.6],
          notes="tMonthly has Sales in column E, Jan 2022 in row 2 and Dec 2025 in row 49.",
          watch="The 3-month average (Oct-Dec) is inflated by the peak; neither is a forecast for January."),

        T("4.5.5", "FORECAST.ETS",
          defn="FORECAST.ETS(target_date, values, timeline) forecasts the next periods from the history, combining the recent "
               "level, the trend and the seasonal pattern.",
          simple="You give it the dates and the values; Excel finds the pattern. It is also behind Data › Forecast Sheet.",
          grid=dict(label="2026 forecast from 2022-25 monthly sales", widths=[2.4, 1.8], rows=[
              ["Item", "Forecast"],
              ["January 2026", Nm(HWF["fc"][0])],
              ["December 2026", Nm(HWF["fc"][11])],
              ["Full year 2026", Nm(HWF["fc"].sum())],
              ["Growth on 2025", f"{HWF['fc'].sum() / YR.iloc[-1] - 1:+.1%}"]],
              note="=FORECAST.ETS(DATE(2026,1,1),tMonthly[Sales],tMonthly[Month])"),
          watch="Excel tunes its smoothing settings slightly differently, so its numbers can differ a little from these."),

        T("4.5.6", "Measuring Forecast Accuracy",
          defn="MAE is the average size of the errors; MAPE is the average error as a % of the actual.",
          simple="Error = actual - forecast. Forecast 2025 from 2022-24 only, then compare with what really happened.",
          grid=dict(label="Forecast made from 2022-24, checked against 2025", widths=[1.4, 1.3, 2.4], rows=[
              ["Measure", "Value", "Formula"],
              ["MAE", Nm(mae(ACT, HWB["fc"]), 2), "AVERAGE(ABS(error))"],
              ["MAPE", P(mape(ACT, HWB["fc"])), "AVERAGE(ABS(error)/actual)"]]),
          watch="MAPE explodes when actuals are near zero; use MAE for small or zero-heavy series."),

        T("4.5.7", "Forecast Charts and Confidence Bands",
          defn="A forecast chart shows the actual history, the forecast and the confidence band in one picture.",
          simple="Data › Forecast Sheet builds it in one step: pick the end date and the confidence level.",
          steps=["Select the Month and Sales columns", "Data › Forecast Sheet", "Forecast end: Dec 2026; Options: confidence 95%",
                 "Create: a new sheet with the table and chart"],
          points=["Actuals solid; forecast dashed.", "Shade the band, do not draw it bold.", "Label the forecast start."],
          watch="A forecast line without its band looks far more certain than it is."),

        T("4.5.8", "Forecast Limitations and Assumptions",
          defn="Every forecast assumes the future behaves like the past: the same trend, the same season, no new shocks.",
          simple="State the assumptions next to the number.",
          grid=dict(label="Assumptions behind the 2026 forecast", widths=[1.6, 3.2], rows=[
              ["Assumption", "Risk"],
              ["Trend continues", "2024-25 growth may slow"],
              ["Same seasonal shape", "A moved promotion shifts the peak"],
              ["No shocks", "FX, fuel prices, a new competitor"],
              ["History is clean", "One-off orders inflate the base"]]),
          watch="Re-forecast monthly and compare with actuals; a forecast is a living number."),

        T("4.5.9", "Choosing a Forecasting Method",
          defn="Choose the simplest method that captures the patterns in the data and has the smallest error on recent actuals.",
          simple="Look at the chart, check for a trend and a season, compare the errors (MAE, MAPE), then pick.",
          grid=dict(label="Which method when?", widths=[1.8, 1.6, 1.4], rows=[
              ["Data shows", "Method", "Excel tool"],
              ["No trend, no season", "Moving average", "AVERAGE"],
              ["Trend only", "Straight line", "FORECAST.LINEAR"],
              ["Trend and season", "Seasonal forecast", "FORECAST.ETS"]]),
          watch=f"On this series the straight line missed 2025 by {P(mape(ACT, LIN), 0)}: it ignored the season."),
    ],
    takeaways=["A time series needs regular, sorted, complete periods.",
               "Separate trend, season and noise before forecasting.",
               "A moving average shows the trend; FORECAST.ETS handles trend and season.",
               "Check a forecast against real actuals with MAE and MAPE.",
               "Pick the simplest method with the smallest error; show the band and state the assumptions."],
)


M46 = dict(
    id="4.6", title="Goal Seek, Data Tables, Scenarios and Solver",
    subtitle="Answer what-if questions with a clean input-output model: Goal Seek for a target, data tables for sensitivity, "
             "scenarios for named cases, and Solver for the best plan within limits.",
    objectives=["Build a model with separate inputs and formula outputs.",
                "Use Goal Seek and know its limits.",
                "Build one- and two-variable data tables and read a sensitivity grid.",
                "Store and compare scenarios with Scenario Manager.",
                "Set up a simple Solver problem, and document the analysis."],
    topics=[
        T("4.6.1", "What Is What-If Analysis?",
          defn="What-if analysis changes a model's inputs to see how its outputs respond: for a target, a range, named cases or "
               "the best answer.",
          simple="Inputs in their own cells, outputs as formulas that use them.",
          grid=dict(label="Atlantic Pearl profit model (tHotel, 2024 stayed stays)", widths=[2.2, 2.4], rows=[
              ["Input", "Value"],
              ["Rooms", f"{ROOMS} (constant)"],
              ["Occupancy (2024 actual)", P(OCC24)],
              ["Average daily rate (2024)", N(ADR24)],
              ["Variable cost per night; fixed cost", f"{N(VC)}; {Nm(FC / 1e3, 1).replace('m', 'bn')} (assumed)"]],
              note="Profit = Rooms x 365 x Occupancy x (ADR - Variable cost) - Fixed cost"),
          watch=f"Base model uses {P(OCC0, 0)} and {N(ADR0)} (rounded 2024 actuals): profit {Nm(profit(OCC0, ADR0))}."),

        T("4.6.2", "Goal Seek",
          defn="Goal Seek (Data › What-If Analysis) finds the one input value that makes a formula cell reach a target.",
          simple="Set cell = the output; To value = the target; By changing cell = one input.",
          grid=dict(label=f"What ADR gives ₦1bn profit at {P(OCC0, 0)} occupancy?", widths=[2.4, 1.8], rows=[
              ["Measure", "Result"],
              ["Current ADR", N(ADR0)],
              ["ADR Goal Seek returns", N(ADR_1BN)],
              ["Increase needed", f"{N(ADR_1BN - ADR0)} a night"]],
              note="Check: (target + fixed) / (rooms x 365 x occupancy) + variable cost"),
          watch="Goal Seek gives one answer to many decimals; round it and check it with algebra."),

        T("4.6.3", "Goal Seek Limitations",
          defn="Goal Seek changes only one input, cannot respect limits, and returns a single value even when it is impossible.",
          simple="Always ask: can this answer happen?",
          grid=dict(label=f"Occupancy needed at ADR {N(ADR0)}", widths=[1.4, 1.8, 1.4], rows=[
              ["Profit target", "Occupancy returned", "Possible?"]] + [
              [Nm(t / 1e3, 0).replace("m", "bn"), P(o), "Yes" if o <= 0.85 else "No" if o > 1 else "Hard"] for t, o in OCC_T.items()]),
          watch=(f"Goal Seek returns {P(OCC_T[6e9])} for ₦6bn without complaint: a hotel cannot be more than full." if OCC_T[6e9] > 1 else
                 f"Goal Seek returns {P(OCC_T[6e9])} for ₦6bn without complaint; check it against the best occupancy ever achieved.")),

        T("4.6.4", "One-Variable Data Tables",
          defn="A one-variable data table (Data › What-If Analysis › Data Table) recalculates a formula for a list of values of "
               "one input.",
          simple="List the values down a column, link the output formula one row up and one column right, then set the Column input cell.",
          grid=dict(label=f"Profit at ADR {N(100_000)} for each occupancy", widths=[1.4, 2.2], rows=[
              ["Occupancy", "Profit"]] + [[P(o, 0), N(profit(o, 100_000))] for o in OCCS[:5]]),
          watch=f"Each 5 points of occupancy adds {Nm(profit(0.45, 1e5) - profit(0.40, 1e5))} at this ADR: occupancy is a big lever."),

        T("4.6.5", "Two-Variable Data Tables",
          defn="A two-variable data table recalculates one formula for every combination of two inputs: one across the top, one "
               "down the side.",
          simple="Formula link in the top-left corner; Row input = ADR, Column input = occupancy.",
          grid=dict(label="Profit by occupancy (rows) and ADR (columns)", widths=[1.0, 1.4, 1.4, 1.4], size=11, rows=[
              ["", "₦80k", "₦100k", "₦120k"]] + [[P(o, 0)] + [Nm(profit(o, a)) for a in ADRS] for o in OCCS[::2]]),
          watch="Data tables are live but slow on big models; set Calculation › Automatic Except Data Tables while building."),

        T("4.6.6", "Reading a Sensitivity Table",
          defn="A sensitivity table shows which input moves the output most and where the result changes sign.",
          simple="Colour the grid (red loss, green profit) and find the break-even line.",
          grid=dict(label="Lowest occupancy in the grid that makes a profit", widths=[1.2, 1.6, 1.8], rows=[
              ["ADR", "Occupancy", "Profit there"]] + [
              [Nk(a), P(o, 0), Nm(profit(o, a))] for a in ADRS for o in [next(x for x in OCCS if profit(x, a) > 0)]]),
          watch=f"At {P(OCC0, 0)} occupancy, break-even ADR is about {N(round(BE_ADR, -2))}; today's {N(ADR0)} is only {P(ADR0 / BE_ADR - 1, 0)} above it."),

        T("4.6.7", "Scenario Manager",
          defn="Scenario Manager (Data › What-If Analysis) stores named sets of input values and swaps them into the model on demand.",
          simple="Use it when several inputs move together: a bad year means lower occupancy and lower rates and higher costs.",
          grid=dict(label="Three 2025 scenarios (assumptions)", widths=[1.3, 1.0, 1.1, 1.2, 1.4], size=11, rows=[
              ["Scenario", "Occupancy", "ADR", "Fixed cost", "Profit"]] + [
              [k, P(o, 0), N(a), Nm(f / 1e3, 2).replace("m", "bn"), Nm(SCP[k])] for k, (o, a, f) in SCEN.items()]),
          watch="A scenario that changes only one input is a data table in disguise."),

        T("4.6.8", "Scenario Summary Reports",
          defn="Scenario Manager › Summary builds a new sheet comparing every scenario's inputs and results side by side.",
          simple="Name the input and result cells first, so the report shows names, not $C$4.",
          grid=dict(label="What management reads from the summary", widths=[2.6, 1.6], rows=[
              ["Measure", "Value"],
              ["Worst case profit", Nm(SCP["Pessimistic"])],
              ["Best case profit", Nm(SCP["Optimistic"])],
              ["Swing to plan for", Nm(SCP["Optimistic"] - SCP["Pessimistic"])]]),
          watch="The summary is static: rerun it after any change to the model."),

        T("4.6.9", "Enabling Solver",
          defn="Solver is a free Excel add-in that finds the best value of a formula by changing several cells within constraints.",
          simple="Enable once: File › Options › Add-ins › Manage Excel Add-ins › Go › tick Solver Add-in. It appears on the Data tab.",
          grid=dict(label=f"The Atlantic Pearl problem: split {ROOMS} rooms across four segments", widths=[1.5, 3.3], rows=[
              ["Solver part", "In this example"],
              ["Objective", "Maximise expected net value per night"],
              ["Changing cells", "Rooms given to each segment"],
              ["Constraints", f"Total <= {ROOMS}; within each segment's demand; OTA <= 50%"]]),
          watch="On a Mac: Tools › Excel Add-ins. Goal Seek cannot handle more than one input or any limit."),

        T("4.6.10", "Formulating an Optimisation Problem",
          defn="An optimisation needs an objective cell, decision cells, and constraints, all linked by formulas.",
          simple="Net value per room = ADR x (1 - cancellation rate) x (1 - commission).",
          grid=dict(label="Net value per room, 2024 bookings (tHotel)", widths=[1.6, 0.9, 0.9, 0.9, 1.0], size=11, rows=[
              ["Segment", "ADR", "Cancel", "Comm.", "Net"]] + [
              [SHORTSEG[s], Nk(SG.adr[s]), P(SG.cr[s], 0), P(COMM[s], 0), N(SG.net[s])] for s in SG.net.sort_values(ascending=False).index]),
          watch="OTA has the highest ADR but the lowest net value after cancellations and commission."),

        T("4.6.11", "Documenting What-If Analysis",
          defn="Documentation records the inputs and their source, the assumptions, the tools and settings used, the results and "
               "the limits.",
          simple="Someone else must be able to rerun and challenge the analysis.",
          grid=dict(label="Documentation for the Atlantic Pearl plan", widths=[1.3, 3.5], rows=[
              ["Item", "Recorded"],
              ["Inputs", "2024 stayed bookings from tHotel"],
              ["Assumptions", f"{ROOMS} rooms; {N(VC)} variable; {Nm(FC / 1e3, 1).replace('m', 'bn')} fixed; scenario values"],
              ["Tools", "Goal Seek, data tables, 3 scenarios, Solver (Simplex LP)"],
              ["Limits", "Annual averages; no seasonality; demand caps assumed"]]),
          watch="Keep the notes on the model sheet; a separate document goes out of date."),
    ],
    takeaways=["Inputs in cells, outputs as formulas: the base of all what-if tools.",
               "Goal Seek finds one input for one target, with no limits.",
               "Data tables show sensitivity; colour them to find break-even.",
               "Scenarios move several inputs together; summaries compare them.",
               "Solver optimises within constraints; document everything."],
)

for _m in (M45, M46):
    for _t in _m["topics"]:
        if _t.get("fx"):
            _t.setdefault("stack", True)

MODULES = [M45, M46]
