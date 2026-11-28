"""Week 4 workbook, modules 4.5 - 4.9 (edition 3): forecasting, what-if analysis and Solver, storytelling, report writing and
presentations.

4.5 works on tMonthly (Sales = Data_Monthly column E, rows 2-49; column F is the one-step smoothing of 4.5.5). FORECAST.ETS
results are cached from the slide's Holt-Winters fit; Excel refits its own parameters when it recalculates, so the check cells
compare your answer with the live key cell, not with a typed number. The backtest uses the slide's Holt-Winters forecast for
2025 as pasted values, so MAE, MAPE and RMSE match the slides. 4.6 has a live Atlantic Pearl profit model (yellow inputs), a
formula version of each data table, three scenarios and the Solver allocation (changing cells pre-filled with Solver's answer).
"""
from __future__ import annotations

import numpy as np
import pandas as pd

import w4_m5_6 as F6
from w1x import ModuleSheet
from wb4_common import N, f, tp

XL0 = pd.Timestamp("1899-12-30")
MSHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MS = "tMonthly[Sales]"
E = "Data_Monthly!$E${}:$E${}"


def xl(d):
    return int((pd.Timestamp(d) - XL0).days)


def build(b, D):
    m45(b, D)
    m46(b, D)
    m47(b, D)
    m48(b, D)
    m49(b, D)


# =====================================================================================================
def m45(b, D):
    MON = D.MON
    S = MON.Sales.values.astype(float)
    YR = MON.groupby("Year").Sales.sum()
    m = ModuleSheet(b, "4.5 Forecasting", "4.5  Time Series and Forecasting",
                    "Tutor walk-through after Module 4.5: trend, seasonal index, moving averages, smoothing, FORECAST.ETS and a backtest",
                    "4.5.1 - 4.5.13")
    m.dataset([("tMonthly", "Sheet Data_Monthly (A1)", "48 months of SuperMart Stores sales, Jan 2022 - Dec 2025. Sales is column E (rows 2-49); "
                "t (column B) numbers the months 1-48."),
               ("SES_Forecast", "Data_Monthly column F", "One-step simple exponential smoothing, alpha 0.3 (4.5.5)."),
               ("tSeason, tBack, tFc", "On this sheet", "Seasonal index, the 2025 backtest and the 2026 forecast with its 95% band.")])

    def line(mm):
        mm.s.chart("line", [{"name": "Monthly sales", "categories": ["Data_Monthly", 1, 0, 48, 0], "values": ["Data_Monthly", 1, 4, 48, 4]}],
                   f"Sales peak every Q4; the year total grew from ₦{YR.iloc[0] / 1e6:.0f}m to ₦{YR.iloc[-1] / 1e6:.0f}m", "", "Sales (₦)", legend=False,
                   y_fmt='₦#,##0,,"m"', rows=15, size=(680, 270),
                   extra=lambda ch: ch.set_x_axis({"num_format": "mmm yy", "num_font": {"size": 8}, "interval_unit": 3}))

    tp(m, "4.5.1",
       [("Months in the series", f"=COUNT({MS})", 48, "int", "One row per month, no gaps."),
        ("2025 sales", f"=SUMIF(tMonthly[Year],2025,{MS})", f(YR[2025]), N, ""),
        ("Growth 2025 on 2024", f"=SUMIF(tMonthly[Year],2025,{MS})/SUMIF(tMonthly[Year],2024,{MS})-1", f(YR[2025] / YR[2024] - 1), "pct", "")],
       [("n", "What was the growth in 2024 on 2023?", f"=SUMIF(tMonthly[Year],2024,{MS})/SUMIF(tMonthly[Year],2023,{MS})-1", f(YR[2024] / YR[2023] - 1), "pct",
         "Growth 2024"),
        ("n", "In how many months of 2025 were sales above ₦60m?", '=COUNTIFS(tMonthly[Year],2025,tMonthly[Sales],">60000000")',
         int(((MON.Year == 2025) & (MON.Sales > 60e6)).sum()), "int", "2025 months above ₦60m")],
       extra=line)

    b2, a2 = np.polyfit(np.arange(1, 37), S[:36], 1)
    bl, al = np.polyfit(np.arange(1, 49), S, 1)
    tp(m, "4.5.2",
       [("Straight-line slope, 2022-24 (per month)", f"=SLOPE({E.format(2, 37)},Data_Monthly!$B$2:$B$37)", f(b2), N, ""),
        ("12-month average, end of 2025", f"=AVERAGE({E.format(38, 49)})", f(F6.MA12), N, "Removes the season."),
        ("12-month average, end of 2024", f"=AVERAGE({E.format(26, 37)})", f(S[24:36].mean()), N, "The trend rose by this much in a year.")],
       [("n", "Straight-line forecast for month 49 (Jan 2026) from all 48 months.", f"=FORECAST.LINEAR(49,{MS},tMonthly[t])", f(al + bl * 49), N,
         "Linear forecast Jan 2026")],
       note="A straight line ignores the season: its January forecast is far above a normal January (4.5.3).")

    tr = MON[MON.Year <= 2024]
    si = tr.groupby("Month_No").Sales.mean() / tr.Sales.mean()
    T = {}

    def season(mm):
        df = pd.DataFrame({"Month_No": range(1, 13), "Month": MSHORT})
        calc = {"Index": ('=AVERAGEIFS(tMonthly[Sales],tMonthly[Month_No],[@Month_No],tMonthly[Year],"<=2024")/AVERAGEIFS(tMonthly[Sales],tMonthly[Year],"<=2024")',
                          "dec2", [f(si[k]) for k in range(1, 13)])}
        mm.h("tSeason: seasonal index, 2022-24 (month average ÷ overall monthly average)")
        T["s"] = mm.s.table(df, "tSeason", fmts={"Month_No": "gen"}, calc=calc, widths={"Month_No": 10, "Month": 8, "Index": 9})
        mm.s.r += 1
        mm.s.chart("column", [{"name": "Seasonal index", "categories": T["s"].ref("Month"), "values": T["s"].ref("Index")}],
                   "Sep, Nov and Dec run far above a normal month; Jan and Feb far below", "", "Index (1 = normal)", legend=False, rows=14,
                   size=(560, 250))

    q4 = MON[MON.Month_No >= 10].Sales.sum() / MON.Sales.sum()
    tp(m, "4.5.3",
       [("December index", "=INDEX(tSeason[Index],12)", f(si[12]), "dec2", "85% above a normal month."),
        ("Weakest month", "=INDEX(tSeason[Month],MATCH(MIN(tSeason[Index]),tSeason[Index],0))", MSHORT[int(si.idxmin()) - 1], "gen", "")],
       [("n", "What is the September index? (2 decimals)", "=INDEX(tSeason[Index],9)", f(si[9]), "dec2", "September index"),
        ("n", "What share of all 2022-25 sales falls in Q4 (Oct-Dec)?", f'=SUMIF(tMonthly[Month_No],">=10",{MS})/SUM({MS})', f(q4), "pct", "Q4 share")],
       extra=season)

    tp(m, "4.5.4",
       [("3-month moving average (Oct-Dec 2025)", "=AVERAGE(Data_Monthly!E47:E49)", f(F6.MA3), N, "Cells quoted on the slide."),
        ("12-month moving average (2025)", "=AVERAGE(Data_Monthly!E38:E49)", f(F6.MA12), N, "")],
       [("n", "What is the 3-month moving average ending September 2025 (rows 45-47)?", "=AVERAGE(Data_Monthly!E45:E47)", f(S[43:46].mean()), N,
         "3-month MA to Sep 2025")],
       do=[("Data › Data Analysis › Moving Average: Input range Data_Monthly!E2:E49, Interval 3, Output G2, tick Chart Output.",
            "A column of 3-month averages (first two #N/A) and a chart: the smoothed line lags the peaks.")])

    ses = F6.SES
    from wb4_common import ses_column
    sc = ses_column(S)
    tp(m, "4.5.5",
       [("Smoothed level after Dec 2025 (alpha 0.3)", "=0.3*Data_Monthly!E49+0.7*Data_Monthly!F49", f(ses), N, "The slide's formula."),
        ("Forecast that was made for Dec 2025", "=Data_Monthly!F49", f(sc[47]), N, "Column F is last month's forecast carried forward.")],
       [("n", "With alpha 0.5 instead, what is 0.5 x Dec 2025 sales + 0.5 x F49?", "=0.5*Data_Monthly!E49+0.5*Data_Monthly!F49", f(.5 * S[47] + .5 * sc[47]), N,
         "Level with alpha 0.5"),
        ("t", "The ToolPak's Exponential Smoothing asks for the damping factor. For alpha 0.3, what do you enter?", "0.7")])

    fc, ci = F6.HWF["fc"], F6.HWF["ci"]
    ETS = "FORECAST.ETS(DATE(2026,{m},1),tMonthly[Sales],tMonthly[Month],12)"
    CI = "FORECAST.ETS.CONFINT(DATE(2026,{m},1),tMonthly[Sales],tMonthly[Month],0.95,12)"

    def fct(mm):
        df = pd.DataFrame({"Month": [xl(f"2026-{k:02d}-01") for k in range(1, 13)]})
        calc = {"Forecast": ("=FORECAST.ETS([@Month],tMonthly[Sales],tMonthly[Month],12)", N, [f(x) for x in fc]),
                "CI_95": ("=FORECAST.ETS.CONFINT([@Month],tMonthly[Sales],tMonthly[Month],0.95,12)", N, [f(x) for x in ci]),
                "Lower": ("=[@Forecast]-[@CI_95]", N, [f(a - c) for a, c in zip(fc, ci)]),
                "Upper": ("=[@Forecast]+[@CI_95]", N, [f(a + c) for a, c in zip(fc, ci)])}
        mm.h("tFc: 2026 forecast with FORECAST.ETS and its 95% confidence interval")
        T["f"] = mm.s.table(df, "tFc", fmts={"Month": "mon"}, calc=calc, widths={"Month": 10, "Forecast": 14, "CI_95": 13, "Lower": 14, "Upper": 14})
        mm.s.r += 1

    tp(m, "4.5.6",
       [("January 2026 forecast", "=" + ETS.format(m=1), f(fc[0]), N, "Triple smoothing: level, trend and season."),
        ("December 2026 forecast", "=" + ETS.format(m=12), f(fc[11]), N, ""),
        ("Full year 2026", "=SUM(tFc[Forecast])", f(fc.sum()), N, "")],
       [("n", "What is the forecast for June 2026?", "=" + ETS.format(m=6), f(fc[5]), N, "June 2026 forecast"),
        ("n", "Forecast growth of 2026 on 2025?", f"=SUM(tFc[Forecast])/SUMIF(tMonthly[Year],2025,{MS})-1", f(fc.sum() / YR[2025] - 1), "pct", "Growth 2026")],
       note="The cached figures are the slide's Holt-Winters fit. Excel fits its own smoothing constants when it recalculates, so its numbers can "
            "differ a little; the check cells follow Excel's live answer.", extra=fct)

    tp(m, "4.5.7",
       [("± 95% interval, January 2026", "=" + CI.format(m=1), f(ci[0]), N, ""),
        ("Lower end, January", "=INDEX(tFc[Lower],1)", f(fc[0] - ci[0]), N, ""), ("Upper end, January", "=INDEX(tFc[Upper],1)", f(fc[0] + ci[0]), N, "")],
       [("n", "What is the ± 95% interval for December 2026?", "=" + CI.format(m=12), f(ci[11]), N, "December CONFINT"),
        ("t", "Does the interval get wider or narrower further ahead?", "wider")])

    ac6 = float(pd.Series(S).autocorr(6))
    tp(m, "4.5.8",
       [("Seasonality Excel detects", "=FORECAST.ETS.SEASONALITY(tMonthly[Sales],tMonthly[Month])", 12, "int", "12 = a yearly pattern in monthly data."),
        ("Correlation with sales 12 months earlier", f"=CORREL({E.format(2, 37)},{E.format(14, 49)})", f(F6.AC12), "dec2", "High: the pattern repeats each year.")],
       [("n", "What is the correlation with sales 6 months earlier? (2 decimals)", f"=CORREL({E.format(2, 43)},{E.format(8, 49)})", f(ac6), "dec2", "Lag-6 correlation")])

    act = F6.ACT
    hw = F6.HWB["fc"]

    def back(mm):
        df = pd.DataFrame({"Month": [xl(f"2025-{k:02d}-01") for k in range(1, 13)], "HW_Forecast": np.round(hw, 0), "Classical": np.round(F6.CLASS, 0)})
        calc = {"Actual": ("=XLOOKUP([@Month],tMonthly[Month],tMonthly[Sales])", N, [f(x) for x in act]),
                "Naive": ("=XLOOKUP(DATE(YEAR([@Month])-1,MONTH([@Month]),1),tMonthly[Month],tMonthly[Sales])", N, [f(x) for x in F6.NAIVE]),
                "Linear": ("=FORECAST.LINEAR(XLOOKUP([@Month],tMonthly[Month],tMonthly[t]),Data_Monthly!$E$2:$E$37,Data_Monthly!$B$2:$B$37)", N,
                           [f(x) for x in F6.LIN]),
                "HW_Error": ("=[@Actual]-[@HW_Forecast]", N, [f(a - round(h_)) for a, h_ in zip(act, hw)])}
        mm.h("tBack: 2025 held out. HW_Forecast and Classical are fitted on 2022-24 only (pasted values); the rest are live")
        T["b"] = mm.s.table(df, "tBack", fmts={"Month": "mon", "HW_Forecast": N, "Classical": N}, calc=calc,
                            widths={"Month": 10, "HW_Forecast": 14, "Classical": 14, "Actual": 14, "Naive": 14, "Linear": 14, "HW_Error": 13})
        mm.s.r += 1
        mm.s.chart("line", [{"name": "Actual 2025", "categories": T["b"].ref("Month"), "values": T["b"].ref("Actual")},
                            {"name": "Holt-Winters forecast", "categories": T["b"].ref("Month"), "values": T["b"].ref("HW_Forecast")}],
                   "The forecast fitted on 2022-24 tracks 2025's shape", "", "", y_fmt='₦#,##0,,"m"', rows=15, size=(620, 260),
                   extra=lambda ch: ch.set_x_axis({"num_format": "mmm", "num_font": {"size": 8}}))

    hwr = np.round(hw, 0)
    err = act - hwr
    tp(m, "4.5.9",
       [("MAE", "=SUMPRODUCT(ABS(tBack[HW_Error]))/COUNT(tBack[HW_Error])", f(np.abs(err).mean()), N, "Average size of a miss."),
        ("MAPE", "=SUMPRODUCT(ABS(tBack[HW_Error])/tBack[Actual])/COUNT(tBack[Actual])", f((np.abs(err) / act).mean()), "pct", "Average % miss."),
        ("RMSE", "=SQRT(SUMPRODUCT(tBack[HW_Error]^2)/COUNT(tBack[HW_Error]))", f(np.sqrt((err ** 2).mean())), N, "Punishes big misses.")],
       [("n", "What is the largest single-month miss (absolute error)?", "=MAX(MAX(tBack[HW_Error]),-MIN(tBack[HW_Error]))", f(np.abs(err).max()), N, "Largest miss"),
        ("t", "Which measure is in % and so compares across series of different size?", "MAPE")],
       extra=back)

    def mp(fc_):
        return float(np.mean(np.abs(act - fc_) / act))
    tp(m, "4.5.10",
       [("MAPE, same month last year", "=SUMPRODUCT(ABS(tBack[Actual]-tBack[Naive])/tBack[Actual])/12", mp(F6.NAIVE), "pct", "The benchmark to beat."),
        ("MAPE, straight-line trend", "=SUMPRODUCT(ABS(tBack[Actual]-tBack[Linear])/tBack[Actual])/12", mp(F6.LIN), "pct", "Worst: ignores the season."),
        ("MAPE, classical decomposition", "=SUMPRODUCT(ABS(tBack[Actual]-tBack[Classical])/tBack[Actual])/12", mp(np.round(F6.CLASS, 0)), "pct", "")],
       [("n", "What 2025 total did the straight-line trend forecast?", "=SUM(tBack[Linear])", f(F6.LIN.sum()), N, "Linear 2025 total"),
        ("t", "Which method had the lowest MAPE?", "Holt-Winters", None, None, ("HW", "holt winters", "FORECAST.ETS", "triple smoothing"))])

    def band(mm):
        mm.s.chart("line", [{"name": "Forecast", "categories": T["f"].ref("Month"), "values": T["f"].ref("Forecast")},
                            {"name": "Lower 95%", "categories": T["f"].ref("Month"), "values": T["f"].ref("Lower"), "color": "#9BB7C9"},
                            {"name": "Upper 95%", "categories": T["f"].ref("Month"), "values": T["f"].ref("Upper"), "color": "#9BB7C9"}],
                   f"2026 forecast ₦{fc.sum() / 1e6:,.0f}m, with a wide band in every month", "", "", y_fmt='₦#,##0,,"m"', rows=15, size=(620, 260),
                   extra=lambda ch: ch.set_x_axis({"num_format": "mmm", "num_font": {"size": 8}}))

    tp(m, "4.5.11",
       [("Widest band (± CI)", "=MAX(tFc[CI_95])", f(ci.max()), N, ""), ("Narrowest band", "=MIN(tFc[CI_95])", f(ci.min()), N, "")],
       [("t", "Should the band lines be drawn darker or lighter than the forecast line?", "lighter")],
       do=[("Select Data_Monthly A1:A49 and E1:E49. Data › Forecast Sheet › Forecast end 2026-12-01 › Options: Seasonality set manually 12 › Create.",
            "A new sheet with the forecast table and a chart: forecast line with lower and upper confidence bounds."),
           ("Compare the new sheet's January 2026 forecast with tFc.", "The same FORECAST.ETS result (Excel's live fit).")], extra=band)

    tp(m, "4.5.12",
       [("Growth 2024 on 2023", f"=SUMIF(tMonthly[Year],2024,{MS})/SUMIF(tMonthly[Year],2023,{MS})-1", f(YR[2024] / YR[2023] - 1), "pct", ""),
        ("Forecast growth 2026 on 2025", f"=SUM(tFc[Forecast])/SUMIF(tMonthly[Year],2025,{MS})-1", f(fc.sum() / YR[2025] - 1), "pct",
         "Assumes the recent growth continues.")],
       [("o", "List two assumptions behind the 2026 forecast and one event that would break it.",
         "The 2024-25 trend continues and the seasonal shape stays the same; a new competitor or a moved promotion would break it.")])

    tp(m, "4.5.13", None,
       [("t", "The data has a trend but no season. Which Excel function?", "FORECAST.LINEAR", None, None, ("forecast linear", "TREND")),
        ("t", "The data has a trend and a season. Which Excel function?", "FORECAST.ETS", None, None, ("forecast ets", "forecast sheet")),
        ("t", "No trend and no season: which simple method?", "moving average", None, None, ("exponential smoothing", "smoothing", "average"))])
    m.finish()


# =====================================================================================================
def m46(b, D):
    H = D.H
    m = ModuleSheet(b, "4.6 What-If", "4.6  What-If Analysis and Solver",
                    "Tutor walk-through after Module 4.6: a live hotel profit model for Goal Seek, data tables, scenarios and Solver",
                    "4.6.1 - 4.6.12")
    m.dataset([("tHotel", "Sheet Data_Hotel (A1)", "Atlantic Pearl reservations: 2024 occupancy, ADR and segment values."),
               ("Model inputs", "Yellow cells on this sheet", "Rooms, occupancy, ADR, variable cost, fixed cost and days: change them and every result updates."),
               ("tOcc, tScen, tSolver", "On this sheet", "The one-variable table, the scenarios and the Solver model.")])
    R, VC, FC, DAYS, OCC0, ADR0 = F6.ROOMS, F6.VC, F6.FC, F6.DAYS, F6.OCC0, F6.ADR0
    A = {}

    def model(mm):
        mm.h("The profit model (yellow = input)")
        A.update(mm.s.inputs([("Rooms", R, "int", "Atlantic Pearl has 320 rooms."), ("Occupancy", OCC0, "pct", "2024 actual, rounded."),
                              ("ADR (average daily rate)", ADR0, N, "2024 actual, rounded to ₦100."),
                              ("Variable cost per occupied night", VC, N, "Assumed."), ("Fixed cost per year", FC, N, "Assumed."),
                              ("Days", DAYS, "int", "")]))

    def P(occ=None, adr=None, fixed=None):
        return (f"{A['Rooms']}*{A['Days']}*{occ or A['Occupancy']}*({adr or A['ADR (average daily rate)']}-{A['Variable cost per occupied night']})"
                f"-{fixed or A['Fixed cost per year']}")

    h24 = "tHotel[Check_In_Year],2024,tHotel[Reservation_Status],\"Stayed\""
    tp(m, "4.6.1",
       lambda mm: [("Room-nights sold, 2024", f"=SUMIFS(tHotel[Nights],{h24})", int(F6.RN), "int", "Stayed bookings with 2024 check-in."),
        ("Occupancy 2024", f"=SUMIFS(tHotel[Nights],{h24})/(320*366)", f(F6.OCC24), "pct", "2024 had 366 days."),
        ("ADR 2024", f"=SUMIFS(tHotel[Reservation_Value_NGN],{h24})/SUMIFS(tHotel[Nights],{h24})", f(F6.ADR24), N, ""),
        ("Profit at the inputs", "=" + P(), f(F6.profit(OCC0, ADR0)), N, "Rooms x days x occupancy x (ADR - variable) - fixed.")],
       [("n", "What is the 2024 room revenue from stayed bookings?", f"=SUMIFS(tHotel[Reservation_Value_NGN],{h24})", f(F6.REV), N, "Room revenue 2024")],
       extra=model)

    tp(m, "4.6.2",
       [("ADR that gives ₦1bn profit (algebra check)", f"=(1000000000+{A['Fixed cost per year']})/({A['Rooms']}*{A['Days']}*{A['Occupancy']})+{A['Variable cost per occupied night']}",
         f(F6.ADR_1BN), N, "What Goal Seek should return."),
        ("Increase on today's ADR", f"=(1000000000+{A['Fixed cost per year']})/({A['Rooms']}*{A['Days']}*{A['Occupancy']})+{A['Variable cost per occupied night']}-{A['ADR (average daily rate)']}",
         f(F6.ADR_1BN - ADR0), N, "")],
       [("n", "What ADR gives ₦2bn profit at the same occupancy?",
         f"=(2000000000+{A['Fixed cost per year']})/({A['Rooms']}*{A['Days']}*{A['Occupancy']})+{A['Variable cost per occupied night']}",
         f((2e9 + FC) / (R * DAYS * OCC0) + VC), N, "ADR for ₦2bn")],
       steps=["Data › What-If Analysis › Goal Seek.", "Set cell: the profit result in 4.6.1. To value: 1000000000. By changing: the yellow ADR cell.",
              "OK; read the new ADR, then Cancel (or Ctrl+Z) to restore the input."])

    tp(m, "4.6.3",
       [(f"Occupancy needed for ₦{t / 1e9:.0f}bn", f"=({t:.0f}+{A['Fixed cost per year']})/({A['Rooms']}*{A['Days']}*({A['ADR (average daily rate)']}-{A['Variable cost per occupied night']}))",
         f(F6.OCC_T[t]), "pct", "Above 100% is impossible: Goal Seek does not know." if F6.OCC_T[t] > 1 else "") for t in (1e9, 3e9, 6e9)],
       [("n", "What occupancy breaks even (profit 0)?", f"={A['Fixed cost per year']}/({A['Rooms']}*{A['Days']}*({A['ADR (average daily rate)']}-{A['Variable cost per occupied night']}))",
         f(FC / (R * DAYS * (ADR0 - VC))), "pct", "Break-even occupancy"),
        ("t", "Goal Seek returns 109% occupancy for ₦6bn. Is the target achievable? (Yes or No)", "No")])

    occs = F6.OCCS

    def onevar(mm):
        df = pd.DataFrame({"Occupancy": occs})
        calc = {"Profit_at_100k": (f"={A['Rooms']}*{A['Days']}*[@Occupancy]*(100000-{A['Variable cost per occupied night']})-{A['Fixed cost per year']}", N,
                                   [f(F6.profit(o, 100_000)) for o in occs])}
        mm.h("tOcc: profit at ADR ₦100,000 for each occupancy (the formula version of a one-variable data table)")
        T["o"] = mm.s.table(df, "tOcc", fmts={"Occupancy": "pct0"}, calc=calc, widths={"Occupancy": 11, "Profit_at_100k": 16})
        mm.s.r += 1

    T = {}
    tp(m, "4.6.4",
       [("Profit at 40% occupancy", "=INDEX(tOcc[Profit_at_100k],1)", f(F6.profit(.4, 1e5)), N, ""),
        ("Profit gained per 5 points of occupancy", "=INDEX(tOcc[Profit_at_100k],2)-INDEX(tOcc[Profit_at_100k],1)", f(F6.profit(.45, 1e5) - F6.profit(.4, 1e5)), N,
         "Straight line: the model is linear in occupancy.")],
       [("n", "What is the profit at 55% occupancy and ADR ₦100,000?", "=INDEX(tOcc[Profit_at_100k],4)", f(F6.profit(.55, 1e5)), N, "Profit at 55%")],
       do=[("Type 40%-70% down a column (steps of 5%). In the cell above and right of the first value, type =the 4.6.1 profit cell. "
            "Select the block › Data › What-If Analysis › Data Table › Column input cell: the yellow Occupancy cell.",
            "Excel fills {=TABLE(,occupancy)}; with ADR set to 100,000 the values equal tOcc.")], extra=onevar)

    adrs = F6.ADRS
    occ4 = [0.4, 0.5, 0.6, 0.7]

    def twovar(mm):
        mm.h("Profit by occupancy (rows) and ADR (columns): formula version of a two-variable data table")
        ws, r0 = mm.s.ws, mm.s.r
        hdr = mm.s.F["hdr"]
        ws.write(r0, 1, "Occupancy \\ ADR", hdr)
        for j, a in enumerate(adrs):
            ws.write(r0, 2 + j, a, b.nf(N, style="input"))
        for i, o in enumerate(occ4):
            ws.write(r0 + 1 + i, 1, o, b.nf("pct0", style="input"))
            for j, a in enumerate(adrs):
                oc = f"$B${r0 + 2 + i}"
                ad = f"{chr(67 + j)}${r0 + 1}"
                ws.write_formula(r0 + 1 + i, 2 + j, "=" + P(occ=oc, adr=ad), b.nf(N), f(F6.profit(o, a)))
        T["g"] = (r0, len(occ4))
        mm.s.r = r0 + len(occ4) + 2

    def cell(i, j):
        r0, _ = T["g"]
        return f"{chr(67 + j)}{r0 + 2 + i}"
    tp(m, "4.6.5", None, [], extra=twovar)
    m.worked([("Profit at 50% and ₦120k", f"={cell(1, 2)}", f(F6.profit(.5, 120_000)), N, ""),
              ("Profit at 70% and ₦80k", f"={cell(3, 0)}", f(F6.profit(.7, 80_000)), N, "")])
    m.turn()
    m.ex("What is the profit at 60% occupancy and ADR ₦80,000?", f"={cell(2, 0)}", f(F6.profit(.6, 80_000)), N, short="Profit 60%/₦80k")
    m.do([("Build the same grid as a real Data Table: Row input cell = ADR, Column input cell = Occupancy.", "The values match the formula grid above.")])

    be = {a: FC / (R * DAYS * (a - VC)) for a in adrs}
    tp(m, "4.6.6",
       [("Break-even occupancy at ₦80k", f"={A['Fixed cost per year']}/({A['Rooms']}*{A['Days']}*(80000-{A['Variable cost per occupied night']}))", f(be[80000]), "pct",
         "Any grid row above this is profitable."),
        ("Break-even occupancy at ₦120k", f"={A['Fixed cost per year']}/({A['Rooms']}*{A['Days']}*(120000-{A['Variable cost per occupied night']}))", f(be[120000]), "pct", "")],
       [("n", "What is the break-even occupancy at ADR ₦100,000?", f"={A['Fixed cost per year']}/({A['Rooms']}*{A['Days']}*(100000-{A['Variable cost per occupied night']}))",
         f(be[100000]), "pct", "Break-even at ₦100k")])

    sc = F6.SCEN

    def scen(mm):
        df = pd.DataFrame({"Scenario": list(sc), "Occupancy": [round(v[0], 4) for v in sc.values()], "ADR": [round(v[1], 2) for v in sc.values()],
                           "Fixed_Cost": [round(v[2], 0) for v in sc.values()]})
        calc = {"Profit": (f"={A['Rooms']}*{A['Days']}*[@Occupancy]*([@ADR]-{A['Variable cost per occupied night']})-[@Fixed_Cost]", N,
                           [f(F6.profit(round(o, 4), round(a, 2), round(fx, 0))) for o, a, fx in sc.values()])}
        mm.h("tScen: three scenarios (the inputs Scenario Manager would store)")
        mm.s.table(df, "tScen", fmts={"Occupancy": "pct", "ADR": N, "Fixed_Cost": N}, calc=calc,
                   widths={"Scenario": 13, "Occupancy": 11, "ADR": 12, "Fixed_Cost": 15, "Profit": 16})
        mm.s.r += 1
    SP = {k: F6.profit(round(o, 4), round(a, 2), round(fx, 0)) for k, (o, a, fx) in sc.items()}
    tp(m, "4.6.7",
       [("Pessimistic profit", '=XLOOKUP("Pessimistic",tScen[Scenario],tScen[Profit])', f(SP["Pessimistic"]), N, ""),
        ("Base profit", '=XLOOKUP("Base",tScen[Scenario],tScen[Profit])', f(SP["Base"]), N, "")],
       [("n", "What is the optimistic profit?", '=XLOOKUP("Optimistic",tScen[Scenario],tScen[Profit])', f(SP["Optimistic"]), N, "Optimistic profit")],
       do=[("Data › What-If Analysis › Scenario Manager › Add three scenarios, each changing the yellow Occupancy, ADR and Fixed cost cells to the tScen values.",
            "Show switches the model to each scenario; the profit cell matches tScen.")], extra=scen)

    tp(m, "4.6.8",
       [("Worst case", "=MIN(tScen[Profit])", f(min(SP.values())), N, ""), ("Best case", "=MAX(tScen[Profit])", f(max(SP.values())), N, "")],
       [("n", "What is the swing between the best and worst case?", "=MAX(tScen[Profit])-MIN(tScen[Profit])", f(max(SP.values()) - min(SP.values())), N, "Swing")],
       do=[("Scenario Manager › Summary › Result cells: the profit cell.", "A Scenario Summary sheet with the three scenarios side by side.")])

    SG, COMM, LO, HI, AL = F6.SG, F6.COMM, F6.LO, F6.HI, F6.ALLOC
    segs = F6.SEGS
    yr = ',tHotel[Check_In_Year],2024'

    def solver(mm):
        df = pd.DataFrame({"Segment": segs, "Commission": [COMM[s] for s in segs], "Min_Rooms": [LO[s] for s in segs], "Max_Rooms": [HI[s] for s in segs],
                           "Rooms": [int(AL[s]) for s in segs]})
        calc = {"ADR": (f"=AVERAGEIFS(tHotel[Room_Rate_NGN],tHotel[Sales_Channel],[@Segment]{yr})", N, [f(SG.adr[s]) for s in segs]),
                "Cancel_Rate": (f"=AVERAGEIFS(tHotel[Is_Cancelled],tHotel[Sales_Channel],[@Segment]{yr})", "pct", [f(SG.cr[s]) for s in segs]),
                "Net_per_Room": ("=[@ADR]*(1-[@Cancel_Rate])*(1-[@Commission])", N, [f(SG.net[s]) for s in segs]),
                "Value": ("=[@Net_per_Room]*[@Rooms]", N, [f(SG.net[s] * AL[s]) for s in segs])}
        mm.h("tSolver: Rooms (yellow-bordered) are the changing cells; they hold Solver's answer")
        T["v"] = mm.s.table(df, "tSolver", fmts={"Commission": "pct0", "Min_Rooms": "int", "Max_Rooms": "int", "Rooms": "int"}, calc=calc,
                            widths={"Segment": 22, "Commission": 11, "Min_Rooms": 10, "Max_Rooms": 10, "Rooms": 8, "ADR": 11, "Cancel_Rate": 11,
                                    "Net_per_Room": 13, "Value": 14})
        mm.s.r += 1

    tp(m, "4.6.9",
       [("Rooms allocated", "=SUM(tSolver[Rooms])", int(sum(AL.values())), "int", "Constraint: at most 320."),
        ("OTA share of rooms", '=XLOOKUP("Online Travel Agent",tSolver[Segment],tSolver[Rooms])/320', AL["Online Travel Agent"] / 320, "pct", "Constraint: at most 50%.")],
       [("t", "In Solver, the Rooms cells are called the … cells.", "changing", None, None, ("variable", "decision"))],
       steps=["File › Options › Add-ins › Manage Excel Add-ins › Go › tick Solver Add-in.", "Data › Solver appears at the right of the Data tab."],
       extra=solver)

    tp(m, "4.6.10",
       [(f"Net value per room: {s}", f'=XLOOKUP("{s}",tSolver[Segment],tSolver[Net_per_Room])', f(SG.net[s]), N, "") for s in ["Airline Crew", "Corporate"]],
       [("n", "What is the net value per room for Online Travel Agent bookings?", '=XLOOKUP("Online Travel Agent",tSolver[Segment],tSolver[Net_per_Room])',
         f(SG.net["Online Travel Agent"]), N, "OTA net per room"),
        ("n", "What is the 2024 cancellation rate of Walk-in / Travel Agent bookings?", f'=AVERAGEIFS(tHotel[Is_Cancelled],tHotel[Sales_Channel],"Walk-in / Travel Agent"{yr})',
         f(SG.cr["Walk-in / Travel Agent"]), "pct", "Walk-in cancel rate")])

    tp(m, "4.6.11",
       [("Objective: expected net value per night", "=SUM(tSolver[Value])", f(F6.LPV), N, "Solver maximises this.")],
       [("n", "How many rooms go to Walk-in / Travel Agent?", '=XLOOKUP("Walk-in / Travel Agent",tSolver[Segment],tSolver[Rooms])', int(AL["Walk-in / Travel Agent"]), "int",
         "Walk-in rooms"),
        ("o", "Why does Corporate get only 70 rooms when it has the second-highest net value?", "Its demand cap is 70: Solver gives it every room it can take.")],
       do=[("Data › Solver. Objective: the 4.6.11 result cell, Max. Changing: tSolver[Rooms]. Constraints: Rooms >= Min_Rooms, Rooms <= Max_Rooms, "
            "SUM(Rooms) <= 320, OTA rooms <= 160, Rooms integer. Method: Simplex LP › Solve.",
            f"Solver keeps Corporate {AL['Corporate']}, Airline Crew {AL['Airline Crew']}, OTA {AL['Online Travel Agent']}, Walk-in {AL['Walk-in / Travel Agent']}.")])

    tp(m, "4.6.12", None,
       [("o", "Write the four-line documentation block (inputs, assumptions, tools, limits) for this sheet.",
         "Inputs: 2024 stayed bookings from tHotel. Assumptions: 320 rooms, ₦28,000 variable and ₦3.6bn fixed cost, scenario values. "
         "Tools: Goal Seek, data tables, Scenario Manager, Solver (Simplex LP). Limits: annual averages, no seasonality, demand caps assumed.")])
    m.finish()


# =====================================================================================================
def m47(b, D):
    D25, STO = D.D25, D.STO
    g = D25.groupby("Store_Name").agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"))
    st = STO.set_index("Store_Name").join(g)
    st["Att"] = st.Sales / st.Annual_Target_2025
    st["Gap"] = st.Sales - st.Annual_Target_2025
    st = st.sort_values("Att")
    wg = st.Gap.idxmin()          # furthest behind = largest gap to target (the slides' WORST)
    tot = float(D25.Sales.sum())
    y24 = float(D.MON[D.MON.Year == 2024].Sales.sum())
    m = ModuleSheet(b, "4.7 Storytelling", "4.7  Data Storytelling",
                    "Tutor walk-through after Module 4.7: the FY2025 SuperMart story, from numbers to a headline, a hierarchy and a finding title",
                    "4.7.1 - 4.7.9")
    m.dataset([("tDash25 / tStores", "Sheets Data_Dash25, Data_Stores (A1)", "FY2025 sales lines and store targets."),
               ("tMonthly", "Sheet Data_Monthly (A1)", "FY2024 total, for the growth figure."),
               ("tAttain, tZone", "On this sheet", "Store attainment and zone sales: every cell is a formula.")])
    T = {}

    def att(mm):
        df = pd.DataFrame({"Store_Name": st.index})
        calc = {"Sales": ("=SUMIF(tDash25[Store_Name],[@Store_Name],tDash25[Sales])", N, [f(x) for x in st.Sales]),
                "Target": ("=XLOOKUP([@Store_Name],tStores[Store_Name],tStores[Annual_Target_2025])", N, [f(x) for x in st.Annual_Target_2025]),
                "Attainment": ("=[@Sales]/[@Target]", "pct", [f(x) for x in st.Att]),
                "Gap": ("=[@Sales]-[@Target]", N, [f(x) for x in st.Gap]),
                "Profit": ("=SUMIF(tDash25[Store_Name],[@Store_Name],tDash25[Profit])", N, [f(x) for x in st.Profit])}
        mm.h("tAttain: FY2025 sales against target, lowest attainment first")
        T["a"] = mm.s.table(df, "tAttain", calc=calc, widths={"Store_Name": 24, "Sales": 14, "Target": 14, "Attainment": 11, "Gap": 14, "Profit": 14})
        mm.s.r += 1
        red = [{"fill": {"color": "#B23A48"}} if a < 1 else {"fill": {"color": "#9BB7C9"}} for a in st.Att]
        mm.s.chart("bar", [{"name": "Attainment", "categories": T["a"].ref("Store_Name"), "values": T["a"].ref("Attainment"), "points": red}],
                   f"{int((st.Att < 1).sum())} of 12 stores missed target; {wg} is furthest behind", "", "Attainment", legend=False,
                   y_fmt="0%", rows=17, size=(640, 310))

    nb = int((st.Att < 1).sum())
    tp(m, "4.7.1",
       [("FY2025 growth", f"=SUM(tDash25[Sales])/SUMIF(tMonthly[Year],2024,tMonthly[Sales])-1", f(tot / y24 - 1), "pct", "The headline number."),
        ("Stores below target", '=COUNTIF(tAttain[Attainment],"<1")', nb, "int", "The complication.")],
       [("n", "How many stores made a loss in FY2025?", '=COUNTIF(tAttain[Profit],"<0")', int((st.Profit < 0).sum()), "int", "Loss-making stores")],
       extra=att)

    tp(m, "4.7.2",
       [("Act 1: FY2025 sales", "=SUM(tDash25[Sales])", f(tot), N, "Situation."), ("Act 2: stores below target", '=COUNTIF(tAttain[Attainment],"<1")', nb, "int", "Complication."),
        ("Act 3: store with the largest gap", "=INDEX(tAttain[Store_Name],MATCH(MIN(tAttain[Gap]),tAttain[Gap],0))", st.Gap.idxmin(), "gen", "Where the resolution starts.")],
       [("o", "Write the three acts in one sentence each.",
         f"Sales grew {tot / y24 - 1:.0%} to ₦{tot / 1e6:.0f}m. But {nb} of 12 stores missed target and {int((st.Profit < 0).sum())} lost money. "
         f"Recovery plans for those stores, starting with {st.Gap.idxmin()}.")])

    tp(m, "4.7.3",
       [("Largest gap (₦)", "=MIN(tAttain[Gap])", f(st.Gap.min()), N, ""),
        ("Headline (built by formula)", '="FY2025 sales grew "&TEXT(SUM(tDash25[Sales])/SUMIF(tMonthly[Year],2024,tMonthly[Sales])-1,"0%")&", but "&COUNTIF(tAttain[Attainment],"<1")&" of 12 stores missed target"',
         f"FY2025 sales grew {tot / y24 - 1:.0%}, but {nb} of 12 stores missed target", "gen", "Updates if the data changes.")],
       [("t", "Which store has the largest gap to target?", st.Gap.idxmin(), "=INDEX(tAttain[Store_Name],MATCH(MIN(tAttain[Gap]),tAttain[Gap],0))", "Largest gap store")])

    zn = D25.groupby("Zone").Sales.sum().reindex(["South-West", "South-South", "North-Central", "South-East"])
    tp(m, "4.7.4",
       [("South-West share of sales", '=SUMIF(tDash25[Zone],"South-West",tDash25[Sales])/SUM(tDash25[Sales])', f(zn["South-West"] / tot), "pct", "Support point 1."),
        ("Biggest loss (₦)", "=MIN(tAttain[Profit])", f(st.Profit.min()), N, "Support point 2."),
        ("Store with the biggest loss", "=INDEX(tAttain[Store_Name],MATCH(MIN(tAttain[Profit]),tAttain[Profit],0))", st.Profit.idxmin(), "gen", "")],
       [("n", "What share of FY2025 sales came from the South-East?", '=SUMIF(tDash25[Zone],"South-East",tDash25[Sales])/SUM(tDash25[Sales])', f(zn["South-East"] / tot),
         "pct", "South-East share")])

    tp(m, "4.7.5", None,
       [("t", "Which audience needs the reconciliation to the ledger: board, zone managers or finance?", "finance"),
        ("t", "Which audience gets a one-slide summary with the decision needed?", "board", None, None, ("the board",))])

    tp(m, "4.7.6",
       [("Finding title (formula)", '=COUNTIF(tAttain[Attainment],"<1")&" of 12 stores missed target; "&INDEX(tAttain[Store_Name],MATCH(MIN(tAttain[Gap]),tAttain[Gap],0))&" is furthest behind"',
         f"{nb} of 12 stores missed target; {wg} is furthest behind", "gen", "The chart title states what the chart shows.")],
       [("o", "Rewrite the label title 'Sales by zone FY2025' as a finding.",
         f"South-West brought {zn['South-West'] / tot:.0%} of FY2025 sales, twice the South-East.")])

    tp(m, "4.7.7",
       [(f"Attainment of {wg}", "=INDEX(tAttain[Attainment],MATCH(MIN(tAttain[Gap]),tAttain[Gap],0))", f(st.Att[wg]), "pct", "The callout on the chart."),
        ("Bars to colour red (below 100%)", '=COUNTIF(tAttain[Attainment],"<1")', nb, "int", "")],
       [("n", "What is the attainment of the best store?", "=MAX(tAttain[Attainment])", f(st.Att.max()), "pct", "Best attainment")],
       do=[(f"On the attainment chart, add a data label to the {wg} bar only (click the bar twice, then + › Data Labels).",
            f"One label: {st.Att[wg]:.0%} on {wg}.")])

    AX = 120e6

    def zone(mm):
        df = pd.DataFrame({"Zone": list(zn.index)})
        calc = {"Sales": ("=SUMIF(tDash25[Zone],[@Zone],tDash25[Sales])", N, [f(x) for x in zn]),
                "Visible_from_120m": ("=[@Sales]-120000000", N, [f(x - AX) for x in zn])}
        mm.h("tZone: what an axis starting at ₦120m shows")
        t = mm.s.table(df, "tZone", calc=calc, widths={"Zone": 14, "Sales": 14, "Visible_from_120m": 18})
        mm.s.r += 1
        mm.s.chart("column", [{"name": "Sales", "categories": t.ref("Zone"), "values": t.ref("Sales")}],
                   "From ₦0: South-West is twice South-East, not 45 times", "", "", legend=False, y_fmt='₦#,##0,,"m"', rows=14, size=(520, 240),
                   extra=lambda ch: ch.set_y_axis({"min": 0, "num_format": '₦#,##0,,"m"', "num_font": {"size": 8}}))

    tp(m, "4.7.8",
       [("True ratio, South-West ÷ South-East", '=XLOOKUP("South-West",tZone[Zone],tZone[Sales])/XLOOKUP("South-East",tZone[Zone],tZone[Sales])',
         f(zn["South-West"] / zn["South-East"]), "dec1", ""),
        ("Ratio the ₦120m axis draws", '=XLOOKUP("South-West",tZone[Zone],tZone[Visible_from_120m])/XLOOKUP("South-East",tZone[Zone],tZone[Visible_from_120m])',
         f((zn["South-West"] - AX) / (zn["South-East"] - AX)), "dec1", "Wildly exaggerated.")],
       [("n", "With the ₦120m axis, how many times longer does South-South look than North-Central? (1 decimal)",
         '=XLOOKUP("South-South",tZone[Zone],tZone[Visible_from_120m])/XLOOKUP("North-Central",tZone[Zone],tZone[Visible_from_120m])',
         f((zn["South-South"] - AX) / (zn["North-Central"] - AX)), "dec1", "Exaggerated ratio")], extra=zone)

    tp(m, "4.7.9",
       [("Shortfall of the below-target stores", '=-SUMIF(tAttain[Gap],"<0")', f(-st.Gap[st.Gap < 0].sum()), N, "Paragraph 3: the cost against plan."),
        ("Best store", "=INDEX(tAttain[Store_Name],MATCH(MAX(tAttain[Attainment]),tAttain[Attainment],0))", st.Att.idxmax(), "gen", "")],
       [("o", "Write the four-paragraph summary (headline, evidence, implication, ask) in four sentences.",
         f"Sales grew {tot / y24 - 1:.0%} to ₦{tot / 1e6:.0f}m, but {nb} of 12 stores missed target. {st.Att.idxmax()} led at {st.Att.max():.0%}; "
         f"{wg} trailed at {st.Att[wg]:.0%}. The below-target stores fell ₦{-st.Gap[st.Gap < 0].sum() / 1e6:.1f}m short of plan. "
         "We ask the board to approve recovery plans for those stores.")])
    m.finish()


# =====================================================================================================
def m48(b, D):
    L, C = D.LOS, D.C23
    m = ModuleSheet(b, "4.8 Report Writing", "4.8  Report Writing and PDF Output",
                    "Tutor walk-through after Module 4.8: the Mercyland 2023 operations report, from the figures to a print-ready PDF", "4.8.1 - 4.8.10")
    m.dataset([("tLOS", "Sheet Data_LOS (A1)", f"{len(L):,} Mercyland admissions in 2023."),
               ("tCensus23", "Sheet Data_Census (A1)", "365 midnight bed counts, 200 staffed beds."),
               ("tDept", "On this sheet", "The report's department table: every cell is a formula.")])
    ws = m.s.ws
    ws.set_landscape()
    ws.set_paper(9)
    ws.fit_to_pages(1, 0)
    ws.set_footer("&LETEH Week 4 - Mercyland 2023 report (training data)&RPage &P of &N")
    dep = L.groupby("Department").agg(n=("Episode_ID", "size"), alos=("LOS_Days", "mean"), re=("Readmitted_30d", lambda x: (x == "Yes").mean())).sort_values("n", ascending=False)
    oc = C.Occupancy_Pct
    mon = pd.to_datetime(C.Census_Date).dt.month
    T = {}

    def dept(mm):
        df = pd.DataFrame({"Department": dep.index})
        calc = {"Admissions": ("=COUNTIF(tLOS[Department],[@Department])", "int", [int(x) for x in dep.n]),
                "ALOS_Days": ("=AVERAGEIF(tLOS[Department],[@Department],tLOS[LOS_Days])", "dec1", [f(x) for x in dep.alos]),
                "Readmit_Rate": ('=COUNTIFS(tLOS[Department],[@Department],tLOS[Readmitted_30d],"Yes")/[@Admissions]', "pct", [f(x) for x in dep.re])}
        mm.h("tDept: departments, 2023 (the report's main table)")
        T["d"] = mm.s.table(df, "tDept", calc=calc, widths={"Department": 24, "Admissions": 11, "ALOS_Days": 10, "Readmit_Rate": 12})
        mm.s.r += 1

    tp(m, "4.8.1",
       [("Admissions in the extract", "=ROWS(tLOS)", len(L), "int", "Quoted in the Method section."), ("Census days", "=ROWS(tCensus23)", len(C), "int", "")],
       [("t", "Which section holds the definitions and the full tables: findings or appendix?", "appendix")])

    q2 = int(((oc > .95) & mon.between(4, 6)).sum())
    tp(m, "4.8.2",
       [("Busiest department", "=INDEX(tDept[Department],MATCH(MAX(tDept[Admissions]),tDept[Admissions],0))", dep.n.idxmax(), "gen", "Finding 1."),
        ("Average occupancy", "=AVERAGE(tCensus23[Occupancy_Pct])", f(oc.mean()), "pct", "Finding 2."),
        ("Days above 95%", '=COUNTIF(tCensus23[Occupancy_Pct],">0.95")', int((oc > .95).sum()), "int", ""),
        ("Of which in April-June", '=COUNTIFS(tCensus23[Occupancy_Pct],">0.95",tCensus23[Census_Date],">="&DATE(2023,4,1),tCensus23[Census_Date],"<="&DATE(2023,6,30))',
         q2, "int", "Why the ask is Q2 surge capacity."),
        ("30-day readmission rate", '=COUNTIF(tLOS[Readmitted_30d],"Yes")/ROWS(tLOS)', f((L.Readmitted_30d == "Yes").mean()), "pct", "Finding 3.")],
       [("t", "Which department has the highest readmission rate?", dep.re.idxmax(), "=INDEX(tDept[Department],MATCH(MAX(tDept[Readmit_Rate]),tDept[Readmit_Rate],0))",
         "Highest readmission"),
        ("n", "What is that department's readmission rate?", "=MAX(tDept[Readmit_Rate])", f(dep.re.max()), "pct", "Highest readmission rate")], extra=dept)

    tp(m, "4.8.3", None, [],
       do=[("Format tDept for the report: one font, 2-3 sizes, left-aligned text, right-aligned numbers, one accent colour for the highest readmission.",
            "A calm table where the eye goes straight to the highlighted row.")])

    def bar(mm):
        mm.s.chart("bar", [{"name": "Readmission rate", "categories": T["d"].ref("Department"), "values": T["d"].ref("Readmit_Rate"),
                            "points": [{"fill": {"color": "#B23A48"}} if d == dep.re.idxmax() else {"fill": {"color": "#9BB7C9"}} for d in dep.index]}],
                   f"{dep.re.idxmax()} readmits {dep.re.max():.0%} within 30 days, the highest rate", "", "", legend=False, y_fmt="0%", rows=14, size=(560, 240))
    tp(m, "4.8.4",
       [("Surgery ALOS (days)", '=XLOOKUP("Surgery",tDept[Department],tDept[ALOS_Days])', f(dep.alos["Surgery"]), "dec1", ""),
        ("Total of the table", "=SUM(tDept[Admissions])", len(L), "int", "Must equal the extract rows.")],
       [("n", "What is the ALOS of Obstetrics & Gynaecology? (1 decimal)", '=XLOOKUP("Obstetrics & Gynaecology",tDept[Department],tDept[ALOS_Days])',
         f(dep.alos["Obstetrics & Gynaecology"]), "dec1", "O&G ALOS")], extra=bar)

    tp(m, "4.8.5", None, [],
       do=[("Open Page Layout on this sheet: it is already Landscape, A4, fit to 1 page wide, with a footer. Check File › Print.",
            "Footer 'Page 1 of N'; every column fits the page width."),
           ("Page Layout › Print Titles › Rows to repeat at top: the title rows.", "The title repeats on every printed page.")])
    tp(m, "4.8.6", None, [],
       do=[("File › Export › Create PDF/XPS › Options: Active sheet(s); Optimise for Standard › Publish.", "A PDF of this sheet only, no raw data sheets."),
           ("Open the PDF and zoom to 200% on the chart.", "Sharp lines and text: the chart was exported as vectors.")])
    tp(m, "4.8.7", None,
       [("t", "The ₦ sign shows as a box in the PDF. What do you check first: the font or the data?", "font", None, None, ("the font",))])
    tp(m, "4.8.8", None, [],
       do=[("File › Info › Check for Issues › Inspect Document.", "A list of comments, hidden sheets and personal information to remove or keep."),
           ("File › Info › Properties: set Title and Author.", "The PDF shows the title in its window bar.")])
    tp(m, "4.8.9", None,
       [("t", "A version number that marks an approved report: v0.2 or v1.0?", "v1.0", None, None, ("1.0",))])
    tp(m, "4.8.10",
       [("Check: table total = extract rows", "=IF(SUM(tDept[Admissions])=ROWS(tLOS),\"OK\",\"CHECK\")", "OK", "gen", ""),
        ("Check: occupancy in the summary = census average", "=IF(ROUND(AVERAGE(tCensus23[Occupancy_Pct]),3)=ROUND(SUMPRODUCT(tCensus23[Inpatients])/(200*ROWS(tCensus23)),3),\"OK\",\"CHECK\")",
         "OK", "gen", "Two routes to the same number.")],
       [("n", "What average occupancy do 365 census days at 200 beds give (inpatients ÷ bed-days)?", "=SUM(tCensus23[Inpatients])/(200*ROWS(tCensus23))",
         f(C.Inpatients.sum() / (200 * len(C))), "pct", "Occupancy from bed-days")])
    m.finish()


# =====================================================================================================
def m49(b, D):
    H = D.H
    m = ModuleSheet(b, "4.9 Presentations", "4.9  Slide Presentations",
                    "Tutor walk-through after Module 4.9: the Atlantic Pearl cancellations talk, from proof numbers to a projected chart", "4.9.1 - 4.9.10")
    m.dataset([("tHotel", "Sheet Data_Hotel (A1)", f"{len(H):,} reservations, 2023-24, with Is_Cancelled."),
               ("tChan, tStory", "On this sheet", "Cancellation by sales channel (slide 3) and the storyboard timing.")])
    canv = float(H.Reservation_Value_NGN[H.Is_Cancelled == 1].sum())
    seg = H.groupby("Sales_Channel").agg(n=("Is_Cancelled", "size"), r=("Is_Cancelled", "mean")).sort_values("n", ascending=False)
    LB = [(0, 7), (8, 30), (31, 90), (91, 180), (181, 99999)]
    lead = [float(H.Is_Cancelled[(H.Days_Ahead >= a) & (H.Days_Ahead <= b_)].mean()) for a, b_ in LB]
    tp(m, "4.9.1",
       [("Value of cancelled bookings", '=SUMIF(tHotel[Reservation_Status],"Cancelled",tHotel[Reservation_Value_NGN])', f(canv), N, "Slide 1 headline."),
        ("Bookings in the talk", "=ROWS(tHotel)", len(H), "int", "")],
       [("n", "What share of total booking value was cancelled?", '=SUMIF(tHotel[Reservation_Status],"Cancelled",tHotel[Reservation_Value_NGN])/SUM(tHotel[Reservation_Value_NGN])',
         f(canv / H.Reservation_Value_NGN.sum()), "pct", "Cancelled value share")])

    tp(m, "4.9.2",
       [("Proof: overall cancel rate", "=AVERAGE(tHotel[Is_Cancelled])", f(H.Is_Cancelled.mean()), "pct", "Slide 2: one booking in three."),
        ("Proof: OTA cancel rate", '=AVERAGEIF(tHotel[Sales_Channel],"Online Travel Agent",tHotel[Is_Cancelled])', f(seg.r["Online Travel Agent"]), "pct", "Slide 3."),
        ("Proof: 181+ days ahead", '=AVERAGEIF(tHotel[Days_Ahead],">=181",tHotel[Is_Cancelled])', lead[4], "pct", "Slide 4.")],
       [("n", "What is the Corporate cancellation rate?", '=AVERAGEIF(tHotel[Sales_Channel],"Corporate",tHotel[Is_Cancelled])', f(seg.r["Corporate"]), "pct",
         "Corporate cancel rate")])

    tp(m, "4.9.3",
       [("Cancel rate, 0-7 days ahead", '=AVERAGEIFS(tHotel[Is_Cancelled],tHotel[Days_Ahead],">=0",tHotel[Days_Ahead],"<=7")', lead[0], "pct", "Grey bar."),
        ("Cancel rate, 181+ days ahead", '=AVERAGEIF(tHotel[Days_Ahead],">=181",tHotel[Is_Cancelled])', lead[4], "pct", "The one teal bar.")],
       [("n", "What is the cancel rate for 91-180 days ahead?", '=AVERAGEIFS(tHotel[Is_Cancelled],tHotel[Days_Ahead],">=91",tHotel[Days_Ahead],"<=180")', lead[3], "pct",
         "Cancel rate 91-180")],
       do=[("Copy the lead-time chart from 4.3 Correlation. Make all bars grey except 181+ (teal), font 18 pt or larger, delete gridlines.",
            "A chart readable from the back of the room: one bar stands out.")])

    t4 = f"Bookings made 6+ months ahead cancel {lead[4]:.0%} of the time"

    def title(mm):
        mm.h("Slide 4 title (type a new title in the yellow cell)")
        A = mm.s.inputs([("Slide 4 title", t4, "gen", "Aim: under 12 words.")])
        T["t"] = A["Slide 4 title"]
    T = {}
    tp(m, "4.9.4", lambda mm: [("Words", f'=LEN(TRIM({T["t"]}))-LEN(SUBSTITUTE(TRIM({T["t"]})," ",""))+1', len(t4.split()), "int", "Words = spaces + 1."),
                               ("Characters", f'=LEN({T["t"]})', len(t4), "int", "")],
       [("n", "How many words are in 'One booking in three is cancelled'? (use the same formula)",
         '=LEN(TRIM("One booking in three is cancelled"))-LEN(SUBSTITUTE(TRIM("One booking in three is cancelled")," ",""))+1', 6, "int", "Word count")],
       extra=title)

    tp(m, "4.9.5", None,
       [("t", "How many accent colours should a deck use to point at the finding?", "1", None, None, ("one",))])

    def chan(mm):
        df = pd.DataFrame({"Sales_Channel": seg.index})
        calc = {"Bookings": ("=COUNTIF(tHotel[Sales_Channel],[@Sales_Channel])", "int", [int(x) for x in seg.n]),
                "Cancel_Rate": ("=AVERAGEIF(tHotel[Sales_Channel],[@Sales_Channel],tHotel[Is_Cancelled])", "pct", [f(x) for x in seg.r])}
        mm.h("tChan: cancellation by sales channel (the slide 3 table)")
        mm.s.table(df, "tChan", calc=calc, widths={"Sales_Channel": 24, "Bookings": 10, "Cancel_Rate": 12})
        mm.s.r += 1
    tp(m, "4.9.6",
       [("Channels", "=ROWS(tChan)", len(seg), "int", "Five rows: small enough for a slide."),
        ("Highest rate", "=MAX(tChan[Cancel_Rate])", f(seg.r.max()), "pct", "")],
       [("n", "How many bookings came through Walk-in / Travel Agent?", '=XLOOKUP("Walk-in / Travel Agent",tChan[Sales_Channel],tChan[Bookings])',
         int(seg.n["Walk-in / Travel Agent"]), "int", "Walk-in bookings")], extra=chan)

    tp(m, "4.9.7", None,
       [("o", "Prepare the answer to 'Is the lead-time link causal?'", "No: it is an association. Early bookers have more time for plans to change; we have no reason field.")])

    story = pd.DataFrame({"Slides": ["1", "2-6", "7-9", "10"], "Role": ["Headline", "Evidence", "Options and recommendation", "Decision needed"],
                          "Minutes": [0.5, 6, 3.5, 1]})

    def board(mm):
        mm.h("tStory: the storyboard")
        mm.s.table(story, "tStory", fmts={"Minutes": "dec1"}, widths={"Slides": 8, "Role": 26, "Minutes": 9})
        mm.s.r += 1
    tp(m, "4.9.8",
       [("Planned talk (minutes)", "=SUM(tStory[Minutes])", 11.0, "dec1", "Slot: 12 minutes + 5 for questions."),
        ("Spare minutes", "=12-SUM(tStory[Minutes])", 1.0, "dec1", "")],
       [("n", "If the evidence section runs 2 minutes over, how many minutes over the slot is the talk?", "=SUM(tStory[Minutes])+2-12", 1.0, "dec1", "Minutes over")],
       extra=board)

    tp(m, "4.9.9", None,
       [("t", "Meaning shown only by colour fails for colour-blind viewers. What do you add?", "value labels", None, None, ("labels", "data labels"))])
    tp(m, "4.9.10", None,
       [("o", "Write the closing sentence: the decision you ask the board for.",
         "We ask the board to approve a deposit policy for bookings made 91 or more days ahead.")])
    m.finish()
