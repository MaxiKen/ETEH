"""Week 3 workbook, modules 3.7 - 3.10 (edition 3): chart anatomy and the 17 chart types, dashboard development, three
dashboard blueprints, and file extracts into SQL, DAX and dashboards.

3.7 builds small summary tables with live formulas (tZoneSum, tMonth, tTier, tPay, tDept, tBins, tStoreMonth) and draws real
Excel charts from them for the nine types an .xlsx written by XlsxWriter can carry (column, bar, line, area, pie, doughnut,
scatter, radar, combo) plus sparklines; bubble and the Excel 2016 types (waterfall, funnel, treemap, sunburst, histogram, box) are
built by the learner from the prepared data (a column chart of tBins stands in for the histogram). 3.8 has a small dashboard driven by a zone selector cell. 3.10 profiles tMsDirty, the raw extract.
"""
from __future__ import annotations

import pandas as pd

from w1x import ModuleSheet
from wb3_common import N, f, net, netn, tp

XL0 = pd.Timestamp("1899-12-30")
MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
ZONES = ["South-West", "South-South", "North-Central", "South-East"]
DASH_DS = ("tDash", "Sheet Data_Dash (A1)", "3,312 SuperMart Stores sales lines, FY2025: the dashboard data.")
STORES_DS = ("tDashStores", "Sheet Data_DashStores (A1)", "12 stores with FY2025 sales, profit, target and attainment.")


def mdates(mn):
    return f'tDash[Order_Date],">="&DATE(2025,{mn},1),tDash[Order_Date],"<="&EOMONTH(DATE(2025,{mn},1),0)'


def build(b, D):
    t = m37(b, D)
    m38(b, D)
    m39(b, D)
    m310(b, D)
    return t


# =====================================================================================================
def m37(b, D):
    DS, DSS, C, TX, L, NT, CV, A = D.DS, D.DSS, D.C, D.TX, D.L, D.NT, D.CV, D.A
    m = ModuleSheet(b, "3.7 Charts", "3.7  Chart Anatomy and the 17 Excel Chart Types",
                    "Tutor walk-through after Module 3.7: summary tables with live formulas, a finished chart for nine types, "
                    "and data plus steps for the rest", "3.7.1 - 3.7.26")
    m.dataset([DASH_DS, STORES_DS, ("tCustomers / tTransactions / tCV / tAdm23", "Data_ sheets (A1)", "Tier, payment, scatter, histogram and box data."),
               ("tZoneSum, tMonth, tTier, tPay, tDept, tBins, tStoreMonth", "On this sheet", "Chart source tables: every cell is a formula.")])
    zs = DS.groupby("Zone").agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"), Lines=("Sales", "size")).reindex(ZONES)
    tg = DSS.groupby("Zone").Target.sum().reindex(ZONES)
    zs["Margin"] = zs.Profit / zs.Sales
    zs["Attainment"] = zs.Sales / tg
    T = {}

    def zone_table(mm):
        df = pd.DataFrame({"Zone": ZONES})
        calc = {"Sales": ("SUMIF(tDash[Zone],[@Zone],tDash[Sales])", N, [f(x) for x in zs.Sales]),
                "Profit": ("SUMIF(tDash[Zone],[@Zone],tDash[Profit])", N, [f(x) for x in zs.Profit]),
                "Lines": ("COUNTIF(tDash[Zone],[@Zone])", "int", [int(x) for x in zs.Lines]),
                "Margin": ("[@Profit]/[@Sales]", "pct", [f(x) for x in zs.Margin]),
                "Attainment": ("[@Sales]/SUMIF(tDashStores[Zone],[@Zone],tDashStores[Target])", "pct", [f(x) for x in zs.Attainment])}
        mm.h("tZoneSum: FY2025 by zone (source of the zone charts)")
        T["z"] = mm.s.table(df, "tZoneSum", calc=calc, widths={"Zone": 14, "Sales": 14, "Profit": 13, "Lines": 8, "Margin": 9, "Attainment": 11})
        mm.s.r += 1

    ms = DS.groupby(DS.Order_Date.dt.month).agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"))

    def month_table(mm):
        df = pd.DataFrame({"Month_No": range(1, 13), "Month": MON})
        calc = {"Sales": (f"SUMIFS(tDash[Sales],{mdates('[@Month_No]')})", N, [f(x) for x in ms.Sales]),
                "Profit": (f"SUMIFS(tDash[Profit],{mdates('[@Month_No]')})", N, [f(x) for x in ms.Profit]),
                "YTD": ("SUMIFS([Sales],[Month_No],\"<=\"&[@Month_No])", N, [f(x) for x in ms.Sales.cumsum()])}
        mm.h("tMonth: FY2025 by month (source of the line and area charts)")
        T["m"] = mm.s.table(df, "tMonth", fmts={"Month_No": "gen"}, calc=calc, widths={"Month_No": 10, "Month": 8, "Sales": 14, "Profit": 13, "YTD": 14})
        mm.s.r += 1

    TOT = f(DS.Sales.sum())
    tp(m, "3.7.1",
       [("Leading zone's sales", "=MAX(tZoneSum[Sales])", f(zs.Sales.max()), N, "The finding the title should state."),
        ("Leading zone", "=INDEX(tZoneSum[Zone],MATCH(MAX(tZoneSum[Sales]),tZoneSum[Sales],0))", zs.Sales.idxmax(), "gen", "")],
       [("t", "Which zone had the lowest FY2025 sales?", zs.Sales.idxmin(), "=INDEX(tZoneSum[Zone],MATCH(MIN(tZoneSum[Sales]),tZoneSum[Sales],0))", "Lowest zone"),
        ("o", "Write a title for the zone chart that states the finding.", f"South-West led FY2025 with ₦{zs.Sales.max() / 1e6:.1f}m of sales.")],
       steps=["Decide the one question the chart answers.", "Pick the chart that answers it (3.7.9).", "Remove 3D, shading, extra colours; write a title that states the finding."],
       extra=zone_table)
    tp(m, "3.7.2",
       [("Series points (zones)", "=ROWS(tZoneSum)", 4, "int", "One column per zone."),
        ("Value axis maximum needed", "=CEILING(MAX(tZoneSum[Sales]),50000000)", f(-(-zs.Sales.max() // 5e7) * 5e7), N, "Round up to a tidy step.")],
       [("n", "Total of the series (sum of zone sales)?", "=SUM(tZoneSum[Sales])", TOT, N, "Sum of the zone series")],
       steps=["Click a part of the chart; read its name in the Format pane (Ctrl+1).", "Chart Elements (+ button) adds or removes titles, labels, legend, gridlines.",
              "Chart Styles (brush) only for quick looks; format parts yourself."], extra=month_table)
    tp(m, "3.7.3",
       [("Months on the line chart", "=ROWS(tMonth)", 12, "int", "Plot area holds 12 points."),
        ("Largest month", "=MAX(tMonth[Sales])", f(ms.Sales.max()), N, "")],
       [("t", "To make room for long axis labels without resizing the chart, which area do you shrink?", "plot area", None, None, ("the plot area",))],
       steps=["Click the chart's edge: chart area (move, size, font).", "Click inside, between gridlines: plot area (resize for labels).",
              "Set the font once on the chart area."])
    sw, se = zs.Sales["South-West"], zs.Sales["South-East"]
    tp(m, "3.7.4",
       [("South-West ÷ South-East (truth)", '=XLOOKUP("South-West",tZoneSum[Zone],tZoneSum[Sales])/XLOOKUP("South-East",tZoneSum[Zone],tZoneSum[Sales])',
         sw / se, "dec1", "Bars from ₦0 show this."),
        ("Ratio drawn with the axis starting at ₦100m", '=(XLOOKUP("South-West",tZoneSum[Zone],tZoneSum[Sales])-100000000)/(XLOOKUP("South-East",tZoneSum[Zone],tZoneSum[Sales])-100000000)',
         (sw - 1e8) / (se - 1e8), "dec1", "Exaggerated.")],
       [("n", "With the axis starting at ₦100m, how many times longer does South-South's bar look than South-East's? (1 decimal)",
         '=(XLOOKUP("South-South",tZoneSum[Zone],tZoneSum[Sales])-100000000)/(XLOOKUP("South-East",tZoneSum[Zone],tZoneSum[Sales])-100000000)',
         (zs.Sales["South-South"] - 1e8) / (se - 1e8), "dec1", "Exaggerated ratio")],
       steps=["Double-click the value axis › Axis Options: Minimum 0.", "Display units: Millions; tick 'Show display units label'.",
              "Category axis: check the order of the categories."])
    tp(m, "3.7.5",
       [("Gridline step for the zone chart", "=50000000", 50000000, N, "Major unit ₦50m: about five gridlines.")],
       [("n", "With a major unit of ₦50m and a maximum of ₦300m, how many major gridlines above zero?", "=300000000/50000000", 6, "int", "Gridlines")],
       steps=["Click a gridline › Delete removes minor gridlines.", "Format major gridlines: light grey, 0.5 pt.", "Axis › Tick Marks: None."])
    tp(m, "3.7.6",
       [("Series in the zone chart", "=1", 1, "int", "One series: delete the legend.")],
       [("t", "A chart with one series: keep or delete the legend?", "delete", None, None, ("delete it", "remove"))],
       steps=["Select the legend › Delete.", "For 2-3 lines, label each line at its end instead (Data Label › Series Name).",
              "If a legend stays, put it at the top, close to the data."])
    tp(m, "3.7.7",
       [("Peak month value", "=MAX(tMonth[Sales])", f(ms.Sales.max()), N, "Label only the peak on a 12-point line."),
        ("Peak month", "=INDEX(tMonth[Month],MATCH(MAX(tMonth[Sales]),tMonth[Sales],0))", MON[int(ms.Sales.idxmax()) - 1], "gen", "")],
       [("t", "Which month was the lowest of FY2025?", MON[int(ms.Sales.idxmin()) - 1], "=INDEX(tMonth[Month],MATCH(MIN(tMonth[Sales]),tMonth[Sales],0))",
         "Lowest month")],
       steps=["Chart Elements › Data Labels for all points (few points).", "For one point: click it twice, then add a label.",
              "Number format of labels: ₦#,##0.0,,\"m\"."])
    tp(m, "3.7.8",
       [("Peak month (annotation)", "=INDEX(tMonth[Month],MATCH(MAX(tMonth[Sales]),tMonth[Sales],0))", MON[int(ms.Sales.idxmax()) - 1], "gen", ""),
        ("Q4 share of the year (subtitle fact)", "=SUM(INDEX(tMonth[Sales],10),INDEX(tMonth[Sales],11),INDEX(tMonth[Sales],12))/SUM(tMonth[Sales])",
         f(ms.Sales.loc[10:12].sum() / ms.Sales.sum()), "pct", "")],
       [("o", "Write a title, subtitle and annotation for the FY2025 monthly line.",
         "Title: Sales build through the year and peak in the fourth quarter. Subtitle: SuperMart Stores, FY2025 monthly sales, ₦m. "
         "Annotation on the peak month with its value.")],
       steps=["Title = the finding (a sentence).", "Subtitle = scope: company, period, unit.", "Annotation = a text box or label pointing at the feature."])
    tp(m, "3.7.9",
       [("Stores (bar chart)", "=ROWS(tDashStores)", 12, "int", "Comparison."),
        ("Months (line chart)", "=ROWS(tMonth)", 12, "int", "Change over time."),
        ("Students (scatter)", "=ROWS(tCV)", len(CV), "int", "Relationship.")],
       [("t", "'How long do patients stay?' needs which chart? (histogram / pie / line)", "histogram", None, None, ()),
        ("t", "'What share of customers is Gold?' needs which chart? (pie / scatter / line)", "pie", None, None, ("doughnut", "bar"))],
       steps=["Name the question type: compare, time, part of whole, relationship, distribution.", "Pick from the matching family.", "Prefer bars when unsure."])

    def ch_column(mm):
        mm.s.chart("column", [{"name": "Sales", "categories": T["z"].ref("Zone"), "values": T["z"].ref("Sales"), "labels": True,
                               "label_fmt": '₦#,##0.0,,"m"'}],
                   f"South-West led FY2025 sales", y_fmt='₦#,##0,,"m"', legend=False, rows=16)

    tp(m, "3.7.10",
       [("South-West sales", '=XLOOKUP("South-West",tZoneSum[Zone],tZoneSum[Sales])', f(sw), N, ""),
        ("South-East sales", '=XLOOKUP("South-East",tZoneSum[Zone],tZoneSum[Sales])', f(se), N, "")],
       [("n", "What is the South-South column's value?", '=XLOOKUP("South-South",tZoneSum[Zone],tZoneSum[Sales])', f(zs.Sales["South-South"]), N,
         "South-South sales")],
       steps=["Select tZoneSum[Zone] and tZoneSum[Sales] (hold Ctrl).", "Insert › Column › Clustered Column.", "Delete the legend; title states the finding.",
              "Value axis: Display units › Millions."], extra=ch_column)
    srt = DSS.sort_values("Sales", ascending=False)

    def ch_bar(mm):
        mm.s.chart("bar", [{"name": "Sales", "categories": ["Data_DashStores", 1, 1, 12, 1], "values": ["Data_DashStores", 1, 3, 12, 3]}],
                   "FY2025 sales by store", y_fmt='₦#,##0,,"m"', legend=False, rows=18,
                   extra=lambda ch: ch.set_y_axis({"reverse": True, "num_font": {"size": 8}}))

    tp(m, "3.7.11",
       [("Top store's sales", "=MAX(tDashStores[Sales])", f(srt.Sales.iloc[0]), N, ""),
        ("Top store", "=INDEX(tDashStores[Store_Name],MATCH(MAX(tDashStores[Sales]),tDashStores[Sales],0))", srt.Store_Name.iloc[0], "gen", "")],
       [("t", "Which store had the lowest FY2025 sales?", srt.Store_Name.iloc[-1],
         "=INDEX(tDashStores[Store_Name],MATCH(MIN(tDashStores[Sales]),tDashStores[Sales],0))", "Lowest store"),
        ("n", "What is the 3rd store's sales?", "=LARGE(tDashStores[Sales],3)", f(srt.Sales.iloc[2]), N, "3rd store")],
       steps=["Sort Data_DashStores by Sales, largest first (or sort the source).", "Insert › Bar › Clustered Bar.",
              "Vertical axis › Categories in reverse order, so the first row is on top."], extra=ch_bar)

    def ch_line(mm):
        mm.s.chart("line", [{"name": "Sales", "categories": T["m"].ref("Month"), "values": T["m"].ref("Sales")}],
                   "Sales build through FY2025", y_fmt='₦#,##0,,"m"', legend=False, rows=16)

    tp(m, "3.7.12",
       [("Lowest month", "=MIN(tMonth[Sales])", f(ms.Sales.min()), N, ""),
        ("Highest month", "=MAX(tMonth[Sales])", f(ms.Sales.max()), N, "")],
       [("n", "How much higher was the best month than the worst? (₦)", "=MAX(tMonth[Sales])-MIN(tMonth[Sales])", f(ms.Sales.max() - ms.Sales.min()), N,
         "Best minus worst month")],
       steps=["Select tMonth[Month] and tMonth[Sales].", "Insert › Line › Line.", "Label the peak; delete the legend."], extra=ch_line)

    def ch_area(mm):
        mm.s.chart("area", [{"name": "Sales to date", "categories": T["m"].ref("Month"), "values": T["m"].ref("YTD")}],
                   "FY2025 sales to date", y_fmt='₦#,##0,,"m"', legend=False, rows=16)

    tp(m, "3.7.13",
       [("Sales to date at June", "=XLOOKUP(6,tMonth[Month_No],tMonth[YTD])", f(ms.Sales.loc[:6].sum()), N, "The area grows to the year total."),
        ("Year total", "=SUM(tMonth[Sales])", TOT, N, "")],
       [("n", "What are sales to date at the end of September?", "=XLOOKUP(9,tMonth[Month_No],tMonth[YTD])", f(ms.Sales.loc[:9].sum()), N, "YTD September")],
       steps=["Select tMonth[Month] and tMonth[YTD].", "Insert › Area › Area.", "For parts over time (zones), use Stacked Area: read only the bottom layer exactly."],
       extra=ch_area)
    tier = C.membership_tier.value_counts()
    tiers = list(tier.index)

    def tier_table(mm):
        df = pd.DataFrame({"Tier": tiers})
        calc = {"Customers": ("COUNTIF(tCustomers[membership_tier],[@Tier])", "int", [int(tier[t]) for t in tiers]),
                "Share": ("[@Customers]/SUM([Customers])", "pct", [f(tier[t] / len(C)) for t in tiers])}
        mm.h("tTier: customers by membership tier")
        T["t"] = mm.s.table(df, "tTier", calc=calc, widths={"Tier": 12, "Customers": 11, "Share": 9})
        mm.s.r += 1
        mm.s.chart("pie", [{"name": "Customers", "categories": T["t"].ref("Tier"), "values": T["t"].ref("Customers"), "labels": True, "label_fmt": "#,##0"}],
                   "SuperMart Online customers by tier", size=(480, 280), rows=15)

    tp(m, "3.7.14",
       [("Silver share", '=XLOOKUP("Silver",tTier[Tier],tTier[Share])', f(tier["Silver"] / len(C)), "pct", ""),
        ("Gold share", '=XLOOKUP("Gold",tTier[Tier],tTier[Share])', f(tier["Gold"] / len(C)), "pct", "Nearly equal slices: a bar shows the order.")],
       [("n", "What share of customers is Platinum?", '=XLOOKUP("Platinum",tTier[Tier],tTier[Share])', f(tier["Platinum"] / len(C)), "pct", "Platinum share")],
       steps=["Select tTier[Tier] and tTier[Customers].", "Insert › Pie › Pie.", "Data labels: Category Name and Percentage; delete the legend."], extra=tier_table)
    pay = TX.groupby("payment_method").total_price.sum().sort_values(ascending=False)

    def pay_table(mm):
        df = pd.DataFrame({"Method": list(pay.index)})
        calc = {"Value": ("SUMIF(tTransactions[payment_method],[@Method],tTransactions[total_price])", N, [f(x) for x in pay]),
                "Share": ("[@Value]/SUM([Value])", "pct", [f(x / pay.sum()) for x in pay])}
        mm.h("tPay: order value by payment method (all lines)")
        T["p"] = mm.s.table(df, "tPay", calc=calc, widths={"Method": 16, "Value": 15, "Share": 9})
        mm.s.r += 1
        mm.s.chart("doughnut", [{"name": "Value", "categories": T["p"].ref("Method"), "values": T["p"].ref("Value")}],
                   "Order value by payment method", size=(480, 280), rows=15)

    tp(m, "3.7.15",
       [("Largest method's share", "=MAX(tPay[Share])", f(pay.iloc[0] / pay.sum()), "pct", ""),
        ("Methods", "=ROWS(tPay)", len(pay), "int", "")],
       [("n", "What share of value was paid by Bank Transfer?", '=XLOOKUP("Bank Transfer",tPay[Method],tPay[Share])', f(pay["Bank Transfer"] / pay.sum()), "pct",
         "Bank Transfer share")],
       steps=["Select tPay[Method] and tPay[Value].", "Insert › Pie › Doughnut.", "Put the total in the hole with a text box linked to a cell (=cell)."],
       extra=pay_table)
    r_att = CV.Final_Score.corr(CV.Lecture_Attendance_Pct)

    def ch_scatter(mm):
        mm.s.chart("scatter", [{"name": "Students", "categories": ["Data_CV", 1, 6, len(CV), 6], "values": ["Data_CV", 1, 8, len(CV), 8]}],
                   "Exam score against attendance (5,840 students)", x_title="Attendance %", y_title="Exam score", legend=False, rows=16,
                   extra=lambda ch: ch.set_x_axis({"name": "Attendance %", "min": 40, "max": 100}))

    tp(m, "3.7.16",
       [("Correlation: score vs attendance", "=CORREL(tCV[Final_Score],tCV[Lecture_Attendance_Pct])", f(r_att), "dec2", ""),
        ("Correlation: score vs sleep", "=CORREL(tCV[Final_Score],tCV[Sleep_Hrs_Night])", f(CV.Final_Score.corr(CV.Sleep_Hrs_Night)), "dec2", "About zero.")],
       [("n", "Correlation between exam score and hours studied per week? (2 decimals)", "=CORREL(tCV[Final_Score],tCV[Study_Hours_Wk])",
         f(CV.Final_Score.corr(CV.Study_Hours_Wk)), "dec2", "r score-hours")],
       steps=["Select Lecture_Attendance_Pct then Final_Score (x first).", "Insert › Scatter.", "Add a trendline: Chart Elements › Trendline."], extra=ch_scatter)
    dep = CV.groupby("Department").agg(H=("Study_Hours_Wk", "mean"), S=("Final_Score", "mean"), n=("Final_Score", "size"))
    deps = list(dep.index)

    def dept_table(mm):
        df = pd.DataFrame({"Department": deps})
        calc = {"Hours": ("AVERAGEIF(tCV[Department],[@Department],tCV[Study_Hours_Wk])", "dec1", [f(x) for x in dep.H]),
                "Score": ("AVERAGEIF(tCV[Department],[@Department],tCV[Final_Score])", "dec1", [f(x) for x in dep.S]),
                "Students": ("COUNTIF(tCV[Department],[@Department])", "int", [int(x) for x in dep.n])}
        mm.h("tDept: one row per department (bubble chart source)")
        T["d"] = mm.s.table(df, "tDept", calc=calc, widths={"Department": 24, "Hours": 8, "Score": 8, "Students": 9})
        mm.s.r += 1

    tp(m, "3.7.17",
       [("Departments (bubbles)", "=ROWS(tDept)", len(dep), "int", ""),
        ("Largest department (students)", "=MAX(tDept[Students])", int(dep.n.max()), "int", "Biggest bubble.")],
       [("t", "Which department has the highest average exam score?", dep.S.idxmax(), "=INDEX(tDept[Department],MATCH(MAX(tDept[Score]),tDept[Score],0))",
         "Top-scoring department")],
       steps=["Select Hours, Score and Students (x, y, size).", "Insert › Scatter › Bubble.", "Format › Scale bubble size; keep 'Area of bubble'."],
       extra=dept_table)

    def ch_radar(mm):
        ms_ = ["Sales", "Profit", "Lines", "Margin"]
        df = pd.DataFrame({"Zone": ZONES})
        calc = {k: (f"XLOOKUP([@Zone],tZoneSum[Zone],tZoneSum[{k}])/MAX(tZoneSum[{k}])", "dec2", [f(zs.loc[z, k] / zs[k].max()) for z in ZONES])
                for k in ms_}
        mm.h("tRadar: each measure as a share of the best zone (1.00 = best)")
        t = mm.s.table(df, "tRadar", calc=calc, widths={"Zone": 14, **{k: 10 for k in ms_}})
        T["r"] = t
        mm.s.r += 1
        c0, c1 = t.ci("Sales"), t.ci("Margin")
        mm.s.chart("radar", [{"name": z, "categories": [t.sheet, t.r0, c0, t.r0, c1], "values": [t.sheet, t.r0 + 1 + i, c0, t.r0 + 1 + i, c1]}
                             for i, z in enumerate(ZONES)], "FY2025 zone profile (1.00 = best zone)", rows=17)

    tp(m, "3.7.18",
       [("North-Central profit vs best zone", '=XLOOKUP("North-Central",tZoneSum[Zone],tZoneSum[Profit])/MAX(tZoneSum[Profit])', f(zs.Profit["North-Central"] / zs.Profit.max()),
         "dec2", "")],
       [("n", "South-East margin as a share of the best margin? (2 decimals)", '=XLOOKUP("South-East",tZoneSum[Zone],tZoneSum[Margin])/MAX(tZoneSum[Margin])',
         f(zs.Margin["South-East"] / zs.Margin.max()), "dec2", "SE margin vs best")],
       steps=["Put measures on the same scale first (share of the best).", "Select the table › Insert › Radar.", "Keep to 3-5 items and 4-6 spokes."],
       extra=ch_radar)

    def ch_combo(mm):
        mm.s.chart("column", [{"name": "Sales", "categories": T["z"].ref("Zone"), "values": T["z"].ref("Sales")}], "Sales and margin by zone, FY2025",
                   y_fmt='₦#,##0,,"m"', secondary={"type": "line", "name": "Margin", "categories": T["z"].ref("Zone"), "values": T["z"].ref("Margin"),
                                                   "y2_title": "Margin", "y2_fmt": "0%"}, rows=16)

    tp(m, "3.7.19",
       [("South-East margin", '=XLOOKUP("South-East",tZoneSum[Zone],tZoneSum[Margin])', f(zs.Margin["South-East"]), "pct", ""),
        ("North-Central margin", '=XLOOKUP("North-Central",tZoneSum[Zone],tZoneSum[Margin])', f(zs.Margin["North-Central"]), "pct", "Lower margin, higher sales.")],
       [("n", "What is the South-West margin?", '=XLOOKUP("South-West",tZoneSum[Zone],tZoneSum[Margin])', f(zs.Margin["South-West"]), "pct", "SW margin")],
       steps=["Select Zone, Sales and Margin.", "Insert › Combo › Clustered Column – Line on Secondary Axis.", "Title both value axes (₦m and %)."],
       extra=ch_combo)
    tp(m, "3.7.20",
       [("Start: 0", "=0", 0, N, "Bridge from zero to total profit, one step per zone."),
        ("+ South-West", '=XLOOKUP("South-West",tZoneSum[Zone],tZoneSum[Profit])', f(zs.Profit["South-West"]), N, ""),
        ("+ South-South", '=XLOOKUP("South-South",tZoneSum[Zone],tZoneSum[Profit])', f(zs.Profit["South-South"]), N, ""),
        ("+ North-Central", '=XLOOKUP("North-Central",tZoneSum[Zone],tZoneSum[Profit])', f(zs.Profit["North-Central"]), N, ""),
        ("+ South-East", '=XLOOKUP("South-East",tZoneSum[Zone],tZoneSum[Profit])', f(zs.Profit["South-East"]), N, ""),
        ("= FY2025 profit (set as total)", "=SUM(tZoneSum[Profit])", f(DS.Profit.sum()), N, "")],
       [("n", "What is the profit of all Furniture lines? (one step of a category bridge)", '=SUMIF(tDash[Category],"Furniture",tDash[Profit])',
         f(DS[DS.Category == "Furniture"].Profit.sum()), N, "Furniture profit")],
       steps=["Select the six labels and amounts above (copy them as values to a range first).", "Insert › Waterfall.",
              "Right-click the last bar › Set as total."])
    shipped, netl = int(TX.delivery_id.notna().sum()), len(NT)
    tp(m, "3.7.21",
       [("Orders placed", "=ROWS(tTransactions)", len(TX), "int", ""),
        ("Shipped (has a delivery)", "=COUNTA(tTransactions[delivery_id])", shipped, "int", ""),
        ("Delivered or completed", "=" + netn(t="tTransactions"), netl, "int", "")],
       [("n", "What % of placed orders end delivered or completed?", "=(" + netn(t="tTransactions") + ")/ROWS(tTransactions)", netl / len(TX), "pct",
         "Funnel conversion")],
       steps=["Put the three stage labels and counts in a range, largest first.", "Insert › Funnel.", "Check each stage is a subset of the one above."])
    sub = NT.groupby("subcategory").total_price.sum().sort_values(ascending=False)
    tp(m, "3.7.22",
       [(f"{sub.index[0]} net revenue", "=" + net(subcategory=sub.index[0]), f(sub.iloc[0]), N, "The largest rectangle."),
        (f"{sub.index[1]} net revenue", "=" + net(subcategory=sub.index[1]), f(sub.iloc[1]), N, "")],
       [("n", f"What is the {sub.index[2]} net revenue?", "=" + net(subcategory=sub.index[2]), f(sub.iloc[2]), N, f"{sub.index[2]} net revenue")],
       steps=["PivotTable: category and subcategory on rows, [Net Revenue] in values; copy as values.", "Select the range › Insert › Treemap.",
              "Label the big rectangles only."])
    tp(m, "3.7.23",
       [("Categories (inner ring)", "=COUNTA(UNIQUE(tLines[category]))", int(L.category.nunique()), "int", ""),
        ("Subcategories (outer ring)", "=COUNTA(UNIQUE(tLines[subcategory]))", int(L.subcategory.nunique()), "int", "")],
       [("n", "Electronics net revenue (its inner-ring segment)?", "=" + net(category="Electronics"), f(NT[NT.category == "Electronics"].total_price.sum()), N,
         "Electronics segment")],
       steps=["Same category + subcategory range as the treemap.", "Insert › Sunburst.", "Use a sorted bar when exact ranking matters."])
    edges = [0, 5, 10, 15, 20, 25, 30, 1000]
    lab = ["0–4", "5–9", "10–14", "15–19", "20–24", "25–29", "30+"]
    cnt = [int(((A.LOS_Days >= a) & (A.LOS_Days < b_)).sum()) for a, b_ in zip(edges[:-1], edges[1:])]

    def bins(mm):
        df = pd.DataFrame({"Days": lab, "From": edges[:-1], "Below": edges[1:]})
        mm.h("tBins: admissions 2023 by length of stay (bin width 5 days)")
        T["b"] = mm.s.table(df, "tBins", fmts={"From": "int", "Below": "int"},
                            calc={"Admissions": ('COUNTIFS(tAdm23[LOS_Days],">="&[@From],tAdm23[LOS_Days],"<"&[@Below])', "int", cnt)},
                            widths={"Days": 8, "From": 7, "Below": 7, "Admissions": 11})
        mm.s.r += 1
        mm.s.chart("column", [{"name": "Admissions", "categories": T["b"].ref("Days"), "values": T["b"].ref("Admissions"), "gap": 5}],
                   "Length of stay, Mercyland admissions 2023", x_title="Days", legend=False, rows=16)

    tp(m, "3.7.24",
       [("Admissions 2023", "=SUM(tBins[Admissions])", len(A), "int", "Every admission falls in one bin."),
        ("Most common bin's count", "=MAX(tBins[Admissions])", max(cnt), "int", "")],
       [("n", "How many 2023 admissions stayed 30 days or more?", '=XLOOKUP("30+",tBins[Days],tBins[Admissions])', cnt[-1], "int", "30+ days")],
       steps=["Select tAdm23[LOS_Days] › Insert › Histogram (Excel builds the bins).", "Format the horizontal axis › Bin width: 5.",
              "The column chart here is the same picture from tBins (gap width 5%)."], extra=bins)
    q = A.LOS_Days.quantile([0.25, 0.5, 0.75])
    tp(m, "3.7.25",
       [("Q1 (all 2023)", "=QUARTILE.INC(tAdm23[LOS_Days],1)", f(q[0.25]), "dec1", ""),
        ("Median", "=MEDIAN(tAdm23[LOS_Days])", f(q[0.5]), "dec1", ""),
        ("Q3", "=QUARTILE.INC(tAdm23[LOS_Days],3)", f(q[0.75]), "dec1", "")],
       [("n", "What is the interquartile range (Q3 − Q1)?", "=QUARTILE.INC(tAdm23[LOS_Days],3)-QUARTILE.INC(tAdm23[LOS_Days],1)",
         f(q[0.75] - q[0.25]), "dec1", "IQR"),
        ("n", "Average length of stay in Cardiology, 2023 (1 decimal)?", '=AVERAGEIF(tAdm23[Department],"Cardiology",tAdm23[LOS_Days])',
         f(A[A.Department == "Cardiology"].LOS_Days.mean()), "dec1", "Cardiology ALOS")],
       steps=["Select Department and LOS_Days from tAdm23.", "Insert › Box and Whisker.", "Format series: show mean markers; compare the boxes."])
    stores = list(DSS.Store_Name)
    sm = DS.assign(M=DS.Order_Date.dt.month).pivot_table(index="Store_Name", columns="M", values="Sales", aggfunc="sum", fill_value=0).reindex(stores)

    def spark(mm):
        df = pd.DataFrame({"Store_Name": stores})
        calc = {MON[k]: (f"SUMIFS(tDash[Sales],tDash[Store_Name],[@Store_Name],{mdates(k + 1)})", N, [f(x) for x in sm[k + 1]]) for k in range(12)}
        mm.h("tStoreMonth: FY2025 sales per store and month, with a sparkline per row")
        t = mm.s.table(df, "tStoreMonth", calc=calc, widths={"Store_Name": 26, **{x: 11 for x in MON}})
        T["s"] = t
        mm.ws.set_column(t.c0 + 13, t.c0 + 13, 16)
        mm.ws.write(t.r0, t.c0 + 13, "Trend", mm.F["hdr"])
        for i in range(len(stores)):
            r = t.r0 + 1 + i
            mm.ws.add_sparkline(r, t.c0 + 13, {"range": f"'{mm.s.name}'!{t.cell(i, 'Jan')}:{t.cell(i, 'Dec')}", "markers": True, "high_point": True,
                                               "series_color": "#0B2545"})
        mm.s.r += 1

    top = DSS.sort_values("Sales", ascending=False).Store_Name.iloc[0]
    tp(m, "3.7.26",
       [(f"{top}: January", f'=XLOOKUP("{top}",tStoreMonth[Store_Name],tStoreMonth[Jan])', f(sm.loc[top, 1]), N, ""),
        (f"{top}: December", f'=XLOOKUP("{top}",tStoreMonth[Store_Name],tStoreMonth[Dec])', f(sm.loc[top, 12]), N, "")],
       [("n", "What were Jos - Rayfield's December sales?", '=XLOOKUP("Jos - Rayfield",tStoreMonth[Store_Name],tStoreMonth[Dec])', f(sm.loc["Jos - Rayfield", 12]),
         N, "Jos - Rayfield December"),
        ("t", "Sparklines › Axis: which option makes every row use the same scale? ('Same for All …')", "Same for All Sparklines", None, None,
         ("same for all", "same for all sparklines"))],
       steps=["Select the 12 monthly cells of the first store.", "Insert › Sparklines › Line; Location: the Trend cell › OK; fill down.",
              "Sparkline tab › tick High Point; Axis › Same for All Sparklines."], extra=spark)
    m.finish()
    return T


# =====================================================================================================
def m38(b, D):
    DS, DSS = D.DS, D.DSS
    m = ModuleSheet(b, "3.8 Dashboard", "3.8  Dashboard Development",
                    "Tutor walk-through after Module 3.8: plan, lay out and build a one-screen dashboard, with a zone selector that drives "
                    "the KPI cards", "3.8.1 - 3.8.19")
    m.dataset([DASH_DS, STORES_DS, ("ZoneSel", "On this sheet (3.8.7)", "A yellow cell with a drop-down: All or one zone. The KPI cards follow it.")])
    S, P, TG = f(DS.Sales.sum()), f(DS.Profit.sum()), f(DSS.Target.sum())
    zs = DS.groupby("Zone").agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"))
    tz = DSS.groupby("Zone").Target.sum()
    on = int((DSS.Attainment >= 1).sum())
    tp(m, "3.8.1",
       [("Net sales", "=SUM(tDash[Sales])", S, N, ""), ("Profit", "=SUM(tDash[Profit])", P, N, ""),
        ("Margin", "=SUM(tDash[Profit])/SUM(tDash[Sales])", P / S, "pct", ""), ("Target attainment", "=SUM(tDash[Sales])/SUM(tDashStores[Target])", S / TG, "pct", "")],
       [("n", "What is the FY2025 target for all stores?", "=SUM(tDashStores[Target])", TG, N, "Total target")],
       steps=["Agree the four KPIs with the audience.", "Work each out once, in a formula or measure.", "Show them as cards across the top."])
    tp(m, "3.8.2",
       [("Line total − store total (must be 0)", "=SUM(tDash[Sales])-SUM(tDashStores[Sales])", 0, N, "Reconciles."),
        ("Zones add up to the total (difference)", "=SUMIF(tDash[Zone],\"South-West\",tDash[Sales])+SUMIF(tDash[Zone],\"South-South\",tDash[Sales])+"
                                                   "SUMIF(tDash[Zone],\"North-Central\",tDash[Sales])+SUMIF(tDash[Zone],\"South-East\",tDash[Sales])-SUM(tDash[Sales])", 0, N, "")],
       [("t", "Where on the screen does the most important KPI go? (top-left / bottom-right)", "top-left", None, None, ("top left",))],
       steps=["One screen, one audience, one purpose.", "Most important top-left; colour only for meaning.", "Every number reconciles to a control total."])
    tp(m, "3.8.3",
       [("Stores below target (to act on)", '=COUNTIF(tDashStores[Attainment],"<1")', len(DSS) - on, "int", "")],
       [("n", "How many stores are at or above target?", '=COUNTIF(tDashStores[Attainment],">=1")', on, "int", "Stores on target"),
        ("o", "Write one question, its decision and its visual for a regional manager.",
         "Question: which of my stores are behind target? Decision: which store to visit or support this month. Visual: sorted bar of attainment.")],
       steps=["Write the audience, their questions and the decision each drives.", "List the measures each question needs.", "Build nothing that serves no decision."])
    tp(m, "3.8.4",
       [("Stores in the bar chart", "=ROWS(tDashStores)", len(DSS), "int", "Middle left."),
        ("Months in the line chart", "=12", 12, "int", "Middle right.")],
       [("t", "Readers scan a dashboard in which letter shape?", "Z", None, None, ("z shape", "z-pattern"))],
       steps=["Set equal column widths for the grid.", "KPIs on top, charts below, notes at the bottom.", "Format › Align › Align Left / Top and Distribute."])
    best = DSS.sort_values("Attainment", ascending=False).iloc[0]
    tp(m, "3.8.5",
       [("Stores in navy (on target)", '=COUNTIF(tDashStores[Attainment],">=1")', on, "int", ""),
        ("Stores in red (below)", '=COUNTIF(tDashStores[Attainment],"<1")', len(DSS) - on, "int", ""),
        ("Best attainment", "=MAX(tDashStores[Attainment])", f(best.Attainment), "pct", "")],
       [("t", "Which store has the best attainment?", best.Store_Name, "=INDEX(tDashStores[Store_Name],MATCH(MAX(tDashStores[Attainment]),tDashStores[Attainment],0))",
         "Best store")],
       steps=["One main colour for data, grey for context.", "Red/green only for status against a target.", "Add a label or symbol too (3.8.17)."])
    tp(m, "3.8.6",
       [("KPI value in ₦m (large type)", "=ROUND(SUM(tDash[Sales])/1000000,1)", round(S / 1e6, 1), "dec1", "")],
       [("n", "Profit in ₦m to 1 decimal?", "=ROUND(SUM(tDash[Profit])/1000000,1)", round(P / 1e6, 1), "dec1", "Profit ₦m")],
       steps=["Three sizes: large values, medium labels and titles, small notes.", "One font family.", "Bold for values only."])
    sel = "All"

    def selector(mm):
        mm.h("Zone selector (a stand-in for a slicer) and KPI cards")
        out = mm.s.inputs([("ZoneSel: All or a zone", sel, "gen", "Pick from the drop-down; the cards below follow it.")])
        a = list(out.values())[0]
        r, c = int("".join(ch for ch in a if ch.isdigit())) - 1, 4
        mm.ws.data_validation(r, c, r, c, {"validate": "list", "source": ["All"] + list(zs.index)})
        mm.__dict__["sel"] = a

    def cards(mm):
        z = mm.sel
        return [("Net sales", f'=IF({z}="All",SUM(tDash[Sales]),SUMIF(tDash[Zone],{z},tDash[Sales]))', S, N, "Follows ZoneSel."),
                ("Profit", f'=IF({z}="All",SUM(tDash[Profit]),SUMIF(tDash[Zone],{z},tDash[Profit]))', P, N, ""),
                ("Margin", f'=IF({z}="All",SUM(tDash[Profit])/SUM(tDash[Sales]),SUMIF(tDash[Zone],{z},tDash[Profit])/SUMIF(tDash[Zone],{z},tDash[Sales]))', P / S, "pct", ""),
                ("Target attainment", f'=IF({z}="All",SUM(tDash[Sales])/SUM(tDashStores[Target]),SUMIF(tDash[Zone],{z},tDash[Sales])/SUMIF(tDashStores[Zone],{z},tDashStores[Target]))',
                 S / TG, "pct", "")]

    tp(m, "3.8.7", cards,
       [("n", "Set ZoneSel to South-East. What does the Net sales card show?", '=SUMIF(tDash[Zone],"South-East",tDash[Sales])', f(zs.Sales["South-East"]), N,
         "SE net sales card"),
        ("n", "And the Target attainment card for South-East?", '=SUMIF(tDash[Zone],"South-East",tDash[Sales])/SUMIF(tDashStores[Zone],"South-East",tDashStores[Target])',
         f(zs.Sales["South-East"] / tz["South-East"]), "pct", "SE attainment card")],
       steps=["Each card: label, value, comparison, status.", "Feed it from a formula that reads the selection (or GETPIVOTDATA).",
              "Set ZoneSel back to All when done."], extra=selector)
    low = DSS.sort_values("Attainment").iloc[0]
    tp(m, "3.8.8",
       [("Lowest attainment", "=MIN(tDashStores[Attainment])", f(low.Attainment), "pct", "Sorted bar: the store behind."),
        ("Store", "=INDEX(tDashStores[Store_Name],MATCH(MIN(tDashStores[Attainment]),tDashStores[Attainment],0))", low.Store_Name, "gen", "")],
       [("t", "'How did sales move month by month?' Which visual? (line / pie / card)", "line", None, None, ("line chart",))],
       steps=["Total = card; trend = line; ranking = sorted bar.", "Share = stacked bar; actual vs target = bar with a target marker.", "One visual per question."])
    tp(m, "3.8.9",
       [(f"{z} net sales", f'=SUMIF(tDash[Zone],"{z}",tDash[Sales])', f(zs.Sales[z]), N, "") for z in ZONES],
       [("n", "Slicer button North-Central: what net sales show?", '=SUMIF(tDash[Zone],"North-Central",tDash[Sales])', f(zs.Sales["North-Central"]), N,
         "North-Central net sales")],
       steps=["Build a PivotTable from tDash › PivotTable Analyze › Insert Slicer › Zone.", "Slicer tab › Columns: 4; size it to the grid row.",
              "Keep the Clear Filter button visible."],
       do=[("Insert a Zone slicer on a PivotTable of tDash and click South-West.", f"The PivotTable shows ₦{zs.Sales['South-West']:,.0f}.")])
    tp(m, "3.8.10",
       [("South-East stores (bars after connecting)", '=COUNTIF(tDashStores[Zone],"South-East")', int((DSS.Zone == "South-East").sum()), "int", "")],
       [("n", "How many South-West stores would the connected store chart show?", '=COUNTIF(tDashStores[Zone],"South-West")', int((DSS.Zone == "South-West").sum()),
         "int", "SW stores")],
       steps=["Right-click the slicer › Report Connections.", "Tick every PivotTable on the dashboard.", "Click a zone: every visual must change."])
    qtr = DS.groupby(DS.Order_Date.dt.quarter).Sales.sum()
    tp(m, "3.8.11",
       [(f"Q{k} net sales", f'=SUMIFS(tDash[Sales],tDash[Order_Date],">="&DATE(2025,{3 * k - 2},1),tDash[Order_Date],"<="&EOMONTH(DATE(2025,{3 * k},1),0))',
         f(qtr[k]), N, "") for k in (1, 4)],
       [("n", "What are Q2 sales?", '=SUMIFS(tDash[Sales],tDash[Order_Date],">="&DATE(2025,4,1),tDash[Order_Date],"<="&DATE(2025,6,30))', f(qtr[2]), N, "Q2 sales"),
        ("n", "Q4 share of the year?", '=SUMIFS(tDash[Sales],tDash[Order_Date],">="&DATE(2025,10,1))/SUM(tDash[Sales])', f(qtr[4] / S), "pct", "Q4 share")],
       steps=["PivotTable Analyze › Insert Timeline › Order_Date.", "Switch the level to QUARTERS.", "Drag across Q2–Q3 to select a range."])
    tp(m, "3.8.12",
       [("All zones", "=SUM(tDash[Sales])", S, N, "What everyone assumes the pivot shows."),
        ("South-South only (a forgotten report filter)", '=SUMIF(tDash[Zone],"South-South",tDash[Sales])', f(zs.Sales["South-South"]), N, "")],
       [("n", "What share of the year would the forgotten South-South filter show?", '=SUMIF(tDash[Zone],"South-South",tDash[Sales])/SUM(tDash[Sales])',
         f(zs.Sales["South-South"] / S), "pct", "Share shown")],
       steps=["Report filter: drop-down above one PivotTable.", "Slicer: visible buttons, can drive many PivotTables.", "Dashboards: slicers only."])
    lowid = low.Store_ID
    lc = DS[DS.Store_ID == lowid].groupby("Category").Sales.sum()
    tp(m, "3.8.13",
       [(f"{low.Store_Name}: lines behind the cell", f'=COUNTIF(tDash[Store_ID],"{lowid}")', int((DS.Store_ID == lowid).sum()), "int", "Double-click drills to these."),
        (f"{low.Store_Name}: Technology sales", f'=SUMIFS(tDash[Sales],tDash[Store_ID],"{lowid}",tDash[Category],"Technology")', f(lc.get("Technology", 0)), N, "")],
       [("n", f"{low.Store_Name}'s Furniture sales?", f'=SUMIFS(tDash[Sales],tDash[Store_ID],"{lowid}",tDash[Category],"Furniture")', f(lc.get("Furniture", 0)), N,
         "Furniture sales, lowest store")],
       steps=["Double-click a value in a PivotTable: a new sheet lists its rows.", "Delete drill sheets before sharing.", "Slicers filter; double-click drills; Clear resets."])
    tp(m, "3.8.14",
       [("Source rows behind the dashboard", "=ROWS(tDash)", len(DS), "int", "")],
       [("t", "When protecting the dashboard sheet, which permission keeps the slicers working? ('Use PivotTable & …')", "PivotChart", None, None,
         ("use pivottable & pivotchart", "pivottable & pivotchart"))],
       steps=["View › untick Gridlines and Headings.", "Align charts (Format › Align).", "Review › Protect Sheet › allow 'Use PivotTable & PivotChart'."])
    tp(m, "3.8.15",
       [("KPI cards to fit on one page", "=4", 4, "int", "")],
       [("t", "Orientation for a dashboard print-out? (portrait / landscape)", "landscape", None, None, ())],
       steps=["Select the dashboard range › Page Layout › Print Area › Set.", "Orientation: Landscape; Scale to Fit: 1 page wide, 1 tall.",
              "File › Export › PDF; open the PDF and check it is readable."])
    tp(m, "3.8.16",
       [("Compact KPI text for a phone", '=TEXT(SUM(tDash[Sales])/1000000,"0.0")&"m"', f"{S / 1e6:.1f}m", "gen", "")],
       [("t", "Hover-only details fail on a phone. What do you show instead?", "data labels", None, None, ("labels", "values as data labels"))],
       steps=["Open the file in Excel for the web and on a phone.", "Fewer visuals, larger values, cards first.", "Test slicers there."])
    stat = DSS.assign(S=lambda d: d.Attainment.ge(1).map({True: "On target", False: "Below"}))
    tp(m, "3.8.17",
       [("ST01 status as text", '=IF(XLOOKUP("ST01",tDashStores[Store_ID],tDashStores[Attainment])>=1,"On target","Below")',
         stat.set_index("Store_ID").S["ST01"], "gen", "Colour plus words.")],
       [("t", f"Status text for {lowid}?", stat.set_index("Store_ID").S[lowid],
         f'=IF(XLOOKUP("{lowid}",tDashStores[Store_ID],tDashStores[Attainment])>=1,"On target","Below")', "Status lowest store")],
       steps=["Contrast 4.5:1 for text.", "Never colour alone: add text or a symbol.", "Right-click each chart › Edit Alt Text."])
    tp(m, "3.8.18",
       [("Columns kept in the source", "=COLUMNS(tDash)", DS.shape[1], "int", "Only what the dashboard needs.")],
       [("t", "Which is volatile and recalculates on every change: SUMIFS or OFFSET?", "OFFSET", None, None, ())],
       steps=["One source table for all PivotTables.", "SUMIFS / INDEX, not OFFSET / INDIRECT.", "Remove unused columns in Power Query."])
    tp(m, "3.8.19",
       [("Net sales = control total (difference)", "=SUM(tDash[Sales])-SUM(tDashStores[Sales])", 0, N, "Pass."),
        ("Negative attainment values (expect 0)", '=COUNTIF(tDashStores[Attainment],"<0")', 0, "int", "Pass.")],
       [("n", "QA: does profit reconcile? Type SUM(tDash[Profit]) − SUM(tDashStores[Profit]).", "=SUM(tDash[Profit])-SUM(tDashStores[Profit])",
         f(DS.Profit.sum() - DSS.Profit.sum()), N, "Profit difference")],
       steps=["List each check with its expected value.", "Record the result and pass/fail.", "Default view: ZoneSel = All."])
    m.finish()


# =====================================================================================================
def m39(b, D):
    DS, DSS, CV, A, CEN = D.DS, D.DSS, D.CV, D.A, D.CEN
    m = ModuleSheet(b, "3.9 Blueprints", "3.9  Three Dashboard Blueprints",
                    "Tutor walk-through after Module 3.9: a KPI specification with live current values for each of the three businesses",
                    "3.9.1 - 3.9.3")
    m.dataset([DASH_DS, ("tCV", "Sheet Data_CV (A1)", "Kingsway students 2024/25."), ("tAdm23 / tCensus", "Data_Adm23, Data_Census (A1)", "Mercyland 2023.")])
    S, P = f(DS.Sales.sum()), f(DS.Profit.sum())
    tp(m, "3.9.1",
       [("Net sales", "=SUM(tDash[Sales])", S, N, "KPI card"),
        ("Target attainment", "=SUM(tDash[Sales])/SUM(tDashStores[Target])", S / f(DSS.Target.sum()), "pct", "Card + bar by store"),
        ("Margin %", "=SUM(tDash[Profit])/SUM(tDash[Sales])", P / S, "pct", "Card + combo by zone"),
        ("Loss-making lines", '=COUNTIF(tDash[Profit],"<0")', int((DS.Profit < 0).sum()), "int", "Drill-down table")],
       [("n", "What value do the loss-making lines lose in total?", '=SUMIF(tDash[Profit],"<0")', f(DS[DS.Profit < 0].Profit.sum()), N, "Total loss"),
        ("o", "Write the definition and source for 'Target attainment'.", "Sales ÷ target for the same stores and period; sales from tDash, targets from tDashStores.")],
       steps=["Audience: regional managers. Slicers: Zone, Month.", "List KPIs with definition, source and visual.", "Fill in current values to check the definitions work."])
    fa = CV.groupby("Faculty").Final_Score.mean()
    tp(m, "3.9.2",
       [("Students enrolled", "=ROWS(tCV)", len(CV), "int", ""), ("Average exam score", "=AVERAGE(tCV[Final_Score])", f(CV.Final_Score.mean()), "dec1", ""),
        ("Grade A share (70+)", '=COUNTIF(tCV[Final_Score],">=70")/ROWS(tCV)', f((CV.Final_Score >= 70).mean()), "pct", ""),
        ("At-risk students", '=COUNTIF(tCV[At_Risk],"Yes")', int((CV.At_Risk == "Yes").sum()), "int", "Restricted list.")],
       [("n", "Average exam score in Engineering?", '=AVERAGEIF(tCV[Faculty],"Engineering",tCV[Final_Score])', f(fa["Engineering"]), "dec1", "Engineering average"),
        ("n", "At-risk students in Science?", '=COUNTIFS(tCV[Faculty],"Science",tCV[At_Risk],"Yes")', int(((CV.Faculty == "Science") & (CV.At_Risk == "Yes")).sum()),
         "int", "At risk, Science")],
       steps=["Audience: senate and deans. Slicers: Faculty, Department, Level.", "Named lists only for advisers (personal data).", "Bar by faculty for scores."])
    c23 = CEN[CEN.Census_Date.dt.year == 2023]
    tp(m, "3.9.3",
       [("Admissions 2023", "=ROWS(tAdm23)", len(A), "int", ""),
        ("Average occupancy 2023", '=AVERAGEIFS(tCensus[Occupancy_Pct],tCensus[Census_Date],">="&DATE(2023,1,1),tCensus[Census_Date],"<="&DATE(2023,12,31))',
         f(c23.Occupancy_Pct.mean()), "pct", "Line with 85% / 95% bands."),
        ("ALOS", "=AVERAGE(tAdm23[LOS_Days])", f(A.LOS_Days.mean()), "dec1", ""),
        ("30-day readmission rate", '=COUNTIF(tAdm23[Readmitted_30d],"Yes")/ROWS(tAdm23)', f((A.Readmitted_30d == "Yes").mean()), "pct", ""),
        ("Average bill", "=AVERAGE(tAdm23[Bill_Amount_NGN])", f(A.Bill_Amount_NGN.mean()), N, "")],
       [("n", "How many 2023 census days were above 95% full?", '=COUNTIFS(tCensus[Occupancy_Pct],">0.95",tCensus[Census_Date],">="&DATE(2023,1,1),tCensus[Census_Date],"<="&DATE(2023,12,31))',
         int((c23.Occupancy_Pct > 0.95).sum()), "int", "Days above 95% in 2023")],
       steps=["Audience: operations team. Slicers: Department, Month, HMO.", "Occupancy from the daily census.", "Bands at 85% and 95%."])
    m.finish()


# =====================================================================================================
def m310(b, D):
    R, A = D.LD, D.A
    m = ModuleSheet(b, "3.10 File Extracts", "3.10  File Extracts into SQL, DAX and Dashboards",
                    "Tutor walk-through after Module 3.10: profile the raw 2023 Mercyland extract, stage it, report exceptions, measure and reconcile",
                    "3.10.1 - 3.10.5")
    m.dataset([("tMsDirty", "Sheet Data_LcDirty (A1)", "Mercyland_Episodes_2023_DIRTY.csv as received: 12,772 rows, every column text."),
               ("tAdm23", "Sheet Data_Adm23 (A1)", "Production: the 2023 admissions in mercyland_hospital.db.")])
    wl = R.Ward[R.Ward != ""].str.lower().nunique()
    tp(m, "3.10.1",
       [("Rows", "=ROWS(tMsDirty)", len(R), "int", ""),
        ("Blank Discharge_Date", '=COUNTIF(tMsDirty[Discharge_Date],"")', int((R.Discharge_Date == "").sum()), "int", "Flag, don't guess."),
        ("Blank HMO", '=COUNTIF(tMsDirty[HMO],"")', int((R.HMO == "").sum()), "int", "Label 'Unknown'."),
        ("Distinct Ward values (Excel ignores case)", "=COUNTA(UNIQUE(tMsDirty[Ward]))", int(wl), "int", "Python/SQL count 16: Excel's UNIQUE treats 'Baobab ward' = 'Baobab Ward'.")],
       [("n", "How many distinct Gender values does Excel see?", "=COUNTA(UNIQUE(tMsDirty[Gender]))", int(R.Gender.str.lower().nunique()), "int", "Distinct Gender"),
        ("n", "How many Admission_Route cells start with 'emerg' (any spelling)?", '=COUNTIF(tMsDirty[Admission_Route],"emerg*")',
         int(R.Admission_Route.str.lower().str.startswith("emerg").sum()), "int", "Emergency spellings")],
       steps=["Count rows and blanks per column.", "List distinct values of every text column.", "Record each problem and the action, before loading anything."])
    com = int(R.Bill_Amount_NGN.str.contains(",").sum())
    tp(m, "3.10.2",
       [("CSV data rows = staging rows", "=ROWS(tMsDirty)", len(R), "int", ""),
        ("Bills written with commas", '=COUNTIF(tMsDirty[Bill_Amount_NGN],"*,*")', com, "int", "CAST would read '2,507,101.35' as 2."),
        ("Bills that are plain numbers once commas go", '=SUMPRODUCT(--ISNUMBER(--SUBSTITUTE(tMsDirty[Bill_Amount_NGN],",","")))', len(R), "int", "")],
       [("n", "How many bills are written with a minus sign?", '=COUNTIF(tMsDirty[Bill_Amount_NGN],"-*")', int(R.Bill_Amount_NGN.str.startswith("-").sum()), "int",
         "Negative bills")],
       steps=["sqlite3: .mode csv, then .import --skip 1 the file into stg_admissions_2023.", "SELECT COUNT(*): must equal the file's data rows.",
              "Convert types in the next step, with REPLACE for the commas."], code=True)
    disc = int(((R.Discharge_Date != "") & (R.Discharge_Date < R.Admit_Date)).sum())
    tp(m, "3.10.3",
       [("Doctor not in production", "=SUMPRODUCT(--(COUNTIF(tAdm23[Consultant_Name],tMsDirty[Consultant_Name])=0))", int((~R.Consultant_Name.isin(A.Consultant_Name)).sum()),
         "int", "LEFT JOIN … IS NULL."),
        ("Discharge before admission", '=SUMPRODUCT((tMsDirty[Discharge_Date]<>"")*(tMsDirty[Discharge_Date]<tMsDirty[Admit_Date]))', disc, "int",
         "ISO text dates compare correctly."),
        ("Bill below zero", '=COUNTIF(tMsDirty[Bill_Amount_NGN],"-*")', int(R.Bill_Amount_NGN.str.startswith("-").sum()), "int", "")],
       [("n", "Production check: how many 2023 admissions in tAdm23 are discharged before admission?", "=SUMPRODUCT(--(tAdm23[Discharge_Date]<tAdm23[Admit_Date]))",
         int((A.Discharge_Date < A.Admit_Date).sum()), "int", "Production rule breaks"),
        ("n", "How many rows have no discharge date (an exception to report)?", '=COUNTIF(tMsDirty[Discharge_Date],"")', int((R.Discharge_Date == "").sum()), "int",
         "No discharge date")],
       steps=["One SELECT per rule, each returning a count.", "Stack them with UNION ALL into one report.", "Run the same report on production."], code=True)
    tp(m, "3.10.4",
       [("[Admissions] 2023", "=ROWS(tAdm23)", len(A), "int", ""),
        ("[ALOS] 2023", "=AVERAGE(tAdm23[LOS_Days])", f(A.LOS_Days.mean()), "dec1", ""),
        ("[Readmission %] 2023", '=COUNTIF(tAdm23[Readmitted_30d],"Yes")/ROWS(tAdm23)', f((A.Readmitted_30d == "Yes").mean()), "pct", ""),
        ("… Admission type = Emergency", '=COUNTIFS(tAdm23[Admission_Route],"Emergency",tAdm23[Readmitted_30d],"Yes")/COUNTIF(tAdm23[Admission_Route],"Emergency")',
         f((A[A.Admission_Route == "Emergency"].Readmitted_30d == "Yes").mean()), "pct", "")],
       [("n", "[Readmission %] for Surgery in 2023?", '=COUNTIFS(tAdm23[Department],"Surgery",tAdm23[Readmitted_30d],"Yes")/COUNTIF(tAdm23[Department],"Surgery")',
         f((A[A.Department == "Surgery"].Readmitted_30d == "Yes").mean()), "pct", "Readmission Surgery"),
        ("n", "[ALOS] for Elective admissions in 2023?", '=AVERAGEIF(tAdm23[Admission_Route],"Elective",tAdm23[LOS_Days])',
         f(A[A.Admission_Route == "Elective"].LOS_Days.mean()), "dec1", "ALOS Elective")],
       steps=["Load the cleaned table to the Data Model.", "Write the three measures (code above).", "Test each under a filter against the twin."], code=True)
    tp(m, "3.10.5",
       [("Admissions: production (SQL)", "=ROWS(tAdm23)", len(A), "int", "Dashboard card must show the same."),
        ("Billing total: production", "=SUM(tAdm23[Bill_Amount_NGN])", f(A.Bill_Amount_NGN.sum()), "naira2", ""),
        ("Raw extract rows − production rows", "=ROWS(tMsDirty)-ROWS(tAdm23)", len(R) - len(A), "int", "Explain every row before publishing.")],
       [("n", "How many raw rows are exact duplicates? (Data › Remove Duplicates on a copy of Data_LcDirty reports it.)", f"={int(R.duplicated().sum())}",
         int(R.duplicated().sum()), "int", "Duplicates in the extract"),
        ("o", "The extract has more rows than production. List what you would check to explain the difference.",
         "Exact duplicates, rows failing the exception rules (negative bills, discharge before admission, no discharge date), and any "
         "admissions outside 2023; then recount until extract minus removed rows equals production.")],
       steps=["Pick the controls: row count and billing total.", "Work each out from the file, the SQL and the dashboard.", "Publish only when all three agree; record the check."])
    m.finish()
