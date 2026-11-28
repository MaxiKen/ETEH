"""Week 3 edition-3 slide copy: Module 3.7 (chart anatomy and the 17 Excel chart types).
Chart data: SuperMart Stores POS lines 2022-2025 (tDash for FY2025), SuperMart Online, Kingsway and Mercyland (w3_fx)."""
import numpy as np
import pandas as pd

from v3kit import T
from slidekit import ORANGE
from w3_fx import CU, CV, DS, I, LC, LN, N, N_LINES, NT, Nm, P, S25, SM, SMONTH, TX, ZS, F1

Z = ZS.index.tolist()                     # zones by FY2025 sales, largest first
SW, SE_ = ZS.loc["South-West"], ZS.loc["South-East"]
RATIO0 = SW.Sales / SE_.Sales
RATIO100 = (SW.Sales - 100e6) / (SE_.Sales - 100e6)
PEAK = SMONTH.idxmax()
DSS = DS.sort_values("Sales", ascending=False)
YEARS = SM.groupby("Year").Sales.sum()
TIER = CU.membership_tier.value_counts()
PAY = TX.groupby("payment_method").total_price.sum().sort_values(ascending=False)
R = {c: CV.Final_Score.corr(CV[c]) for c in ("Lecture_Attendance_Pct", "Study_Hours_Wk", "Sleep_Hrs_Night")}
ZZ = S25.groupby("Zone").agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"), Orders=("Order_ID", "nunique"))
ZZ["Margin"] = ZZ.Profit / ZZ.Sales
RAD = ZZ / ZZ.max()
P24 = SM[SM.Year == 2024].groupby("Category").Profit.sum()
P25 = S25.groupby("Category").Profit.sum()
SUB = NT.groupby(["category", "subcategory"]).total_price.sum().sort_values(ascending=False)
NCAT, NSUB = NT.category.nunique(), NT.subcategory.nunique()
CATN = NT.groupby("category").total_price.sum().sort_values(ascending=False)
BINS = pd.cut(LC.LOS_Days, [0, 5, 10, 15, 20, 25, 30, 1000], right=False,
              labels=["0–4", "5–9", "10–14", "15–19", "20–24", "25–29", "30+"]).value_counts().sort_index()
BOX = LC.groupby("Department").LOS_Days.quantile([0.25, 0.5, 0.75]).unstack()
SP = S25.assign(M=S25.Order_Date.dt.month).pivot_table(index="Store_Name", columns="M", values="Sales", aggfunc="sum", fill_value=0)
NOT_CANC = int((TX.order_status != "Cancelled").sum())
SHIPPED = int(TX.delivery_id.notna().sum())
NETL = int(TX.order_status.isin(["Delivered", "Completed"]).sum())


def ym(p):
    return pd.Period(p).strftime("%b %Y")


def yr_lohi(y):
    s = SMONTH[[p for p in SMONTH.index if p.year == y]]
    return f"{s.idxmin().strftime('%b')} {Nm(s.min())}", f"{s.idxmax().strftime('%b')} {Nm(s.max())}"


def zone_rows(cols=("Sales", "Margin")):
    head = ["Zone", "FY2025 sales", "Margin"] if cols == ("Sales", "Margin") else ["Zone"] + list(cols)
    return [head] + [[z, N(ZS.loc[z, "Sales"]), P(ZS.loc[z, "Margin"])] for z in Z]


M37 = dict(
    id="3.7", title="Chart Anatomy and the 17 Excel Chart Types",
    subtitle="What makes a chart effective, the parts of every chart, and which of the 17 types answers which question.",
    objectives=["State what makes a chart effective in five seconds.",
                "Name and format every part of an Excel chart.",
                "Write titles that state the finding, with scope and annotations.",
                "Choose the chart type from the question being asked.",
                "Build the 17 chart types from column to sparkline, and know their traps."],
    topics=[
        T("3.7.1", "What Makes a Chart Effective",
          defn="An effective chart answers one question accurately, quickly and honestly, with nothing that distracts from the message.",
          simple="One clear message, a title that states the finding, and no 3D, heavy borders or extra colours.",
          grid=dict(label="Zone sales chart, before and after", widths=[1.3, 1.8, 2.4], rows=[
              ["Element", "Before", "After"],
              ["Title", "Sum of Sales by Zone", f"South-West led FY2025 with {Nm(SW.Sales)}"],
              ["Style", "3D, shaded, four colours", "Flat, one colour"],
              ["Time to answer", "Read every bar", "Read the title"]]),
          watch="A chart that needs a paragraph to explain it is the wrong chart."),

        T("3.7.2", "Chart Anatomy Overview",
          defn="Every Excel chart is built from the same elements: chart area, plot area, axes, gridlines, legend, data "
               "labels, titles and one or more data series.",
          simple="Click any part, then press Ctrl+1 to open its Format pane.",
          grid=dict(label="Parts of the FY2025 zone chart", widths=[1.4, 2, 1.8], rows=[
              ["Part", "In this chart", "Typical fix"],
              ["Data series", f"{len(Z)} columns, one per zone", "One colour"],
              ["Value axis", f"₦0 to ₦{int(np.ceil(SW.Sales / 50e6) * 50)}m", "Display units: Millions"],
              ["Chart title", "Top line", "State the finding"],
              ["Legend", "'Sales'", "Delete (one series)"]]),
          watch="Excel's default chart is a starting point, not a finished chart."),

        T("3.7.3", "Chart Area and Plot Area",
          defn="The chart area is the whole chart object; the plot area is the inner region where the series are drawn.",
          simple="Resize the plot area to make room for labels without shrinking the chart.",
          grid=dict(label="Which to select", widths=[1.3, 2, 2], rows=[
              ["Selection", "Contains", "Use it to"],
              ["Chart area", "Everything", "Move, size, set the font"],
              ["Plot area", "Series and gridlines only", "Make room for labels"]]),
          watch="Setting a font on the plot area does not change the titles; set fonts on the chart area."),

        T("3.7.4", "Category and Value Axes",
          defn="The category axis lists the groups or periods; the value axis shows the numeric scale.",
          simple="Start column and bar value axes at zero, and use Display units (Millions) to keep labels short.",
          grid=dict(label="South-West versus South-East, FY2025", widths=[1.8, 2, 1], rows=[
              ["Value axis starts at", "South-West bar looks", "Truth"],
              ["₦100m", f"{RATIO100:.1f} times longer", f"{RATIO0:.1f} times"],
              ["₦0", f"{RATIO0:.1f} times longer", f"{RATIO0:.1f} times"]]),
          watch="A truncated value axis on a column chart exaggerates differences; it is the commonest misleading chart."),

        T("3.7.5", "Gridlines and Tick Marks",
          defn="Gridlines are reference lines across the plot area; tick marks are the small marks on an axis at each interval.",
          simple="Use light major gridlines only, or none when every column has a data label.",
          grid=dict(label="Clean-up", widths=[1.6, 1.4, 1.4], rows=[
              ["Element", "Before", "After"],
              ["Minor gridlines", "Shown", "Removed"],
              ["Major gridlines", "Dark", "Light grey"],
              ["Tick marks", "Every month", "None"]]),
          watch="Dark gridlines compete with the data for attention."),

        T("3.7.6", "The Legend",
          defn="The legend is the key that links each colour or marker to its series.",
          simple="With one series, delete it. With two or three, label the series directly at the end of each line.",
          grid=dict(label="When to keep a legend", widths=[2, 1, 1.6], rows=[
              ["Series in chart", "Legend?", "Instead"],
              ["1 (zone sales)", "No", "The title says it"],
              ["2 (sales + margin)", "Optional", "Direct labels"],
              [f"{len(Z)}+ (all zones by month)", "Yes", "Legend at the top"]]),
          watch="A legend far from the data makes the eye jump back and forth for every point."),

        T("3.7.7", "Data Labels, Data Tables and Value Readability",
          defn="Data labels print each point's value on the chart; a data table shows the values in a grid under the chart.",
          simple="Label every point when there are few; with many, label only the ones that matter, such as the peak.",
          grid=dict(label="How much to label", widths=[2, 0.8, 2.6], size=11.5, rows=[
              ["Chart", "Points", "Labels"],
              ["Zone sales (column)", str(len(Z)), ", ".join(Nm(x) for x in ZS.Sales)],
              ["Monthly sales (line)", str(len(SMONTH)), f"Peak only: {ym(PEAK)} {Nm(SMONTH.max())}"]]),
          watch="Labels with too many digits (₦250,128,366) are unreadable; show ₦250.1m."),

        T("3.7.8", "Titles, Subtitles and Annotations",
          defn="The title states the finding, the subtitle states the scope, and an annotation points to the feature being discussed.",
          simple="Together they tell the reader what to see, in what data, and where.",
          grid=dict(label="Monthly sales line", widths=[1.2, 3.6], rows=[
              ["Text", "Wording"],
              ["Title", f"Sales build to a late-year peak; {ym(PEAK)} was the best month"],
              ["Subtitle", f"SuperMart Stores, monthly sales {SMONTH.index.min().year}–{SMONTH.index.max().year}, ₦m"],
              ["Annotation", f"{ym(PEAK)}: {Nm(SMONTH.max())}"]]),
          watch="A title that only names the data ('Sales by Month') leaves the reader to find the message."),

        T("3.7.9", "Choosing the Right Chart Type",
          defn="Chart choice follows the question: comparison, change over time, part of a whole, relationship or distribution.",
          simple="Compare = column or bar; time = line; part of whole = pie or stacked bar; relationship = scatter; spread = histogram or box plot.",
          grid=dict(label="Question, chart, data", widths=[2.4, 1, 2], size=11.5, rows=[
              ["Question", "Chart", "Data"],
              ["Which store sold most?", "Bar", f"{len(DS)} SuperMart stores"],
              ["How did sales move monthly?", "Line", f"{len(SMONTH)} months"],
              ["Does study time relate to score?", "Scatter", f"{I(len(CV))} Kingsway students"],
              ["How long do patients stay?", "Histogram", f"{I(len(LC))} admissions"]]),
          watch="Pick the chart for the question, not for variety."),

        T("3.7.10", "Column Chart",
          defn="A column chart compares values across categories with vertical bars.",
          simple="Clustered columns compare two or three series side by side; stacked columns show parts of a total.",
          grid=dict(label="FY2025 by zone (tDash)", widths=[1.4, 1.6, 0.9], rows=zone_rows()),
          steps=["Select the zone summary", "Insert › Column › Clustered Column", "Delete the legend; title states the finding",
                 "Value axis: Display units › Millions"],
          order=["grid", "steps"], split=0.52,
          watch="Long category names on a column chart turn diagonal; switch to a bar chart."),

        T("3.7.11", "Bar Chart",
          defn="A bar chart compares values across categories with horizontal bars.",
          simple="Use it when names are long or there are many categories. Sort the bars, largest at the top.",
          grid=dict(label="FY2025 sales by store", widths=[2.4, 1.6], rows=[
              ["Store", "FY2025 sales"]] + [[f"{r.Store_Name}" + (" (top)" if i == 0 else ""), N(r.Sales)]
                                             for i, r in enumerate(DSS.head(3).itertuples())] + [
              [f"{DSS.iloc[-1].Store_Name} (bottom)", N(DSS.iloc[-1].Sales)]]),
          watch="Excel draws the first row at the bottom: tick 'Categories in reverse order' to put the largest on top."),

        T("3.7.12", "Line Chart",
          defn="A line chart shows how a value changes over an ordered, usually time-based, category axis.",
          simple="It reveals trends, seasons and peaks better than any other chart.",
          grid=dict(label="Monthly sales, lowest and highest month", widths=[0.8, 1.6, 1.6], rows=[
              ["Year", "Lowest month", "Highest month"]] + [[str(y), *yr_lohi(y)] for y in YEARS.index]),
          watch="A line chart on unordered categories (such as zones) implies a trend that does not exist."),

        T("3.7.13", "Area Chart",
          defn="An area chart is a line chart with the area beneath the line filled, emphasising volume.",
          simple="A stacked area shows parts adding up over time, but only the bottom layer is easy to read.",
          grid=dict(label="Annual sales", widths=[1, 1.8], rows=[["Year", "Sales"]] + [[str(y), N(x)] for y, x in YEARS.items()]),
          watch="In a stacked area, the upper layers' changes are hard to judge; use lines for exact comparisons."),

        T("3.7.14", "Pie Chart",
          defn="A pie chart shows the parts of one whole as slices of a circle.",
          simple="It works only with a few slices (about five) that add up to 100%. Otherwise use a bar chart.",
          grid=dict(label="SuperMart Online customers by tier", widths=[1.2, 1.2, 1], rows=[
              ["Tier", "Customers", "Share"]] + [[t, I(n), P(n / len(CU))] for t, n in TIER.items()]),
          watch=f"Slices of {P(TIER.iloc[0] / len(CU))} and {P(TIER.iloc[1] / len(CU))} look identical in a pie; a bar shows the order."),

        T("3.7.15", "Doughnut Chart",
          defn="A doughnut chart is a pie chart with a hole in the centre, which can hold a total or a KPI.",
          simple="It shows nothing more than a pie; use it for KPI cards, not detailed comparison.",
          grid=dict(label="Order value by payment method", widths=[1.8, 1.6, 0.9], rows=[
              ["Payment method", "Value", "Share"]] + [[k, N(x), P(x / PAY.sum())] for k, x in list(PAY.items())[:4]]),
          watch=f"With {len(PAY)} methods the small slices become unreadable; show the top four and 'Other'."),

        T("3.7.16", "Scatter Chart",
          defn="A scatter chart plots pairs of numeric values as points, showing the relationship between two variables.",
          simple="It shows correlation, clusters and outliers. Both axes are numbers.",
          grid=dict(label=f"Exam score against … ({I(len(CV))} students)", widths=[2, 1.2, 1.6], rows=[
              ["Variable", "Correlation r", "Reading"],
              ["Attendance %", f"{R['Lecture_Attendance_Pct']:.2f}", "Moderate, positive"],
              ["Hours studied per week", f"{R['Study_Hours_Wk']:.2f}", "Moderate, positive"],
              ["Sleep hours", f"{R['Sleep_Hrs_Night']:.2f}", "No relationship"]]),
          watch="Correlation is not cause: attendance and score may both reflect motivation."),

        T("3.7.17", "Bubble Chart",
          defn="A bubble chart is a scatter chart in which a third numeric variable sets the size of each point.",
          simple="Use it only when the third number matters; bubble sizes are hard to compare precisely.",
          grid=dict(label="Kingsway: one bubble per department", widths=[1.6, 1.4, 2], rows=[
              ["Variable", "Shown as", "Read it as"],
              ["Hours studied (avg)", "Across", "Further right = more study"],
              ["Exam score (avg)", "Up", "Higher = better"],
              ["Students", "Bubble size", "Bigger = more students"]]),
          watch="Excel sizes bubbles by area by default; switching to width exaggerates big values."),

        T("3.7.18", "Radar Chart",
          defn="A radar chart plots several measures on spokes from a centre, drawing one profile per item.",
          simple="Use it for a few items and a few measures on the same scale; here each is a share of the best zone.",
          grid=dict(label="FY2025 zone profile (1.00 = best zone)", widths=[1.5] + [0.9] * 4, size=11.5, rows=[
              ["Zone", "Sales", "Profit", "Orders", "Margin"]] + [[z] + [f"{RAD.loc[z, m]:.2f}" for m in ("Sales", "Profit", "Orders", "Margin")] for z in Z]),
          watch="Radar charts with many spokes or items become tangled; a small table often reads better."),

        T("3.7.19", "Combo Chart with Secondary Axis",
          defn="A combo chart draws two chart types together, with a secondary value axis for a series on a different scale.",
          simple="Columns for sales, a line for margin on a % secondary axis. Label both axes.",
          grid=dict(label="FY2025 by zone", widths=[1.4, 1.6, 0.9], rows=zone_rows()),
          steps=["Insert › Combo › Clustered Column – Line on Secondary Axis", "Sales = column; Margin = line, secondary",
                 "Title both value axes (₦m and %)"],
          order=["grid", "steps"], split=0.52,
          watch=f"South-East ({P(ZS.loc['South-East', 'Margin'])}) beats North-Central ({P(ZS.loc['North-Central', 'Margin'])}) "
                "on margin despite lower sales; unlabelled axes hide that."),

        T("3.7.20", "Waterfall Chart",
          defn="A waterfall chart shows how a starting value becomes an ending value through a series of increases and decreases.",
          simple="The standard chart for explaining a change, such as this year's profit against last year's.",
          grid=dict(label="Profit bridge FY2024 → FY2025", widths=[2.2, 1.6], rows=[
              ["Step", "Amount"], ["FY2024 profit", N(P24.sum())]] + [
              [f"{c} change", ("+" if P25[c] - P24[c] >= 0 else "") + N(P25[c] - P24[c])] for c in P24.index] + [
              ["FY2025 profit", N(P25.sum())]]),
          watch="Right-click the last bar › Set as total; otherwise Excel draws it as another increase."),

        T("3.7.21", "Funnel Chart",
          defn="A funnel chart shows values through sequential stages, each a subset of the one before.",
          simple="Use it when stages happen in order, such as order placed, shipped, delivered.",
          grid=dict(label="SuperMart Online order funnel", widths=[2.2, 1, 1.2], rows=[
              ["Stage", "Lines", "% of placed"],
              ["Orders placed", I(N_LINES), "100.0%"],
              ["Not cancelled (shipped)", I(SHIPPED), P(SHIPPED / N_LINES)],
              ["Delivered or completed", I(NETL), P(NETL / N_LINES)]]),
          watch="Stages that are not subsets of each other do not belong in a funnel."),

        T("3.7.22", "Treemap Chart",
          defn="A treemap shows a hierarchy as nested rectangles whose areas are proportional to their values.",
          simple="Best for many categories when the question is 'which parts are biggest?'",
          grid=dict(label="Net revenue by subcategory", widths=[1.4, 1.4, 1.6], rows=[
              ["Subcategory", "Category", "Net revenue"]] + [[s, c, N(x)] for (c, s), x in list(SUB.items())[:4]]),
          watch="Small rectangles cannot be labelled; treemaps show the big parts, not exact values."),

        T("3.7.23", "Sunburst Chart",
          defn="A sunburst chart shows a hierarchy as concentric rings: the inner ring is the top level, outer rings the detail.",
          simple="Use it for two or three levels, such as category then subcategory.",
          grid=dict(label="Net revenue rings", widths=[0.8, 1.6, 2.6], rows=[
              ["Ring", "Level", "Largest"],
              ["Inner", f"Category ({NCAT})", f"{CATN.index[0]} {Nm(CATN.iloc[0])}"],
              ["Outer", f"Subcategory ({NSUB})", f"{SUB.index[0][1]} {Nm(SUB.iloc[0])}, {SUB.index[1][1]} {Nm(SUB.iloc[1])}"]]),
          watch="Outer-ring segments are hard to compare by angle; use a sorted bar for precise ranking."),

        T("3.7.24", "Histogram Chart",
          defn="A histogram shows the distribution of a numeric variable by counting values in consecutive ranges (bins).",
          simple="The bars touch, because the axis is a continuous scale.",
          grid=dict(label="Mercyland length of stay", widths=[1.2, 1.2], rows=[
              ["Days", "Admissions"]] + [[str(k), I(n)] for k, n in BINS.items()]),
          watch="Bin width changes the picture; state it, and try two widths before choosing."),

        T("3.7.25", "Box and Whisker Chart",
          defn="A box and whisker plot shows a distribution's median, quartiles, range and outliers, one box per group.",
          simple="It is the best chart for comparing distributions side by side.",
          grid=dict(label="Length of stay (days) by department", widths=[1.8, 0.7, 0.9, 0.7], rows=[
              ["Department", "Q1", "Median", "Q3"]] + [[d, F1(r[0.25]), F1(r[0.5]), F1(r[0.75])] for d, r in BOX.head(5).iterrows()]),
          watch="Compare medians and box heights: Surgery stays run longest; Obstetrics & Gynaecology stays are short and tightly grouped."),

        T("3.7.26", "Sparklines",
          defn="A sparkline is a small chart drawn inside one cell: a line, columns or win/loss.",
          simple="Put sparklines next to the numbers so readers see each row's trend without a full chart.",
          grid=dict(label="FY2025 monthly sales per store", widths=[2.2, 1, 1, 1.2], rows=[
              ["Store", "Jan", "Dec", "Shape"]] + [
              [s, Nm(SP.loc[s, 1]), Nm(SP.loc[s, 12]),
               "Up" if SP.loc[s, 12] > SP.loc[s, 1] * 1.2 else "Down" if SP.loc[s, 12] < SP.loc[s, 1] * 0.8 else "Flat"]
              for s in DSS.Store_Name.head(4)]),
          steps=["Select the 12 monthly cells of the first store", "Insert › Sparklines › Line",
                 "Location: the Trend cell › OK; fill down", "Sparkline tab › tick High Point"],
          order=["grid", "steps"], split=0.55,
          watch="By default each sparkline has its own scale; set Axis › Same for All to compare rows."),
    ],
    takeaways=["One question, one message, a title that states the finding.",
               "Format every part with Ctrl+1; start value axes at zero.",
               "Label few points directly; delete single-series legends.",
               "Comparison, time, part-whole, relationship, distribution: each has its chart.",
               "17 types, from column to sparkline; each has a trap."],
)

# ---------------------------------------------------------------- charts (edition 3): every series comes from the data above
FM = '₦#,##0,,"m"'                         # millions, the Display units › Millions look
MON3 = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
ZM = S25.assign(M=S25.Order_Date.dt.month).pivot_table(index="M", columns="Zone", values="Sales", aggfunc="sum").reindex(range(1, 13)).fillna(0)
SM_LAB = [p.strftime("%b %y") for p in SMONTH.index]
PEAK_I = list(SMONTH.index).index(PEAK)
BYYR = {y: [float(SMONTH.get(pd.Period(f"{y}-{m:02d}", "M"), 0)) for m in range(1, 13)] for y in YEARS.index}
DEPT = CV.groupby("Department").agg(H=("Study_Hours_Wk", "mean"), S=("Final_Score", "mean"), n=("Final_Score", "size")).sort_values("n", ascending=False)


def _whisk(d):
    x = LC.loc[LC.Department == d, "LOS_Days"]
    q1, q3 = x.quantile(.25), x.quantile(.75)
    lo, hi = x[x >= q1 - 1.5 * (q3 - q1)].min(), x[x <= q3 + 1.5 * (q3 - q1)].max()
    return d, float(lo), float(q1), float(x.median()), float(q3), float(hi)


TOPSUB = list(SUB.items())[:11]
OTHER = float(SUB.sum() - sum(v for _, v in TOPSUB))
CHARTS = {
    "3.7.1": dict(kind="column", cats=Z, series=[("Sales", ZS.Sales.tolist())], fmt=FM, labels=True, legend=False,
                  title=f"South-West led FY2025 with {Nm(SW.Sales)}", label="AFTER: ONE COLOUR, FINDING IN THE TITLE"),
    "3.7.2": dict(kind="column", cats=Z, series=[("Sales", ZS.Sales.tolist())], fmt=FM, labels=True, legend=True, y_title="₦m", x_title="Zone",
                  title="Chart title: FY2025 sales by zone", label="EVERY PART SWITCHED ON"),
    "3.7.4": dict(kind="column", cats=Z, series=[("Sales", ZS.Sales.tolist())], fmt=FM, legend=False, y_min=100e6, colors=["B42B1E"],
                  title=f"Axis from ₦100m: South-West looks {RATIO100:.1f}× South-East", label="THE TRAP: TRUNCATED VALUE AXIS",
                  note=f"The truth is {RATIO0:.1f}×. Set the axis minimum back to 0."),
    "3.7.6": dict(kind="line", cats=MON3, series=[(z, ZM[z].tolist()) for z in Z], fmt=FM, legend=True, h=2.6,
                  title="FY2025 monthly sales by zone: four series need a key"),
    "3.7.7": dict(kind="column", cats=Z, series=[("Sales", ZS.Sales.tolist())], fmt=FM, label_fmt='₦#,##0.0,,"m"', labels=True, legend=False,
                  title="Four points: label every one", h=2.5),
    "3.7.8": dict(kind="line", cats=SM_LAB, series=[("Sales", SMONTH.tolist())], fmt=FM, label_fmt='₦#,##0.0,,"m"', labels=[PEAK_I], legend=False,
                  title=f"Late-year peak; {ym(PEAK)} was the best month", h=2.6),
    "3.7.10": dict(kind="column", cats=Z, series=[("Sales", ZS.Sales.tolist())], fmt=FM, labels=True, legend=False,
                   title=f"South-West led FY2025 with {Nm(SW.Sales)}"),
    "3.7.11": dict(kind="bar", cats=DSS.Store_Name.tolist(), series=[("Sales", DSS.Sales.tolist())], fmt=FM, legend=False, h=3.0,
                   point_colors=[ORANGE] + [None] * (len(DSS) - 1), title=f"{DSS.iloc[0].Store_Name} sold most in FY2025"),
    "3.7.12": dict(kind="line", cats=MON3, series=[(str(y), v) for y, v in BYYR.items()], fmt=FM, legend=True,
                   title="Every year peaks late: one line per year"),
    "3.7.13": dict(kind="area", cats=MON3, series=[(z, ZM[z].tolist()) for z in Z], fmt=FM, legend=True, stacked=True,
                   title="FY2025 monthly sales, stacked by zone"),
    "3.7.14": dict(kind="pie", cats=TIER.index.tolist(), series=[("Customers", TIER.tolist())], title="Customers by membership tier", legend=True),
    "3.7.15": dict(kind="doughnut", cats=list(PAY.index[:4]) + ["Other"], series=[("Value", list(PAY.iloc[:4]) + [float(PAY.iloc[4:].sum())])],
                   title="Order value by payment method"),
    "3.7.16": dict(kind="scatter", points=[("Students", CV.Lecture_Attendance_Pct.tolist(), CV.Final_Score.tolist())], trend=True, marker=3,
                   x_title="Lecture attendance %", y_title="Final score", x_min=float(np.floor(CV.Lecture_Attendance_Pct.min() / 10) * 10), x_max=100,
                   title=f"Attendance and score: r = {R['Lecture_Attendance_Pct']:.2f}"),
    "3.7.17": dict(kind="bubble", points=[("Departments", DEPT.H.tolist(), DEPT.S.tolist(), DEPT.n.tolist())], scale=35,
                   x_title="Average study hours per week", y_title="Average final score", h=2.9,
                   title="Kingsway departments: study, score and size"),
    "3.7.18": dict(kind="radar", cats=["Sales", "Profit", "Orders", "Margin"], series=[(z, [float(RAD.loc[z, m]) for m in ("Sales", "Profit", "Orders", "Margin")]) for z in Z],
                   fmt="0.00", legend=True, title="Zone profile (1.00 = best zone)", h=2.9),
    "3.7.19": dict(kind="combo", cats=Z, series=[("Sales", ZS.Sales.tolist()), ("Margin", ZS.Margin.tolist())], fmt=FM, y2_fmt="0%",
                   title="Sales (columns) and margin (line, right axis)"),
    "3.7.20": dict(kind="waterfall", cats=["FY2024"] + [f"{c}" for c in P24.index] + ["FY2025"], start=float(P24.sum()),
                   steps=[float(P25[c] - P24[c]) for c in P24.index], fmt=FM, label_fmt='₦#,##0.0,,"m"', title="Profit bridge, FY2024 to FY2025", h=2.7),
    "3.7.21": dict(kind="funnel", cats=["Orders placed", "Shipped", "Delivered or completed"], values=[N_LINES, SHIPPED, NETL], label_fmt="#,##0",
                   title="SuperMart Online order funnel (lines)", h=2.3),
    "3.7.22": dict(kind="treemap", tree=[(sc, float(v)) for (c, sc), v in TOPSUB] + [("Other", OTHER)], title="Net revenue by subcategory", h=2.9),
    "3.7.23": dict(kind="sunburst", rings=[(c, [(sc, float(v)) for (cc, sc), v in SUB.items() if cc == c]) for c in CATN.index],
                   title="Category (inner) and subcategory (outer)", h=2.9),
    "3.7.24": dict(kind="histogram", cats=[str(k) for k in BINS.index], series=[("Admissions", BINS.tolist())], fmt="#,##0", legend=False,
                   x_title="Length of stay (days), bin width 5", title="Most stays are short; a long right tail"),
    "3.7.25": dict(kind="box", box=[_whisk(d) for d in BOX.index], fmt="0", y_title="Days", title="Length of stay by department (whiskers: 1.5 × IQR)", h=2.8),
    "3.7.26": dict(kind="sparklines", rows=[(s, SP.loc[s].tolist()) for s in DSS.Store_Name], ncol=2, label_w=1.95, h=3.0,
                   title="FY2025 monthly sales, one sparkline per store (orange = high point)", label="SPARKLINES FROM OUR DATA"),
}
ORDER = {"3.7.10": ["chart", "steps"], "3.7.19": ["chart", "steps"], "3.7.26": ["chart", "steps"]}
SPLIT = {"3.7.10": 0.58, "3.7.19": 0.58, "3.7.26": 0.62}
for t in M37["topics"]:
    if t["tag"] in CHARTS:
        t["chart"] = CHARTS[t["tag"]]
        if t["tag"] in ORDER:
            t["order"] = ORDER[t["tag"]]
            t["split"] = SPLIT[t["tag"]]

MODULES = [M37]
