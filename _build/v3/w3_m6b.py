"""Week 3 edition-3 slide copy: Module 3.6B (DAX time intelligence, KPIs and performance).
Every result is a pandas twin of the measure shown, on SuperMart Online (net revenue), SuperMart Stores FY2025,
Lifecare and Crestview data (w3_fx)."""
import pandas as pd

from v3kit import T
from w3_fx import (AOV, CAT, CEN, CODE, CV, FY25, GROSS, I, LC, LN, M24, M25, N, N_LINES, N_NET, NETREV, NT, ORDERS25, P,
                   PROFIT25, S25, SE, TARGET25, TX, YEAR, YTD25, ZONE, BUYERS, mon_full)

PREV = {1: float(NT[(NT.year == 2024) & (NT.month == 12)].total_price.sum())}
for m in range(2, 13):
    PREV[m] = float(M25[m - 1])
R3 = {m: float(M25[max(1, m - 2):m].sum()) if m >= 3 else None for m in range(1, 13)}
RUN = YEAR.cumsum()
SELLER = NT.groupby("seller_id").total_price.sum().sort_values(ascending=False)
SNAME = SE.set_index("seller_id").store_name
ACTIVE = TX.seller_id.nunique()
RET = LN.groupby("category").order_status.apply(lambda s: (s == "Returned").mean()).sort_values()
N_RET = int((TX.order_status == "Returned").sum())
NO_SALE = sorted(set(SE.seller_id) - set(TX.seller_id))
YOY25 = YEAR[2025] / YEAR[2024] - 1
AR = int(((CV.Exam_Score < 62) | (CV.Attendance_Pct < 65)).sum())
AR2 = int(((CV.Exam_Score < 60) | (CV.Attendance_Pct < 65)).sum())
READMIT = (LC.Readmit_30d == "Yes").mean()

MEASURES = ["Gross Bookings", "Orders", "Net Revenue", "Net Orders", "Net Rev Big Orders", "Cat Share", "Share of Zone",
            "Categories Sold", "Revenue X", "Avg Net Spend per Customer", "AOV", "Order Size", "Comm %", "Rev YTD", "Rev LY",
            "YoY %", "Rev PM", "MoM %", "Rev R3M", "Rev Running", "Zone Share", "Seller Rank", "Rev 0", "Return Rate"]

M36B = dict(
    id="3.6B", title="Power Pivot and DAX: Time Intelligence, KPIs and Performance",
    subtitle="Year to date, last year, rolling windows, ranks and KPI sets, then organising, formatting, debugging and documenting.",
    objectives=["Write YTD, prior-year, prior-month and rolling measures on a marked date table.",
                "Build running totals, percentage of total and ranks.",
                "Handle BLANK, zero and errors deliberately.",
                "Organise, format and document measures; prefer explicit measures.",
                "Define retail, hospital and university KPI sets, then debug and verify them."],
    topics=[
        T("3.6.22", "Time Intelligence: TOTALYTD",
          defn="TOTALYTD evaluates an expression from the start of the year to the last date in the current filter context, "
               "using a marked date table.",
          simple="Put DimDate months on rows: each row shows the total from 1 January to the end of that month.",
          **CODE("""
Rev YTD :=
TOTALYTD (
    [Net Revenue],
    tDimDate[Date]
)""", label="DAX MEASURE"),
          grid=dict(label="2025, by month", widths=[1.3, 1.6, 1.6], rows=[
              ["Month", "[Net Revenue]", "[Rev YTD]"]] + [
              [mon_full(m), N(M25[m]), N(YTD25[m])] for m in (3, 6, 9, 12)]),
          split=0.45,
          watch="Without a marked, gap-free date table related to the facts, YTD results are wrong or blank."),

        T("3.6.23", "SAMEPERIODLASTYEAR",
          defn="SAMEPERIODLASTYEAR returns the dates in the current filter context shifted back by one year.",
          simple="Use it inside CALCULATE for last year's figure, then DIVIDE for year-on-year growth. No year is typed in.",
          **CODE("""
Rev LY :=
CALCULATE ( [Net Revenue],
    SAMEPERIODLASTYEAR ( tDimDate[Date] ) )

YoY % :=
DIVIDE ( [Net Revenue] - [Rev LY], [Rev LY] )""", label="DAX MEASURES", size=11.5),
          grid=dict(label="By year", widths=[0.7, 1.5, 1.5, 0.9], size=11.5, rows=[
              ["Year", "[Net Revenue]", "[Rev LY]", "[YoY %]"]] + [
              [str(y), N(YEAR[y]), N(YEAR[y - 1]), P(YEAR[y] / YEAR[y - 1] - 1)] for y in (2023, 2024, 2025)]),
          split=0.47,
          watch=f"The first year ({YEAR.index.min()}) has no prior year, so [YoY %] is blank there; that is correct."),

        T("3.6.24", "DATEADD",
          defn="DATEADD shifts the dates in the current filter context by a number of days, months, quarters or years.",
          simple="DATEADD ( tDimDate[Date], -1, MONTH ) gives the previous month, for month-on-month change.",
          **CODE("""
Rev PM :=
CALCULATE ( [Net Revenue],
    DATEADD ( tDimDate[Date], -1, MONTH ) )

MoM % :=
DIVIDE ( [Net Revenue] - [Rev PM], [Rev PM] )""", label="DAX MEASURES", size=11.5),
          grid=dict(label="2025 months", widths=[1.1, 1.4, 1.4, 0.8], size=11.5, rows=[
              ["Month", "[Net Revenue]", "[Rev PM]", "[MoM %]"]] + [
              [mon_full(m), N(M25[m]), N(PREV[m]), P(M25[m] / PREV[m] - 1)] for m in (1, 2, 12)]),
          split=0.47,
          watch="January's previous month is December of the year before: the date table must cover it."),

        T("3.6.25", "DATESYTD and DATESINPERIOD",
          defn="DATESYTD returns the dates from the start of the year to the current date; DATESINPERIOD returns a window "
               "of a given length ending on a date.",
          simple="A rolling three-month total smooths out month-to-month noise.",
          **CODE("""
Rev R3M :=
CALCULATE ( [Net Revenue],
    DATESINPERIOD ( tDimDate[Date],
        MAX ( tDimDate[Date] ), -3, MONTH ) )""", label="DAX MEASURE", size=11.5),
          grid=dict(label="Rolling 3 months", widths=[1.3, 1.4, 1.5], rows=[
              ["Month shown", "Window", "[Rev R3M]"],
              ["June 2025", "Apr–Jun 2025", N(R3[6])],
              ["December 2025", "Oct–Dec 2025", N(R3[12])]]),
          split=0.52,
          watch="CALCULATE ( [M], DATESYTD ( tDimDate[Date] ) ) equals TOTALYTD; use whichever reads better."),

        T("3.6.26", "Running Totals",
          defn="A running total accumulates a measure over all dates up to the current one, without restarting each year.",
          simple="The pattern: remove the date filter with ALL, then keep dates up to the last date in view.",
          **CODE("""
Rev Running :=
CALCULATE ( [Net Revenue],
    FILTER ( ALL ( tDimDate ),
        tDimDate[Date] <= MAX ( tDimDate[Date] ) ) )""", label="DAX MEASURE", size=11.5),
          grid=dict(label="By year", widths=[0.7, 1.6, 1.6], rows=[
              ["Year", "[Net Revenue]", "[Rev Running]"]] + [
              [str(y), N(YEAR[y]), N(RUN[y])] for y in (2020, 2022, 2024, 2025)]),
          split=0.52,
          watch="The last running total must equal the grand total of the measure; that is the check."),

        T("3.6.27", "Percentage of Total",
          defn="A percentage of total divides a measure by the same measure evaluated with the relevant filters removed.",
          simple="DIVIDE ( [M], CALCULATE ( [M], ALL ( table ) ) ). The shares must add up to 100%.",
          **CODE("""
Zone Share :=
DIVIDE ( [Net Revenue],
    CALCULATE ( [Net Revenue],
        ALL ( tCustomers[zone] ) ) )""", label="DAX MEASURE"),
          grid=dict(label="By customer zone", widths=[1.4, 1.6, 0.9], rows=[
              ["Zone", "[Net Revenue]", "Share"]] + [
              [z, N(x), P(x / NETREV)] for z, x in list(ZONE.items())[:4]] + [["All 6 zones", N(NETREV), "100.0%"]]),
          split=0.48,
          watch="ALL on the wrong column gives 100% on every row; ALL on the whole fact table ignores every slicer."),

        T("3.6.28", "Ranking with RANKX",
          defn="RANKX ranks each item by the value of an expression evaluated over a table of items.",
          simple="Use ALL ( column ) so each item is ranked against all items. Ties: DENSE gives 1, 2, 2, 3; the default gives 1, 2, 2, 4.",
          like="A league table: every team placed by its points.",
          **CODE("""
Seller Rank :=
RANKX ( ALL ( tSellers[seller_id] ),
    [Net Revenue], , DESC, DENSE )""", label="DAX MEASURE"),
          grid=dict(label="Top sellers", widths=[0.5, 2.6, 1.5], size=11.5, rows=[
              ["#", "Seller", "[Net Revenue]"]] + [
              [str(i), f"{s} {SNAME[s]}", N(x)] for i, (s, x) in enumerate(list(SELLER.items())[:4], 1)]),
          split=0.45,
          watch="Rank sellers by seller_id: store names repeat, and ranking by name would merge different sellers."),

        T("3.6.29", "Handling BLANK and Zero",
          defn="BLANK is DAX's empty value, returned when there is no data; a PivotTable hides rows whose measures are all blank.",
          simple="Adding + 0 shows a zero instead. Use it only where seeing the zero matters.",
          **CODE("""
Rev 0 :=
[Net Revenue] + 0""", label="DAX MEASURE"),
          grid=dict(label="Sellers on PivotTable rows", widths=[1.4, 1.2, 1.8], rows=[
              ["Measure", "Rows shown", "Sellers with no net sales"],
              ["[Net Revenue]", I(SELLER.size), "Hidden (BLANK)"],
              ["[Rev 0]", I(len(SE)), "Shown as 0"]]),
          split=0.42,
          watch="+ 0 on a measure placed against a large dimension fills the pivot with zero rows; use it sparingly."),

        T("3.6.30", "Error Handling in DAX",
          defn="Defensive DAX prevents errors by design: DIVIDE for ratios, ISBLANK tests and filtering out invalid rows, "
               "rather than hiding errors with IFERROR.",
          simple="IFERROR hides real mistakes as well as expected ones, so use it sparingly.",
          **CODE("""
Return Rate :=
DIVIDE (
    CALCULATE ( [Orders],
        tTransactions[order_status] = "Returned" ),
    [Orders], 0 )""", label="DAX MEASURE"),
          grid=dict(label="[Return Rate]", widths=[1.8, 1, 1.4], rows=[
              ["Filter", "Rate", "Note"],
              ["All lines", P(N_RET / N_LINES), f"{I(N_RET)} of {I(N_LINES)}"],
              [f"{RET.index[-1]} (highest)", P(RET.iloc[-1]), ""],
              [f"{RET.index[0]} (lowest)", P(RET.iloc[0]), ""],
              [f"{NO_SALE[0]} (no orders)", "0.0%", "DIVIDE alternate"]]),
          split=0.48,
          watch="An IFERROR that turns a broken relationship into 0 makes a wrong report look right."),

        T("3.6.31", "Organising a Measures Table",
          defn="A measures table is an empty table that holds only measures, so the field list has one place for every calculation.",
          simple="Create a one-column table called _Measures, move every measure into it and hide the dummy column.",
          grid=dict(label="Practice and result", widths=[2.4, 2.2], rows=[
              ["Practice", "Result"],
              ["Every measure in _Measures", f"{len(MEASURES)} measures from this module in one place"],
              ["Name prefixes (Rev, Orders, Share)", "Easy to scan"],
              ["Dummy column hidden", "Only measures show"]]),
          steps=["Enter one cell 'x' › Ctrl+T › name it _Measures", "Add it to the Data Model",
                 "Each measure: Table name = _Measures", "Hide the column x"],
          order=["steps", "grid"], split=0.45,
          watch="Measures left in fact tables are hard to find, and get deleted with the table."),

        T("3.6.32", "Format Strings for Measures",
          defn="A measure's format string sets how its values display wherever the measure is used.",
          simple="Set it once on the measure, not on each PivotTable: #,##0 for counts, 0.0% for rates, ₦#,##0 for money.",
          grid=dict(label="Formats", widths=[1.6, 1.2, 1.6], rows=[
              ["Measure", "Format", "Displays"],
              ["[Net Revenue]", "₦#,##0", N(NETREV)],
              ["[YoY %] (2025)", "0.0%", P(YOY25)],
              ["[Orders]", "#,##0", I(N_LINES)]]),
          watch="Formatting pivot cells instead of the measure is lost the next time the pivot is rebuilt."),

        T("3.6.33", "Implicit versus Explicit Measures",
          defn="An implicit measure is created by dragging a column into Values; an explicit measure is written in DAX, named "
               "and reusable.",
          simple="Professional models use explicit measures only: they can be reused, formatted and documented.",
          grid=dict(label="Two reports, two totals", widths=[0.8, 2.6, 1.6], rows=[
              ["Report", "Measure", "Value"],
              ["A", "Sum of total_price (implicit)", N(GROSS)],
              ["B", "[Net Revenue] (explicit)", N(NETREV)]]),
          watch="Report A includes cancelled and returned orders; an implicit sum cannot carry that rule."),

        T("3.6.34", "DAX Performance Basics",
          defn="DAX performance depends mostly on model design: narrow fact tables, few calculated columns on large tables "
               "and simple CALCULATE filters.",
          simple="Iterators such as FILTER over a big fact table are the usual slow spot.",
          grid=dict(label="Slow pattern and better pattern", widths=[2.4, 2.2], rows=[
              ["Slow", "Better"],
              ["FILTER ( tTransactions, … ) for a simple test", "CALCULATE with a column filter"],
              ["Calculated column for a total", "A measure"],
              ["Unused columns in the fact table", "Remove them in Power Query"],
              ["Many-to-many or bidirectional filters", "Single-direction star"]]),
          watch="Speed problems are rarely fixed by clever DAX; fix the model shape first."),

        T("3.6.35", "Retail KPI Measures",
          defn="A retail KPI set measures trading: sales, profit, margin, orders, basket value, units per basket and target attainment.",
          simple="Each KPI is one documented measure on the SuperMart Stores FY2025 lines (tDash).",
          grid=dict(label="SuperMart Stores, FY2025", widths=[2, 1.6, 2.2], rows=[
              ["KPI", "Result", "Definition"],
              ["Sales", N(FY25), "SUM ( Sales )"],
              ["Margin %", P(PROFIT25 / FY25, 2), "Profit ÷ Sales"],
              ["Orders / basket value", f"{I(ORDERS25)} / {N(FY25 / ORDERS25)}", "Distinct orders; Sales ÷ Orders"],
              ["Units per basket", f"{S25.Quantity.sum() / ORDERS25:.1f}", "Quantity ÷ Orders"],
              ["Target attainment", P(FY25 / TARGET25), "Sales ÷ Annual target"]]),
          watch="Basket value uses distinct orders, not lines: one order has several lines."),

        T("3.6.36", "Hospital KPI Measures",
          defn="A hospital KPI set measures flow and quality: admissions, average and median length of stay, 30-day "
               "readmission rate, bed occupancy and cost per admission.",
          simple="Report the median with the average: a few very long stays pull the average up.",
          grid=dict(label=f"Lifecare, {LC.Admission_Date.min():%b %Y} to {LC.Admission_Date.max():%b %Y}", widths=[2.4, 1.8], rows=[
              ["KPI", "Result"],
              ["Admissions", I(len(LC))],
              ["ALOS / median LOS", f"{LC.Length_of_Stay.mean():.1f} / {LC.Length_of_Stay.median():.0f} days"],
              ["30-day readmission rate", P(READMIT)],
              ["Average occupancy", P(CEN.Occupancy_Pct.mean())],
              ["Days above 95% full", f"{I((CEN.Occupancy_Pct > 0.95).sum())} of {I(len(CEN))}"],
              ["Cost per admission", N(LC.Billing_Amount.mean())]]),
          watch="Occupancy comes from the daily bed census, not from admissions; the grains differ, so keep separate fact tables."),

        T("3.6.37", "University KPI Measures",
          defn="A university KPI set measures progression: enrolment, average exam score, grade A rate, attendance and the "
               "number of students at risk.",
          simple="Write each definition down. At risk here means exam score below 62 or attendance below 65%.",
          grid=dict(label="Crestview University, 2024/25", widths=[2.2, 1.6], rows=[
              ["KPI", "Result"],
              ["Enrolment", I(len(CV))],
              ["Average exam score", f"{CV.Exam_Score.mean():.1f}"],
              ["Grade A rate (70+)", P((CV.Exam_Score >= 70).mean())],
              ["Average attendance", P(CV.Attendance_Pct.mean() / 100)],
              ["At-risk students", f"{I(AR)} ({P(AR / len(CV))})"]]),
          watch=f"Definitions change results: with 'below 60' instead of 'below 62', {I(AR)} at risk becomes {I(AR2)}."),

        T("3.6.38", "Debugging DAX Measures",
          defn="Debugging a measure means testing it in a simple PivotTable, replacing parts with simpler measures and "
               "checking relationships until the wrong part is found.",
          simple="The commonest bug: the same total on every row, because the table on rows is not related to the facts.",
          grid=dict(label="tDimDate[Year] on rows, [Net Revenue] in values", widths=[0.8, 1.8, 1.8], rows=[
              ["Year", "No relationship", "After relating order_date"]] + [
              [str(y), N(NETREV), N(YEAR[y])] for y in (2020, 2024, 2025)]),
          steps=["Put one field on rows and the measure in values", "Swap the measure for COUNTROWS: does it change by row?",
                 "Check Diagram View: is the relationship there and active?", "Fix, then compare with a SUMIFS twin"],
          order=["steps", "grid"], split=0.5,
          watch="Excel also shows a yellow 'Relationships between tables may be needed' bar; read it."),

        T("3.6.39", "Performance Analyzer and Model Verification",
          defn="Verification proves a measure returns the right number, by comparing it with an independent calculation of "
               "the same quantity.",
          simple="Excel has no Performance Analyzer (that is in Power BI Desktop); the free DAX Studio times Excel queries. "
                 "Verify with SUMIFS twins on the data sheets.",
          grid=dict(label="Measure versus formula twin", widths=[1.6, 1.6, 1.6, 0.8], size=11.5, rows=[
              ["Measure", "DAX", "SUMIFS / COUNTIFS twin", "Match"],
              ["[Net Revenue]", N(NETREV), N(TX[TX.order_status.isin(["Delivered", "Completed"])].total_price.sum()), "Yes"],
              ["[Net Revenue] 2025", N(YEAR[2025]), N(NT[NT.order_date.dt.year == 2025].total_price.sum()), "Yes"],
              ["[Orders]", I(N_LINES), I(len(TX)), "Yes"],
              ["Customers", I(BUYERS), I(TX.customer_id.nunique()), "Yes"]]),
          watch="A measure that has never been compared with an independent total has not been tested."),

        T("3.6.40", "Documenting a DAX Model",
          defn="Model documentation records, for every measure, its name, plain-English definition, grain, source, owner "
               "and the date it was last checked.",
          simple="Put it in the measure's Description box and on a Measures sheet in 00_README.",
          grid=dict(label="Entry for [Net Revenue]", widths=[1.4, 3.2], rows=[
              ["Field", "Entry"],
              ["Definition", "Sum of total_price for Delivered and Completed lines"],
              ["Grain / source", "Order line / tTransactions"],
              ["Expected total", N(NETREV)],
              ["Owner / checked", "BI analyst / date of the last twin check"]]),
          watch="'Revenue' with no definition will be read three ways by three managers."),

        T("3.6.41", "DAX Quick Reference",
          defn="The core of DAX: context first, CALCULATE to change filters, iterators for row-by-row work, time intelligence "
               "on a marked date table, and DIVIDE for every ratio.",
          simple="Most measures in practice are one of these five patterns.",
          grid=dict(label="Pattern and result in our model", widths=[1.2, 2.4, 1.8], size=11.5, rows=[
              ["Need", "Pattern", "Result"],
              ["Total", "SUM ( column )", N(GROSS)],
              ["Filtered total", "CALCULATE ( [M], filter )", N(NETREV)],
              ["Year to date", "TOTALYTD ( [M], tDimDate[Date] )", N(YTD25[12]) + " (Dec 2025)"],
              ["Growth", "DIVIDE ( [M] - [LY], [LY] )", P(YOY25) + " (2025)"],
              ["Share", "DIVIDE ( [M], CALCULATE ( [M], ALL ( … ) ) )", P(CAT.iloc[0] / NETREV) + " Electronics"]]),
          watch="When a number looks wrong, ask first: what filter context is this cell in?"),
    ],
    takeaways=["Time intelligence needs a marked, gap-free date table.",
               "YTD, LY, PM and rolling windows are CALCULATE plus a date function.",
               "Shares use ALL; ranks use RANKX over ALL; BLANK hides rows.",
               "Explicit, formatted, documented measures in one _Measures table.",
               "Debug with a simple pivot; verify with formula twins."],
)

MODULES = [M36B]
