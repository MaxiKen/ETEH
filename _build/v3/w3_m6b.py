"""Week 3 edition-3 slide copy: Module 3.6B (DAX time intelligence, KPIs and performance).
Every result is a pandas twin of the measure shown, on SuperMart Online (net revenue), SuperMart Stores FY2025,
Mercyland and Kingsway data (w3_fx)."""
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
AR = int(((CV.Final_Score < 62) | (CV.Lecture_Attendance_Pct < 65)).sum())
AR2 = int(((CV.Final_Score < 60) | (CV.Lecture_Attendance_Pct < 65)).sum())
READMIT = (LC.Readmitted_30d == "Yes").mean()

MEASURES = ["Gross Bookings", "Orders", "Net Revenue", "Net Orders", "Net Rev Big Orders", "Cat Share", "Share of Zone",
            "Categories Sold", "Revenue X", "Avg Net Spend per Customer", "AOV", "Order Size", "Comm %", "Rev YTD", "Rev LY",
            "YoY %", "Rev PM", "MoM %", "Rev R3M", "Rev Running", "Zone Share", "Seller Rank", "Rev 0", "Return Rate"]

M36B = dict(
    id="3.6B", title="Power Pivot and DAX: Time Intelligence and KPIs",
    subtitle="Year to date and last year, percentage of total, BLANK, tidy measures, three KPI sets and simple debugging.",
    objectives=["Write year-to-date and last-year measures on a marked date table.",
                "Build a percentage-of-total measure.",
                "Handle BLANK and zero deliberately.",
                "Organise and format explicit measures in one measures table.",
                "Define retail, hospital and university KPI sets, then debug them."],
    topics=[
        T("3.6.17", "Time Intelligence: TOTALYTD",
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

        T("3.6.18", "SAMEPERIODLASTYEAR",
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

        T("3.6.19", "Percentage of Total",
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

        T("3.6.20", "Handling BLANK and Zero",
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

        T("3.6.21", "Organising a Measures Table",
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

        T("3.6.22", "Format Strings for Measures",
          defn="A measure's format string sets how its values display wherever the measure is used.",
          simple="Set it once on the measure, not on each PivotTable: #,##0 for counts, 0.0% for rates, ₦#,##0 for money.",
          grid=dict(label="Formats", widths=[1.6, 1.2, 1.6], rows=[
              ["Measure", "Format", "Displays"],
              ["[Net Revenue]", "₦#,##0", N(NETREV)],
              ["[YoY %] (2025)", "0.0%", P(YOY25)],
              ["[Orders]", "#,##0", I(N_LINES)]]),
          watch="Formatting pivot cells instead of the measure is lost the next time the pivot is rebuilt."),

        T("3.6.23", "Implicit versus Explicit Measures",
          defn="An implicit measure is created by dragging a column into Values; an explicit measure is written in DAX, named "
               "and reusable.",
          simple="Professional models use explicit measures only: they can be reused, formatted and documented.",
          grid=dict(label="Two reports, two totals", widths=[0.8, 2.6, 1.6], rows=[
              ["Report", "Measure", "Value"],
              ["A", "Sum of total_price (implicit)", N(GROSS)],
              ["B", "[Net Revenue] (explicit)", N(NETREV)]]),
          watch="Report A includes cancelled and returned orders; an implicit sum cannot carry that rule."),

        T("3.6.24", "Retail KPI Measures",
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

        T("3.6.25", "Hospital KPI Measures",
          defn="A hospital KPI set measures flow and quality: admissions, average and median length of stay, 30-day "
               "readmission rate, bed occupancy and cost per admission.",
          simple="Report the median with the average: a few very long stays pull the average up.",
          grid=dict(label=f"Mercyland, {LC.Admit_Date.min():%b %Y} to {LC.Admit_Date.max():%b %Y}", widths=[2.4, 1.8], rows=[
              ["KPI", "Result"],
              ["Admissions", I(len(LC))],
              ["ALOS / median LOS", f"{LC.LOS_Days.mean():.1f} / {LC.LOS_Days.median():.0f} days"],
              ["30-day readmission rate", P(READMIT)],
              ["Average occupancy", P(CEN.Occupancy_Pct.mean())],
              ["Days above 95% full", f"{I((CEN.Occupancy_Pct > 0.95).sum())} of {I(len(CEN))}"],
              ["Cost per admission", N(LC.Bill_Amount_NGN.mean())]]),
          watch="Occupancy comes from the daily bed census, not from admissions; the grains differ, so keep separate fact tables."),

        T("3.6.26", "University KPI Measures",
          defn="A university KPI set measures progression: enrolment, average exam score, grade A rate, attendance and the "
               "number of students at risk.",
          simple="Write each definition down. At risk here means exam score below 62 or attendance below 65%.",
          grid=dict(label="Kingsway University, 2024/25", widths=[2.2, 1.6], rows=[
              ["KPI", "Result"],
              ["Enrolment", I(len(CV))],
              ["Average exam score", f"{CV.Final_Score.mean():.1f}"],
              ["Grade A rate (70+)", P((CV.Final_Score >= 70).mean())],
              ["Average attendance", P(CV.Lecture_Attendance_Pct.mean() / 100)],
              ["At-risk students", f"{I(AR)} ({P(AR / len(CV))})"]]),
          watch=f"Definitions change results: with 'below 60' instead of 'below 62', {I(AR)} at risk becomes {I(AR2)}."),

        T("3.6.27", "Debugging DAX Measures",
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

        T("3.6.28", "DAX Quick Reference",
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
               "YTD and last year are a measure plus one date function.",
               "Shares use ALL; BLANK hides rows, + 0 shows them.",
               "Explicit, formatted measures in one _Measures table.",
               "Debug with a simple pivot; verify with formula twins."],
)

MODULES = [M36B]
