"""Week 3 edition-3.3 slide copy: Module 3.6B (KPIs and time comparisons with PivotTables, no DAX).
Every result is computed in pandas exactly as the PivotTable setting shown returns it, on SuperMart Online (net revenue),
SuperMart Stores FY2025, Mercyland and Kingsway data (w3_fx)."""
import pandas as pd

from v3kit import T
from w3_fx import (AOV, CAT, CEN, CV, FY25, GROSS, I, LC, LN, M24, M25, N, N_LINES, N_NET, NETREV, NT, ORDERS25, P,
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

M36B = dict(
    id="3.6B", title="KPIs and Time Comparisons with PivotTables",
    subtitle="Years and months from the date table, year to date, growth on last year, shares, tidy value fields, three "
             "KPI sets and how to check a PivotTable. All with PivotTable settings, no formula language.",
    objectives=["Put years and months from the date table on a PivotTable.",
                "Show year to date with Running Total In, and growth with % Difference From.",
                "Show shares with % of Column Total, and show empty cells as 0.",
                "Name and format value fields so the PivotTable reads clearly.",
                "Define retail, hospital and university KPI sets, then check the numbers."],
    topics=[
        T("3.6.15", "Years and Months from the Date Table",
          defn="With tDimDate related to tTransactions on the order date, its Year and Month fields can go on Rows or "
               "Columns of any Data Model PivotTable.",
          simple="Use the date table's fields, not the order date itself: every month appears, in calendar order.",
          grid=dict(label="Filter sale_type = Net; Rows: tDimDate Year", widths=[0.9, 1.8, 1], rows=[
              ["Year", "Net revenue", "Lines"]] + [
              [str(y), N(YEAR[y]), I((NT.year == y).sum())] for y in (2023, 2024, 2025)]),
          steps=["Rows: tDimDate › Year, then Month", "Month sorts Jan to Dec because it is sorted by Month_No (3.5)",
                 "Collapse or expand a year with the − / + buttons"],
          order=["grid", "steps"], split=0.5,
          watch="If every year shows the same total, the date table is not related to the fact table (3.6.24)."),

        T("3.6.16", "Year to Date with Running Total In",
          defn="Show Values As › Running Total In adds each month to the months before it, giving a year-to-date figure.",
          simple="With Year above Month on Rows (or Year in Filters), the running total starts again each January.",
          grid=dict(label="2025, by month", widths=[1.3, 1.6, 1.6], rows=[
              ["Month", "Net revenue", "Year to date"]] + [
              [mon_full(m), N(M25[m]), N(YTD25[m])] for m in (3, 6, 9, 12)]),
          steps=["Values: total_price twice; Filter: Year = 2025", "Second copy › Show Values As › Running Total In",
                 "Base field: Month › OK", "Rename it Year to date"],
          order=["grid", "steps"], split=0.5,
          watch="Pick Base field Month, not Year. Check: December's year-to-date must equal the year's total."),

        T("3.6.17", "Growth on Last Year with % Difference From",
          defn="Show Values As › % Difference From compares each item with another item of the same field, such as the "
               "previous year.",
          simple="Choose Base field Year and Base item (previous). No year is typed in, so it keeps working next year.",
          grid=dict(label="Rows: Year; values shown two ways", widths=[0.7, 1.5, 1.4, 0.9], size=11.5, rows=[
              ["Year", "Net revenue", "Difference From", "% Diff."]] + [
              [str(y), N(YEAR[y]), N(YEAR[y] - YEAR[y - 1]), P(YEAR[y] / YEAR[y - 1] - 1)] for y in (2023, 2024, 2025)]),
          steps=["Values: total_price twice", "Second copy › Show Values As › % Difference From",
                 "Base field: Year; Base item: (previous)"],
          order=["grid", "steps"], split=0.55,
          watch=f"The first year ({YEAR.index.min()}) has no previous year, so its cell is blank; that is correct."),

        T("3.6.18", "Share by Zone with % of Column Total",
          defn="% of Column Total shows each row as a share of its column's total; with one value column it equals % of "
               "Grand Total.",
          simple="Use it when Columns holds years: each year's zone shares then add up to 100% on their own.",
          grid=dict(label="Filter sale_type = Net; Rows: customer zone", widths=[1.4, 1.6, 0.9], rows=[
              ["Zone", "Net revenue", "Share"]] + [
              [z, N(x), P(x / NETREV)] for z, x in list(ZONE.items())[:4]] + [["All 6 zones", N(NETREV), "100.0%"]]),
          steps=["Rows: tCustomers › zone", "Values: total_price twice",
                 "Second copy › Show Values As › % of Column Total"],
          order=["grid", "steps"], split=0.55,
          watch="% of Grand Total with years on Columns gives shares of all years together, which is rarely what you want."),

        T("3.6.19", "Empty Cells and Items with No Data",
          defn="A PivotTable leaves a cell empty when no rows match, and hides items that have no data at all.",
          simple="PivotTable Options can show 0 in empty cells and list items with no data, for example sellers with no sales.",
          grid=dict(label="Sellers on Rows, sale_type = Net", widths=[2.1, 1.1, 1.6], rows=[
              ["Setting", "Rows shown", "Sellers with no net sales"],
              ["Default", I(SELLER.size), "Hidden"],
              ["Show items with no data", I(len(SE)), "Listed, cell empty"],
              ["+ For empty cells show: 0", I(len(SE)), "Listed as 0"]]),
          steps=["Right-click › PivotTable Options › Layout & Format", "For empty cells show: 0",
                 "Display tab › Show items with no data on rows"],
          order=["grid", "steps"], split=0.52,
          watch="Showing every item against a large table fills the PivotTable with zero rows; use it where the zeros matter."),

        T("3.6.20", "Naming and Formatting Value Fields",
          defn="Value Field Settings sets each value field's custom name and number format for the whole PivotTable.",
          simple="Rename 'Sum of total_price' to Net revenue and give it ₦#,##0; set 0.0% on shares and growth.",
          grid=dict(label="Before and after", widths=[1.8, 1.4, 1.5], rows=[
              ["Default name", "Custom name", "Format"],
              ["Sum of total_price", "Net revenue", "₦#,##0"],
              ["Count of transaction_id", "Lines", "#,##0"],
              ["Sum of total_price2", "Growth %", "0.0%"]]),
          steps=["Right-click a value › Value Field Settings", "Custom Name: type the new name",
                 "Number Format: pick the format › OK"],
          order=["grid", "steps"], split=0.55,
          watch="Formatting the cells instead of the field is lost when the PivotTable refreshes or changes shape."),

        T("3.6.21", "Retail KPIs",
          defn="A retail KPI set measures trading: sales, profit, margin, orders, basket value, units per basket and target attainment.",
          simple="Totals and counts come from the PivotTable; each ratio is one total divided by another in a cell beside it.",
          grid=dict(label="SuperMart Stores, FY2025", widths=[2, 1.6, 2.4], size=11.5, rows=[
              ["KPI", "Result", "How in Excel"],
              ["Sales", N(FY25), "Sum of Sales"],
              ["Margin %", P(PROFIT25 / FY25, 2), "Sum of Profit ÷ Sum of Sales"],
              ["Orders / basket value", f"{I(ORDERS25)} / {N(FY25 / ORDERS25)}", "Distinct Count of Order_ID; Sales ÷ Orders"],
              ["Units per basket", f"{S25.Quantity.sum() / ORDERS25:.1f}", "Sum of Quantity ÷ Orders"],
              ["Target attainment", P(FY25 / TARGET25), "Sales ÷ total annual target"]]),
          watch="Basket value uses distinct orders, not lines: one order has several lines."),

        T("3.6.22", "Hospital KPIs",
          defn="A hospital KPI set measures flow and quality: admissions, average and median length of stay, 30-day "
               "readmission rate, bed occupancy and cost per admission.",
          simple="Report the median with the average: a few very long stays pull the average up.",
          grid=dict(label=f"Mercyland, {LC.Admit_Date.min():%b %Y} to {LC.Admit_Date.max():%b %Y}", widths=[2.2, 1.7, 2.2],
                    size=11.5, rows=[
              ["KPI", "Result", "How in Excel"],
              ["Admissions", I(len(LC)), "Count of rows"],
              ["ALOS / median LOS", f"{LC.LOS_Days.mean():.1f} / {LC.LOS_Days.median():.0f} days", "Average; MEDIAN on the column"],
              ["30-day readmission rate", P(READMIT), "Readmitted Yes ÷ admissions"],
              ["Average occupancy", P(CEN.Occupancy_Pct.mean()), "Average of the daily census"],
              ["Cost per admission", N(LC.Bill_Amount_NGN.mean()), "Average of the bill"]]),
          watch="Occupancy comes from the daily bed census, not from admissions; the grains differ, so keep separate tables."),

        T("3.6.23", "University KPIs",
          defn="A university KPI set measures progression: enrolment, average exam score, grade A rate, attendance and the "
               "number of students at risk.",
          simple="Write each definition down. At risk here means exam score below 62 or attendance below 65%: a Conditional "
                 "Column flag in Power Query.",
          grid=dict(label="Kingsway University, 2024/25", widths=[2.2, 1.6], rows=[
              ["KPI", "Result"],
              ["Enrolment", I(len(CV))],
              ["Average exam score", f"{CV.Final_Score.mean():.1f}"],
              ["Grade A rate (70+)", P((CV.Final_Score >= 70).mean())],
              ["Average attendance", P(CV.Lecture_Attendance_Pct.mean() / 100)],
              ["At-risk students", f"{I(AR)} ({P(AR / len(CV))})"]]),
          watch=f"Definitions change results: with 'below 60' instead of 'below 62', {I(AR)} at risk becomes {I(AR2)}."),

        T("3.6.24", "Checking a PivotTable",
          defn="Checking a PivotTable means comparing its totals with a simple formula on the source and testing one field "
               "at a time until the wrong part is found.",
          simple="The commonest fault: the same total on every row, because the table on Rows is not related to the facts.",
          grid=dict(label="tDimDate Year on Rows, net revenue in Values", widths=[0.8, 1.8, 1.8], rows=[
              ["Year", "No relationship", "After relating order_date"]] + [
              [str(y), N(NETREV), N(YEAR[y])] for y in (2020, 2024, 2025)]),
          steps=["One field on Rows, one value in Values", "Same number on every row? Check Diagram View for the relationship",
                 "Compare the Grand Total with SUMIFS on the source", "After any source change: Data › Refresh All"],
          order=["steps", "grid"], split=0.5,
          watch="Excel also shows a yellow 'Relationships between tables may be needed' bar; read it."),

        T("3.6.25", "PivotTable and Power Query Quick Reference",
          defn="Most questions this week are answered by one of six PivotTable or Power Query settings.",
          simple="Make flags and bands in Power Query; summarise and compare in the PivotTable.",
          grid=dict(label="Need, where in Excel, result in our data", widths=[1.3, 2.6, 1.8], size=11.5, rows=[
              ["Need", "Where in Excel", "Result"],
              ["Total", "Values › Sum of total_price", N(GROSS)],
              ["Net total", "sale_type flag (Conditional Column) in Filters", N(NETREV)],
              ["Year to date", "Show Values As › Running Total In", N(YTD25[12]) + " (Dec 2025)"],
              ["Growth", "Show Values As › % Difference From (previous)", P(YOY25) + " (2025)"],
              ["Share", "Show Values As › % of Grand Total", P(CAT.iloc[0] / NETREV) + " Electronics"],
              ["Unique customers", "Value Field Settings › Distinct Count", I(BUYERS)]]),
          watch="When a number looks wrong, ask first: which filters is this cell under?"),
    ],
    takeaways=["Years and months come from the related date table.",
               "Running Total In gives year to date; % Difference From (previous) gives growth.",
               "% of Column Total gives shares; PivotTable Options shows empty cells as 0.",
               "Name and format value fields in Value Field Settings.",
               "Check every PivotTable total against a simple formula on the source."],
)

MODULES = [M36B]
