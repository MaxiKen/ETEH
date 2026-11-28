"""Week 3 edition-3 slide copy: Module 3.6A (Power Pivot and DAX: foundations and context).
DAX results are computed in pandas as exact twins of each measure on the SuperMart Online tables (w3_fx)."""
from v3kit import T
from w3_fx import (AOV, CAT, CAT_N, CODE, COMM, CU, DV, GROSS, I, LN, N, N_LINES, N_NET, NETREV, NO_DELIV, NT, P, PR, SE,
                   TX, BUYERS, NET_BUYERS)
import w3_data as D

DD = D.dimdate()
BIG = TX.loc[TX.total_price.idxmax()]
NET_BIG = NT[NT.total_price > 500000]
SHARE = CAT / NETREV
SW = NT[NT.customer_zone == "South-West"].groupby("membership_tier").total_price.sum().sort_values(ascending=False)
KCAT = LN.groupby("seller_id").category.nunique().value_counts().sort_index()
REL = LN.groupby("category").size().sort_values(ascending=False)
OC = TX.customer_id.value_counts()
AOV_CAT = NT.groupby("category").total_price.mean().sort_values(ascending=False)
NO_SALE = sorted(set(SE.seller_id) - set(TX.seller_id))
HEADS = LN.groupby("category").customer_id.nunique().sum()


def band(x):
    return "Large" if x >= 200000 else "Medium" if x >= 50000 else "Small"


M36A = dict(
    id="3.6A", title="Power Pivot and DAX: Foundations and Context",
    subtitle="Measures and calculated columns, row and filter context, CALCULATE, FILTER, ALL, iterators and safe division.",
    objectives=["Explain Power Pivot and load tables into the Data Model.",
                "Tell a calculated column from a measure, and write DAX syntax.",
                "Explain row context and filter context.",
                "Change filters with CALCULATE, FILTER, ALL and ALLEXCEPT.",
                "Use RELATED, iterators, COUNTROWS, DISTINCTCOUNT, DIVIDE, SWITCH and VAR."],
    topics=[
        T("3.6.1", "What Is Power Pivot?",
          defn="Power Pivot is Excel's data modelling engine: it holds related tables in compressed memory and evaluates "
               "DAX formulas over them.",
          simple="It goes far beyond the sheet's 1,048,576-row limit and adds DAX measures that work across tables.",
          like="Like moving from a filing tray on your desk to a records room with an index.",
          grid=dict(label="Our model", widths=[2.2, 1.2, 1.6], rows=[
              ["Table", "Rows", "Role"],
              ["tTransactions", I(N_LINES), "Fact"],
              ["tDeliveries", I(len(DV)), "Fact details"],
              ["tCustomers / tProducts / tSellers", f"{I(len(CU))} / {I(len(PR))} / {I(len(SE))}", "Dimensions"],
              ["tDimDate", I(len(DD)), "Calendar"]]),
          watch="Power Pivot does not change the source data: it reads it on refresh."),

        T("3.6.2", "Enabling the Power Pivot Add-in",
          defn="The Power Pivot add-in adds the Power Pivot ribbon tab, with Manage, Diagram View, calculated columns and "
               "Mark as Date Table.",
          simple="The Data Model is always present; only the tab is hidden by default.",
          steps=["File › Options › Add-ins", "Manage: COM Add-ins › Go", "Tick Microsoft Power Pivot for Excel › OK",
                 "The Power Pivot tab appears on the ribbon"],
          grid=dict(label="Do you need the tab?", widths=[2.6, 1], rows=[
              ["Task", "Tab needed?"],
              ["Add tables via Power Query or PivotTable", "No"],
              ["Diagram View, drawing relationships", "Yes"],
              ["Calculated columns, Mark as Date Table", "Yes"]]),
          order=["steps", "grid"], split=0.45,
          watch="Some Microsoft 365 plans and Excel for Mac lack Power Pivot; check before the class depends on it."),

        T("3.6.3", "Adding Tables to the Data Model",
          defn="Tables enter the Data Model through Power Query (Close & Load To › Add this data to the Data Model), Power "
               "Pivot › Add to Data Model, or the PivotTable dialog.",
          simple="Power Query is the best route: the cleaning steps are kept and re-run on every refresh.",
          grid=dict(label="Three routes", widths=[2.4, 1.3, 1.3], rows=[
              ["Route", "Cleaning kept?", "Use for"],
              ["Power Query › Data Model", "Yes", "Everything"],
              ["Power Pivot › Add to Data Model", "No steps", "Quick tests"],
              ["PivotTable › Add to Data Model", "No steps", "Quick tests"]]),
          watch="Loading the same table twice (once to a sheet, once to the model) doubles memory and confuses users."),

        T("3.6.4", "Calculated Columns versus Measures",
          defn="A calculated column is evaluated row by row and stored in the table; a measure is evaluated on demand for the "
               "current filter context and stores nothing.",
          simple="Use columns to slice or filter; use measures for totals and ratios.",
          like="A column is a printed label on every item; a measure is the calculator you use at the till.",
          grid=dict(label="Example on tTransactions", widths=[1.4, 2, 2], rows=[
              ["Aspect", "Calculated column", "Measure"],
              ["Example", "Is Net (Yes/No)", "[Net Revenue]"],
              ["Stored", f"{I(N_LINES)} values", "Nothing; calculated live"],
              ["Use in PivotTable", "Rows or slicer", "Values"],
              ["Result", f"{I(N_NET)} Yes, {I(N_LINES - N_NET)} No", N(NETREV)]]),
          watch="A calculated column for a total stores the same number on every row and ignores slicers; write a measure."),

        T("3.6.5", "DAX Syntax Basics",
          defn="DAX (Data Analysis Expressions) is a functional formula language whose formulas refer to tables, columns "
               "and measures rather than cells.",
          simple="Columns are written Table[column]; measures are named with := and called in brackets, [Measure].",
          **CODE("""
Gross Bookings :=
    SUM ( tTransactions[total_price] )

Orders :=
    COUNTROWS ( tTransactions )""", label="DAX MEASURES"),
          grid=dict(label="Result in a PivotTable with no filters", widths=[1.4, 1.4], rows=[
              ["Measure", "Value"], ["[Gross Bookings]", N(GROSS)], ["[Orders]", I(N_LINES)]]),
          split=0.52,
          watch="Always qualify columns with their table, tTransactions[total_price], so readers know where they come from."),

        T("3.6.6", "Row Context",
          defn="Row context is the current row, which exists while a calculated column is computed and inside iterator "
               "functions such as SUMX.",
          simple="A plain measure has no current row, so a bare column reference there cannot be resolved.",
          like="Row context is a teacher marking one script at a time; a measure looks at the whole pile.",
          grid=dict(label="The same expression in three places", widths=[1.6, 2.2, 1.8], size=11.5, rows=[
              ["Where", "Formula", "Result"],
              ["Calculated column", "= [quantity] * [unit_price]", f"{N(BIG.total_price)} on {BIG.transaction_id}"],
              ["Measure, no aggregation", "= tTransactions[quantity]", "Error: no single value"],
              ["Measure with iterator", "= SUMX ( tTransactions, [quantity] * [unit_price] )", N(GROSS)]]),
          watch="'A single value for column cannot be determined' means you referred to a column where there is no row."),

        T("3.6.7", "Filter Context",
          defn="Filter context is the set of filters in force when a measure is evaluated: PivotTable rows and columns, "
               "slicers and report filters.",
          simple="The same measure gives a different answer in each cell, because each cell has different filters.",
          grid=dict(label="[Net Revenue] by products[category]", widths=[1.8, 1.1, 1.6], rows=[
              ["Category (filter)", "Net lines", "[Net Revenue]"]] + [
              [c, I(CAT_N[c]), N(CAT[c])] for c in CAT.index[:3]] + [
              ["Beauty (lowest)", I(CAT_N["Beauty"]), N(CAT["Beauty"])], ["Grand total (no filter)", I(N_NET), N(NETREV)]]),
          watch="The grand total is not the sum of the visible rows when a measure is a ratio or a distinct count."),

        T("3.6.8", "CALCULATE",
          defn="CALCULATE evaluates an expression in a filter context modified by its filter arguments.",
          simple="It is the most important DAX function. Each filter argument replaces any existing filter on that column.",
          **CODE("""
Net Revenue :=
CALCULATE (
    [Gross Bookings],
    tTransactions[order_status]
        IN { "Delivered", "Completed" }
)""", label="DAX MEASURE"),
          grid=dict(label="Result", widths=[1.6, 0.9, 1.5], rows=[
              ["Measure", "Lines", "Value"],
              ["[Gross Bookings]", I(N_LINES), N(GROSS)],
              ["[Net Revenue]", I(N_NET), N(NETREV)],
              ["Cancelled + Returned", I(N_LINES - N_NET), N(GROSS - NETREV)]]),
          split=0.48,
          watch="A slicer on order_status is overridden inside this measure: CALCULATE replaces that filter."),

        T("3.6.9", "FILTER",
          defn="FILTER returns the rows of a table for which a condition is true, testing the rows one at a time.",
          simple="Use it inside CALCULATE when a simple column filter is not enough, such as a test on a measure or on two columns.",
          **CODE("""
Net Rev Big Orders :=
CALCULATE (
    [Net Revenue],
    FILTER ( tTransactions,
        tTransactions[total_price] > 500000 )
)""", label="DAX MEASURE"),
          grid=dict(label="Result", widths=[1.8, 1, 1.5], rows=[
              ["Measure", "Lines", "Value"],
              ["[Net Rev Big Orders]", I(len(NET_BIG)), N(NET_BIG.total_price.sum())],
              ["[Net Revenue]", I(N_NET), N(NETREV)]]),
          split=0.52,
          watch="FILTER over a whole fact table is slow on big models; for a simple column test let CALCULATE filter the column."),

        T("3.6.10", "ALL",
          defn="ALL returns all rows of a table, or all values of a column, ignoring any filters on them.",
          simple="Divide a measure by the same measure with ALL to get each item's share of the whole.",
          **CODE("""
Cat Share :=
DIVIDE (
    [Net Revenue],
    CALCULATE ( [Net Revenue],
                ALL ( tProducts[category] ) )
)""", label="DAX MEASURE"),
          grid=dict(label="Result by category", widths=[1.4, 1.6, 0.9], rows=[
              ["Category", "[Net Revenue]", "[Cat Share]"]] + [
              [c, N(CAT[c]), P(SHARE[c])] for c in CAT.index[:3]] + [["Beauty", N(CAT["Beauty"]), P(SHARE["Beauty"])]]),
          split=0.5,
          watch="The shares must add up to 100%. If they don't, the ALL is on the wrong column."),

        T("3.6.11", "ALLEXCEPT",
          defn="ALLEXCEPT removes every filter on a table except those on the columns you list.",
          simple="Use it for 'share of the parent': each tier's share of its own zone, not of the whole company.",
          **CODE("""
Share of Zone :=
DIVIDE (
    [Net Revenue],
    CALCULATE ( [Net Revenue],
        ALLEXCEPT ( tCustomers,
                    tCustomers[zone] ) )
)""", label="DAX MEASURE"),
          grid=dict(label="South-West tiers", widths=[1.1, 1.6, 1.1], rows=[
              ["Tier", "[Net Revenue]", "[Share of Zone]"]] + [
              [t, N(x), P(x / SW.sum())] for t, x in SW.items()]),
          split=0.5,
          watch="ALLEXCEPT keeps only the listed columns' filters on that table; a slicer on another column of tCustomers stops working."),

        T("3.6.12", "VALUES and DISTINCT",
          defn="VALUES returns the distinct values of a column visible in the current filter context, plus a blank for "
               "unmatched keys; DISTINCT returns them without that blank.",
          simple="Wrap it in COUNTROWS to count the values in view, or use it to build dynamic titles.",
          **CODE("""
Categories Sold :=
COUNTROWS (
    VALUES ( tProducts[category] )
)""", label="DAX MEASURE"),
          grid=dict(label="[Categories Sold] per seller", widths=[1.4, 1], rows=[
              ["Categories", "Sellers"]] + [[str(k), I(n)] for k, n in KCAT.items()]),
          split=0.55,
          watch=f"{I(KCAT.sum())} sellers have sales; none sells all {len(CAT)} categories (the most is {KCAT.index.max()})."),

        T("3.6.13", "RELATED",
          defn="RELATED returns a value from the 'one' side of a relationship to the current row on the 'many' side.",
          simple="It is the model's VLOOKUP: no column numbers, no ranges, and it follows the relationship.",
          **CODE("""
Category =
RELATED ( tProducts[category] )
-- calculated column in tTransactions""", label="DAX CALCULATED COLUMN"),
          grid=dict(label="Lines per category via RELATED", widths=[1.5, 1, 1], rows=[
              ["Category", "Lines", "Share"]] + [
              [c, I(n), P(n / N_LINES)] for c, n in list(REL.items())[:3]] + [["All 7", I(N_LINES), "100.0%"]]),
          split=0.5,
          watch="RELATED needs a relationship and a row context; in a measure, use it inside an iterator such as SUMX."),

        T("3.6.14", "RELATEDTABLE",
          defn="RELATEDTABLE returns the rows on the 'many' side that relate to the current row on the 'one' side.",
          simple="From each customer, count or sum that customer's transactions.",
          **CODE("""
Orders Placed =
COUNTROWS (
    RELATEDTABLE ( tTransactions )
)
-- calculated column in tCustomers""", label="DAX CALCULATED COLUMN"),
          grid=dict(label="Results on tCustomers", widths=[2, 1.6], rows=[
              ["Check", "Result"],
              ["Highest order count", f"{OC.max()} ({OC.idxmax()})"],
              ["Average per customer", f"{N_LINES / len(CU):.1f} orders"],
              ["Customers with no orders", f"{I(len(CU) - OC.size)} (blank)"]]),
          split=0.5,
          watch="COUNTROWS of an empty table is blank, not 0: add + 0 if the zero must show."),

        T("3.6.15", "SUMX",
          defn="SUMX evaluates an expression for each row of a table and returns the sum of the results.",
          simple="Use it when you must multiply or combine columns before adding up.",
          **CODE("""
Revenue X :=
SUMX ( tTransactions,
    tTransactions[quantity]
      * tTransactions[unit_price] )""", label="DAX MEASURE"),
          grid=dict(label="Three totals", widths=[2, 1.5, 1.2], rows=[
              ["Measure", "Result", "Meaning"],
              ["SUMX ( qty × unit_price )", N(GROSS), "Correct revenue"],
              ["SUM ( total_price )", N(float(TX.total_price.sum())), "Same: a check"],
              ["SUM ( unit_price )", N(float(TX.unit_price.sum())), "Meaningless"]]),
          split=0.45,
          watch="SUM(quantity) × SUM(unit_price) is not revenue: multiply per row first, then add."),

        T("3.6.16", "AVERAGEX",
          defn="AVERAGEX evaluates an expression for each row of a table and returns the average of the results.",
          simple="The table you iterate sets the grain: average per order line and average per customer are different questions.",
          **CODE("""
Avg Net Spend per Customer :=
AVERAGEX (
    VALUES ( tTransactions[customer_id] ),
    [Net Revenue]
)""", label="DAX MEASURE"),
          grid=dict(label="Three averages, three grains", widths=[2, 1.4, 1.1], size=11.5, rows=[
              ["Measure", "Grain", "Result"],
              ["Average quantity per line", "Order line", f"{TX.quantity.mean():.2f}"],
              ["Average net order value", f"Net line ({I(N_NET)})", N(AOV)],
              ["Avg net spend per customer", f"Customer ({I(NET_BUYERS)})", N(NETREV / NET_BUYERS)]]),
          split=0.5,
          watch="State the grain in the measure name: 'average' alone is ambiguous."),

        T("3.6.17", "COUNTROWS",
          defn="COUNTROWS returns the number of rows in a table in the current filter context; COUNT counts only the "
               "non-blank values of one column.",
          simple="COUNTROWS is the clearest way to count events such as orders or admissions.",
          grid=dict(label="Two counts on tTransactions", widths=[2.4, 1, 1.8], rows=[
              ["Measure", "Result", "Why"],
              ["COUNTROWS ( tTransactions )", I(N_LINES), "Every order line"],
              ["COUNT ( tTransactions[delivery_id] )", I(N_LINES - NO_DELIV), f"Skips {I(NO_DELIV)} blanks"]]),
          watch="COUNT on a column with blanks quietly gives a smaller number; count rows unless blanks must be excluded."),

        T("3.6.18", "DISTINCTCOUNT",
          defn="DISTINCTCOUNT returns the number of different values in a column in the current filter context.",
          simple="How many different customers bought? Distinct counts do not add up across groups.",
          grid=dict(label="Counting customers", widths=[2.4, 1, 1.8], rows=[
              ["Measure", "Value", "Meaning"],
              ["Orders (COUNTROWS)", I(N_LINES), "Order lines"],
              ["Customers (DISTINCTCOUNT)", I(BUYERS), "Unique buyers"],
              ["Sum of 7 category headcounts", I(HEADS), "Wrong: people counted repeatedly"]]),
          watch="Summing distinct counts by category counts a customer once for every category they bought in."),

        T("3.6.19", "DIVIDE",
          defn="DIVIDE(numerator, denominator, [alternate]) divides safely, returning blank, or the alternate result, when "
               "the denominator is 0 or blank.",
          simple="Use DIVIDE for every ratio instead of the / operator.",
          **CODE("""
AOV :=
DIVIDE (
    [Net Revenue],
    [Net Orders]
)""", label="DAX MEASURE"),
          grid=dict(label="[AOV] by seller", widths=[1.8, 1, 1.6], rows=[
              ["Seller", "Net lines", "[AOV]"],
              ["All sellers", I(N_NET), N(AOV)],
              [f"{NO_SALE[0]} (no sales)", "0", "Blank, not an error"]]),
          split=0.45,
          watch="With the / operator a zero denominator shows Infinity or NaN in the PivotTable; DIVIDE returns blank."),

        T("3.6.20", "IF and SWITCH in DAX",
          defn="IF returns one of two values; SWITCH(TRUE(), test1, result1, …, default) returns the result of the first true test.",
          simple="SWITCH ( TRUE (), … ) is the tidy way to write three or more bands, like IFS in Excel.",
          **CODE("""
Order Size :=
SWITCH ( TRUE (),
    [AOV] >= 200000, "Large",
    [AOV] >= 50000,  "Medium",
    "Small" )""", label="DAX MEASURE"),
          grid=dict(label="By category", widths=[1.3, 1.2, 0.9], rows=[
              ["Category", "[AOV]", "Size"]] + [
              [c, N(x), band(x)] for c, x in list(AOV_CAT.items())[:4]] + [["All", N(AOV), band(AOV)]]),
          split=0.5,
          watch="Tests run in order: put the largest threshold first, or every value lands in the first band."),

        T("3.6.21", "Variables in DAX with VAR",
          defn="VAR name = expression … RETURN result stores an intermediate result once, so it can be reused in the formula.",
          simple="Measures become easier to read and faster, because the expression is not calculated twice.",
          **CODE("""
Comm % :=
VAR rev  = [Gross Bookings]
VAR comm = SUM ( tTransactions[commission_amount] )
RETURN
    DIVIDE ( comm, rev )""", label="DAX MEASURE"),
          grid=dict(label="Values", widths=[1.4, 1.6], rows=[
              ["Part", "Value"], ["rev", N(GROSS)], ["comm", N(COMM, 2)], ["[Comm %]", P(COMM / GROSS)]]),
          split=0.55,
          watch="A variable is calculated once, where it is defined; CALCULATE later in the formula cannot change it."),
    ],
    takeaways=["Columns are stored per row; measures calculate per cell.",
               "Row context = current row; filter context = what the cell is filtered by.",
               "CALCULATE changes filters; ALL and ALLEXCEPT remove them.",
               "Iterators work row by row; the table sets the grain.",
               "DIVIDE every ratio; SWITCH for bands; VAR for clarity."],
)

MODULES = [M36A]
