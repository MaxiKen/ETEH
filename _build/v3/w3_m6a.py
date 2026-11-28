"""Week 3 edition-3.3 slide copy: Module 3.6A (Power Pivot and Data Model PivotTables, no DAX).
The course analyses in Excel with PivotTables and Power Query (M) only: helper columns and flags are made with Add Column in
Power Query, totals, counts, distinct counts and shares with PivotTable Value Field Settings, and ratios in a cell beside the
PivotTable. Every result is computed in pandas on the SuperMart Online tables (w3_fx), exactly as the PivotTable shows it."""
from v3kit import T
from w3_fx import (AOV, BUYERS, CAT, CAT_N, CU, DV, GROSS, I, LN, N, N_LINES, N_NET, NETREV, NT, P, PR, SE, TX)
import w3_data as D

DD = D.dimdate()
BIG = TX.loc[TX.total_price.idxmax()]
GCAT = LN.groupby("category").total_price.sum().sort_values(ascending=False)          # all lines, by category
ELEC_Z = LN[LN.category == "Electronics"].groupby("customer_zone").total_price.sum().sort_values(ascending=False)
SHARE = CAT / NETREV
HEADS = LN.groupby("category").customer_id.nunique().sum()
AOV_CAT = NT.groupby("category").total_price.mean()
EX36 = "Week_3 › Ex3.4_Files"


def band(x):
    return "Large" if x >= 200000 else "Medium" if x >= 50000 else "Small"


BAND = NT.assign(b=NT.total_price.map(band)).groupby("b").total_price.agg(["size", "sum"]).reindex(["Large", "Medium", "Small"])


def M(text, size=12, note=""):
    return dict(code=text.strip("\n"), code_label="M  ·  WHAT POWER QUERY WRITES IN THE FORMULA BAR", code_size=size,
                code_note=note)


M36A = dict(
    id="3.6A", title="Power Pivot and Data Model PivotTables",
    subtitle="Put the related tables in the Data Model, add helper columns in Power Query, and answer questions with "
             "PivotTables: sums, counts, distinct counts, shares and simple ratios. No new formula language.",
    objectives=["Explain Power Pivot and load tables into the Data Model.",
                "Build a PivotTable from the Data Model using fields from several related tables.",
                "Add helper columns and flags with Power Query's Custom Column and Conditional Column.",
                "Summarise with Sum, Count, Average and Distinct Count, and show values as % of the total.",
                "Work out ratios such as average order value in a cell beside the PivotTable."],
    topics=[
        T("3.6.1", "What Is Power Pivot?",
          defn="Power Pivot is Excel's data modelling engine: it holds related tables in compressed memory so that one "
               "PivotTable can use fields from all of them.",
          simple="It goes far beyond the sheet's 1,048,576-row limit. In this course you use it through Power Query and "
                 "PivotTables, so there is no new formula language to learn.",
          like="Like moving from a filing tray on your desk to a records room with an index.",
          grid=dict(label="Our model", widths=[2.2, 1.2, 1.6], rows=[
              ["Table", "Rows", "Role"],
              ["tTransactions", I(N_LINES), "Fact"],
              ["tDeliveries", I(len(DV)), "Fact details"],
              ["tCustomers / tProducts / tSellers", f"{I(len(CU))} / {I(len(PR))} / {I(len(SE))}", "Dimensions"],
              ["tDimDate", I(len(DD)), "Calendar"]]),
          watch="Power Pivot does not change the source data: it reads it again on every refresh."),

        T("3.6.2", "Enabling the Power Pivot Add-in",
          defn="The Power Pivot add-in adds the Power Pivot ribbon tab, with Manage, Diagram View and Mark as Date Table.",
          simple="The Data Model is always present; only the tab is hidden by default.",
          steps=["File › Options › Add-ins", "Manage: COM Add-ins › Go", "Tick Microsoft Power Pivot for Excel › OK",
                 "The Power Pivot tab appears on the ribbon"],
          grid=dict(label="Do you need the tab?", widths=[2.6, 1], rows=[
              ["Task", "Tab needed?"],
              ["Add tables via Power Query or PivotTable", "No"],
              ["Diagram View, drawing relationships", "Yes"],
              ["Mark as Date Table", "Yes"]]),
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

        T("3.6.4", "A PivotTable from the Data Model",
          defn="Insert › PivotTable › From Data Model builds a PivotTable whose field list shows every table in the model.",
          simple="Drag a field from one table to Rows and a number from another table to Values; the relationship joins them.",
          steps=["Insert › PivotTable › From Data Model › New Worksheet",
                 "In the field list click All to see every table",
                 "tProducts › category to Rows",
                 "tTransactions › total_price to Values"],
          grid=dict(label="Sum of total_price by category (all lines)", widths=[1.6, 1.8], rows=[
              ["category", "Sum of total_price"]] + [[c, N(x)] for c, x in list(GCAT.items())[:3]] + [
              ["Grand Total", N(GROSS)]]),
          order=["steps", "grid"], split=0.5,
          watch="If every row shows the same total, the two tables are not related: build the relationship first (3.5)."),

        T("3.6.5", "Fields from Several Related Tables",
          defn="A Data Model PivotTable can put fields from different tables in Rows, Columns, Filters and Values, as long as "
               "each table is related to the fact table.",
          simple="Here zone comes from tCustomers, category from tProducts and total_price from tTransactions: no VLOOKUP "
                 "columns were added.",
          grid=dict(label="Filter category = Electronics; Rows: tCustomers zone", widths=[1.6, 1.8], rows=[
              ["zone", "Sum of total_price"]] + [[z, N(x)] for z, x in list(ELEC_Z.items())[:4]] + [
              ["Grand Total", N(GCAT["Electronics"])]]),
          steps=["Filters: tProducts › category = Electronics",
                 "Rows: tCustomers › zone",
                 "Values: tTransactions › total_price",
                 "Sort largest to smallest"],
          order=["steps", "grid"], split=0.45,
          watch="A yellow 'Relationships between tables may be needed' bar means a table you used is not related yet."),

        T("3.6.6", "Helper Columns with Custom Column",
          defn="Add Column › Custom Column in Power Query adds a new column calculated on every row from other columns.",
          simple="Type the calculation in the dialog using column names in square brackets. Power Query writes the M step for you.",
          like="Like adding a 'line total' column to a paper invoice before you add up the page.",
          **M("""
= Table.AddColumn(transactions,
    "line_value",
    each [quantity] * [unit_price])"""),
          grid=dict(label=f"Check on {BIG.transaction_id}", widths=[1.6, 1.6], rows=[
              ["Column", "Value"],
              ["quantity", I(BIG.quantity)],
              ["unit_price", N(BIG.unit_price)],
              ["line_value (new)", N(BIG.quantity * BIG.unit_price)],
              ["total_price (given)", N(BIG.total_price)]]),
          split=0.58,
          watch="Check a new column against one you trust: here line_value equals total_price on every one of the "
                f"{I(N_LINES)} rows."),

        T("3.6.7", "A Net / Not Net Flag with Conditional Column",
          defn="Add Column › Conditional Column builds an if-then-else column by filling in a form, with no formula typing.",
          simple="Our rule: an order line counts as net revenue when order_status is Delivered or Completed.",
          steps=["Add Column › Conditional Column; New column name: sale_type",
                 "If order_status equals Delivered › Output: Net",
                 "Add Clause: order_status equals Completed › Output: Net",
                 "Else: Not net › OK, then Close & Load"],
          grid=dict(label="sale_type on tTransactions", widths=[1.2, 1, 1.7], rows=[
              ["sale_type", "Lines", "Sum of total_price"],
              ["Net", I(N_NET), N(NETREV)],
              ["Not net", I(N_LINES - N_NET), N(GROSS - NETREV)],
              ["Total", I(N_LINES), N(GROSS)]]),
          order=["steps", "grid"], split=0.52,
          watch="Write the rule once in Power Query; every PivotTable then filters on the same flag and agrees."),

        T("3.6.8", "Filtering the PivotTable to Net Sales",
          defn="Putting the flag in Filters (or on a slicer) restricts every value in the PivotTable to the rows that pass.",
          simple="Filter sale_type = Net and Sum of total_price becomes net revenue.",
          grid=dict(label="Filter sale_type = Net", widths=[1.5, 1, 1.7], rows=[
              ["category", "Lines", "Sum of total_price"]] + [
              [c, I(CAT_N[c]), N(CAT[c])] for c in CAT.index[:3]] + [
              ["Beauty (lowest)", I(CAT_N["Beauty"]), N(CAT["Beauty"])], ["Grand Total", I(N_NET), N(NETREV)]]),
          steps=["Drag sale_type to Filters", "Choose Net › OK", "Or: PivotTable Analyze › Insert Slicer › sale_type"],
          order=["grid", "steps"], split=0.58,
          watch=f"Check the filter is on: the Grand Total falls from {N(GROSS)} (3.3.10) to {N(NETREV)} once cancelled and "
                "returned lines are left out."),

        T("3.6.9", "Summarise Values By: Sum, Count and Average",
          defn="Value Field Settings › Summarize Values By changes how a field in Values is summarised: Sum, Count, Average, "
               "Max or Min.",
          simple="Drag total_price into Values several times and set a different summary on each copy.",
          grid=dict(label="Net lines (sale_type = Net), total_price", widths=[1.8, 1.8], rows=[
              ["Summarize by", "Result"],
              ["Sum", N(NETREV)],
              ["Count", I(N_NET)],
              ["Average", N(AOV)],
              ["Max / Min", f"{N(NT.total_price.max())} / {N(NT.total_price.min())}"]]),
          steps=["Right-click a value › Value Field Settings", "Summarize Values By: pick the summary",
                 "Number Format › Currency ₦ or Number with 1000 separator"],
          order=["grid", "steps"], split=0.5,
          watch="Count counts lines (rows), not orders or customers; use Distinct Count for those (3.6.10)."),

        T("3.6.10", "Distinct Count",
          defn="Distinct Count gives the number of different values in a field; it is offered only in PivotTables built "
               "on the Data Model.",
          simple="How many different customers bought? Count gives order lines; Distinct Count gives people.",
          grid=dict(label="Counting customers (all lines)", widths=[2.4, 1, 1.8], rows=[
              ["Value field", "Result", "Meaning"],
              ["Count of customer_id", I(N_LINES), "Order lines"],
              ["Distinct Count of customer_id", I(BUYERS), "Unique buyers"],
              ["Adding the 7 category distinct counts", I(HEADS), "Wrong: people counted repeatedly"]]),
          steps=["Values: customer_id", "Value Field Settings › Distinct Count",
                 "Not in the list? The PivotTable is not on the Data Model"],
          order=["grid", "steps"], split=0.55,
          watch="Distinct counts do not add up across rows: a customer who bought in three categories counts once in the total."),

        T("3.6.11", "Show Values As: % of Grand Total",
          defn="Show Values As re-expresses a value field, for example as a percentage of the grand total, without any formula.",
          simple="Put total_price in Values twice; show the second copy as % of Grand Total to get each category's share.",
          grid=dict(label="Net revenue share by category", widths=[1.4, 1.6, 0.9], rows=[
              ["category", "Net revenue", "Share"]] + [
              [c, N(CAT[c]), P(SHARE[c])] for c in CAT.index[:3]] + [["Beauty", N(CAT["Beauty"]), P(SHARE["Beauty"])],
                                                                   ["Grand Total", N(NETREV), "100.0%"]]),
          steps=["Drag total_price to Values a second time", "Right-click it › Show Values As › % of Grand Total",
                 "Rename it Share (3.6.20)"],
          order=["grid", "steps"], split=0.55,
          watch="The shares must add up to 100%. If they do not, check the filter: a share of filtered data is still 100% of what is shown."),

        T("3.6.12", "Ratios Beside the PivotTable",
          defn="A ratio KPI, such as average order value, is one total divided by another; work it out in a cell next to the "
               "PivotTable that points at the two totals.",
          simple="Type = and click the net revenue cell, type /, click the net lines cell. Excel writes GETPIVOTDATA, which "
                 "keeps working when the PivotTable moves.",
          grid=dict(label="Average order value (net lines)", widths=[1.8, 1.8], rows=[
              ["Item", "Value"],
              ["Net revenue (Sum)", N(NETREV)],
              ["Net lines (Count)", I(N_NET)],
              ["AOV = revenue ÷ lines", N(AOV)],
              ["Average of the 7 category averages", N(AOV_CAT.mean())]]),
          watch="Never average the averages: the categories have very different line counts, so the last row is wrong. "
                "Divide the two totals."),

        T("3.6.13", "Bands with a Conditional Column",
          defn="A band column puts each row into a range, such as Small, Medium or Large, so a PivotTable can count and sum by band.",
          simple="Conditional Column tests run from the top: put the largest threshold first.",
          steps=["Add Column › Conditional Column; name: size_band",
                 "If total_price is greater than or equal to 200000 › Large",
                 "Else if total_price is greater than or equal to 50000 › Medium",
                 "Else: Small; then Rows: size_band in the PivotTable"],
          grid=dict(label="Net lines by size_band", widths=[1.1, 1, 1.7], rows=[
              ["size_band", "Lines", "Sum of total_price"]] + [
              [b_, I(r["size"]), N(r["sum"])] for b_, r in BAND.iterrows()] + [["Total", I(N_NET), N(NETREV)]]),
          order=["steps", "grid"], split=0.52,
          watch="Put the smallest test first and every line over ₦50,000 lands in the first band you wrote."),

        T("3.6.14", "Group By in Power Query",
          defn="Transform › Group By turns many rows into one row per group, with sums or counts, as a new fixed table.",
          simple="Use it when you need a summary table to load, merge or chart; use a PivotTable when the reader should re-slice.",
          **M("""
= Table.Group(net_lines, {"category"}, {
    {"net_revenue", each List.Sum([total_price])},
    {"lines", each Table.RowCount(_)}})"""),
          grid=dict(label="Group By category (net lines)", widths=[1.4, 1.6, 0.8], rows=[
              ["category", "net_revenue", "lines"]] + [
              [c, N(CAT[c]), I(CAT_N[c])] for c in CAT.index[:4]] + [["(7 rows in all)", N(NETREV), I(N_NET)]]),
          split=0.55,
          watch="A Group By table does not follow slicers. It updates on Refresh, not when someone clicks a filter."),
    ],
    takeaways=["The Data Model lets one PivotTable use fields from several related tables.",
               "Helper columns, flags and bands are made once in Power Query (Custom and Conditional Column).",
               "Sum, Count, Average and Distinct Count come from Value Field Settings.",
               "Show Values As gives shares; ratios are total ÷ total in a cell beside the PivotTable.",
               "Never average averages or add distinct counts."],
)

MODULES = [M36A]
