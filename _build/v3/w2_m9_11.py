"""Week 2 edition-3 slide copy: Modules 2.9 (M language), 2.10 (PivotTables), 2.11 (PivotCharts).
Figures from w2_fx (LQ = the Mercyland cleaning query replayed; SMF = the full SuperMart table tSM; CVF = Kingsway)."""
from v3kit import T
from w2_fx import LF, LQ, CVF, SMF, CTRL, N, I, P, Nm, S
import w2_data as D

ZY = SMF["zy"]
YEARS = [2022, 2023, 2024, 2025]
ZORD = ZY.sort_values(2025, ascending=False).index.tolist()
SMM = SMF["store_month"]
IKM = SMM.loc["Lagos - Ikeja"].Sales
DEPT = D.table("tDept23")
DEPT = DEPT[DEPT.Department != "Total"] if "Total" in set(DEPT.Department) else DEPT


def m(v):
    return f"₦{v / 1e6:,.1f}m"


M29 = dict(
    id="2.9", title="M Language",
    subtitle="The language Power Query writes for you: read it, edit it in the formula bar and the Advanced Editor, and "
             "write small steps yourself.",
    objectives=["Explain what M is and where it runs.", "Read let … in, step names and the step chain.",
                "Recognise M data types and null.", "Use the formula bar and the Advanced Editor.",
                "Use Text, Number, Date, List and Table functions in real queries."],
    topics=[
        T("2.9.1", "What Is the M Language?",
          defn="M (the Power Query Formula Language) is the functional language in which every Power Query step is written.",
          simple="Each click in the editor writes one line of M. You can read it, change it, or type steps yourself.",
          grid=dict(label="A first query on Mercyland_Episodes_2023_DIRTY.csv", widths=[1.6, 3], size=12, rows=[
              ["Applied step", "The M behind it"], ["Source", "Csv.Document(File.Contents(…))"],
              ["Promoted Headers", "Table.PromoteHeaders(Source, …)"],
              ["Proper Names", "Table.TransformColumns(…, {{\"Patient_Name\", Text.Proper}})"],
              ["Result", f"'{LQ['name_example']}' → '{LQ['name_example'].title()}'"]])),

        T("2.9.2", "Where M Runs and Why It Is Case-Sensitive",
          simple="M runs inside Power Query in Excel, Power BI and Dataflows. Unlike Excel formulas, it is case-sensitive: "
                 "names, functions and text comparisons must match exactly.",
          grid=dict(label=f"After removing duplicates ({LQ['dedup']:,} rows)", widths=[2.8, 1.6], size=12, rows=[
              ["Test", "Result"], ["Ward spellings in the raw file", str(LQ["ward_raw_n"])],
              ["Rows exactly \"Obeche Ward\" (raw)", I(LQ["ruby_raw"])], ["Obeche Ward rows after the ward mapping", I(LQ["ruby_map"])],
              ["text.proper instead of Text.Proper", "Error: name not recognised"]]),
          watch="\"obeche ward\" = \"Obeche Ward\" is false in M. Lower-case both sides (Text.Lower) before comparing text."),

        T("2.9.3", "Introduction to M Syntax",
          simple="A query is let, a list of named steps separated by commas, then in and the step to return.",
          grid=dict(label="Kingsway: attendance band as a custom column", widths=[1], size=11.5, rows=[
              ["M code"],
              ["Band = Table.AddColumn(Typed, \"Attendance_Band2\", each"],
              ["    if [Lecture_Attendance_Pct] >= 90 then \"90%+\""],
              ["    else if [Lecture_Attendance_Pct] >= 75 then \"75-89%\" else \"Below 75%\", type text)"]]),
          points=[f"90%+: {CVF['bands']['90%+']:,} students", f"75-89%: {CVF['bands']['75-89%']:,}",
                  f"Below 75%: {CVF['bands']['Below 75%']:,}", f"Total: {CVF['n']:,}"],
          points_label="RESULT", order=["grid", "points"], split=0.7,
          watch="if needs both then and else, all lower case. each is shorthand for a function of the current row, _."),

        T("2.9.4", "M Data Types",
          defn="A type tells M what kind of value a column holds: type text, type number, Int64.Type, type date, type logical, "
               "Currency.Type, and so on. null means no value.",
          grid=dict(label="Types on the dirty Mercyland file", widths=[1.4, 1.8, 2.2], size=12, rows=[
              ["Column", "M type", "Result"],
              ["Bill_Amount_NGN", "type number, locale en-US", f"{LF['bill_text']} values like '{LQ['bill_example']}' parsed"],
              ["Admit_Date", "type date", "All valid"], ["Discharge_Date", "type date", f"{LF['dis_blank']} null"],
              ["Age", "Int64.Type", "All whole numbers"]]),
          watch="null is not 0 and not \"\". null + 1 is null; use List.Sum, which skips nulls, for totals."),

        T("2.9.5", "let Expressions and the Step Chain",
          simple="Each step names the step before it. That chain is what Applied Steps shows. Break a link and every later step fails.",
          grid=dict(label="The Mercyland cleaning query", widths=[2.2, 1.4, 1], size=12, rows=[
              ["Step", "Uses", "Rows after"], ["Promoted", "Source", I(LF["rows"])], ["Dedup", "Promoted", I(LQ["dedup"])],
              ["Clean (exceptions removed)", "Exception", I(LQ["clean"])], ["Names, if Dedup is deleted", "Dedup (gone)", "Error"]]),
          points=["Steps run in the order they refer to each other", "The value after in is what loads",
                  "Names with spaces need #\"…\", e.g. #\"Removed Columns\""],
          points_label="RULES", order=["grid", "points"], split=0.58),

        T("2.9.6", "The Power Query Formula Bar",
          simple="Select a step to see its M in the formula bar (View › Formula Bar). Edit it there to change the step.",
          diagram=("pqeditor", dict(highlight=("formulabar",), query="qMercylandClean",
                                    cols=("Patient_Name", "Ward", "Age"),
                                    types=("ABC", "ABC", "123"),
                                    steps=("Source", "Promoted", "Dedup", "Names", "WardMap"),
                                    formula="= Table.Distinct(Promoted)",
                                    labels={"Formula bar": "the M of the selected step",
                                            "Dedup selected": f"{LQ['dedup']:,} rows; Promoted had {LF['rows']:,}",
                                            "fx button": "inserts a new step that refers to the previous one"})),
          order=["diagram"]),

        T("2.9.7", "The Advanced Editor",
          simple="Home › Advanced Editor shows the whole query as one script. Copy it to share or document a query; paste one in "
                 "to reuse it.",
          grid=dict(label="The complete Mercyland cleaning query, pasted in", widths=[2.4, 1.1], size=12.5, rows=[
              ["Check", "Result"], ["Steps in the script", "12"], ["Duplicates removed", I(LF["dups"])],
              ["Exception rows removed", I(sum(LQ["exceptions"].values()))], ["Clean rows", I(LQ["clean"])]]),
          points=["Display Options › line numbers and word wrap", "Comments: // to end of line",
                  "A syntax error is reported with its line"],
          points_label="TIPS", order=["grid", "points"], split=0.55),

        T("2.9.8", "Basic M Functions",
          simple="M functions are grouped by family. The prefix says what they work on.",
          grid=dict(label="One example from each family", widths=[0.9, 3.3, 1.2], size=12, rows=[
              ["Family", "Example (our data)", "Result"],
              ["Text.", "Text.BeforeDelimiter(\"Benin City, Edo\", \",\")", "\"Benin City\""],
              ["Number.", "Number.Round(0.12749, 3)", "0.127"],
              ["Date.", "Date.EndOfMonth(#date(2024, 2, 10))", "29-Feb-2024"],
              ["List.", "List.Contains({\"m\", \"male\"}, \"male\")", "true"],
              ["Table.", "Table.RowCount(Dedup)", I(LQ["dedup"])]]),
          watch="M counts from 0: Text.Middle(\"SM-2025-107503\", 3, 4) returns \"2025\" (Excel's MID starts at 1)."),

        T("2.9.9", "Text, Number and Date Functions in M",
          simple="Most cleaning is Text functions applied with Table.TransformColumns.",
          fx=[["each if List.Contains({\"m\", \"male\"}, Text.Lower(_)) then \"Male\" else \"Female\"", "Gender"],
              ["each let t = Text.Lower(Text.Trim(_)) in if Text.StartsWith(t, \"emerg\") then \"Emergency\" else Text.Proper(t)",
               "Admission_Route"],
              ["Duration.Days([Discharge_Date] - [Admit_Date])", "LOS_Days"]],
          fx_head=["M expression", "Column"], fx_widths=[4.6, 1.2],
          grid=dict(label=f"After de-duplication ({LQ['dedup']:,} rows)", widths=[1.4, 1, 2.6], size=12, rows=[
              ["Column", "Before", "After"],
              ["Gender", f"{LQ['gender_raw_n']} spellings", " / ".join(f"{k} {v:,}" for k, v in LQ["gender"].items())],
              ["Admission_Route", f"{LQ['atype_raw_n']} spellings", " / ".join(f"{k} {v:,}" for k, v in LQ["atype"].items())]]),
          order=["fx", "grid"], stack=True),

        T("2.9.10", "Table Functions in M",
          simple="Table functions work on whole tables: select rows, group, join, unpivot. Most applied steps are Table functions.",
          fx=[["Table.SelectRows(t, each [Profit] < 0)", "Filter rows"],
              ["Table.Group(t, {\"Store_Name\", \"Month\"}, {{\"Sales\", each List.Sum([Sales]), type number}})", "Group By"],
              ["Table.NestedJoin(a, \"Store_ID\", b, \"Store_ID\", \"Stores\", JoinKind.LeftOuter)", "Merge Queries"],
              ["Table.UnpivotOtherColumns(t, {\"Store\"}, \"Month\", \"Sales_m\")", "Unpivot"]],
          fx_head=["M function", "Menu equivalent"], fx_widths=[4.6, 1.2],
          grid=dict(label="FY2025 sales grouped by store and month", widths=[2.6, 2], size=12, rows=[
              ["Check", "Result"], ["Rows after Table.Group", f"{len(SMM)} (12 stores × 12 months)"],
              ["Lagos - Ikeja, best month", f"{IKM.idxmax()}: {N(IKM.max())}"],
              ["Abuja - Garki, Feb 2025 profit", N(SMM.loc[("Abuja - Garki", "2025-02")].Profit)]]),
          order=["fx", "grid"], stack=True),

        T("2.9.11", "Practical M Examples Across the Three Datasets",
          simple="The same pattern on three datasets: source, clean, add logic, return one tidy table.",
          grid=dict(label="One query per dataset", widths=[1.1, 2.6, 1.8], size=12, rows=[
              ["Dataset", "Rule in M", "Result"],
              ["Mercyland", "Remove duplicates and exception rows", f"{LF['rows']:,} → {LQ['clean']:,} rows"],
              ["Kingsway", "Final_Score < 60 and Lecture_Attendance_Pct < 75", f"{CVF['mentor']} students need a mentor"],
              ["SuperMart", "Group by store and month", f"{len(SMM)} rows"]]),
          points=[f"{k}: {v}" for k, v in LQ["exceptions"].items()],
          points_label="MERCYLAND EXCEPTION ROWS", order=["grid", "points"], split=0.62),
    ],
    takeaways=["Every applied step is a line of M; you can read and edit it.",
               "M is case-sensitive: Text.Proper, not text.proper; compare text after Text.Lower.",
               "let … in: each step names the previous one; in returns the last.",
               "null is not zero; set types deliberately, with a locale when needed.",
               "Text., Number., Date., List. and Table. functions cover most cleaning."],
)

M210 = dict(
    id="2.10", title="PivotTables",
    subtitle="Summarise the 9,994 SuperMart lines (2022-2025) by any field in seconds, then group, compare, format, filter and refresh.",
    objectives=["Explain what a PivotTable does and why analysts use it.", "Prepare a source that meets the requirements.",
                "Build a pivot with the four areas.", "Choose aggregations, group dates and numbers, and show values as %.",
                "Format, sort, filter, refresh, add calculated fields, slicers and timelines."],
    topics=[
        T("2.10.1", "What Is a PivotTable?",
          defn="A PivotTable is an interactive summary of a table: it groups rows by the fields you choose and aggregates "
               "the values, without formulas.",
          like="Like a till report at closing time: thousands of receipts, summarised by department in seconds.",
          grid=dict(label=f"Sales by zone and year, from tSM ({SMF['rows']:,} lines)", widths=[1.4, 1, 1, 1.2], size=12, rows=[
              ["Zone", "2022", "2025", "Grand Total"]] + [[z, m(ZY.loc[z, 2022]), m(ZY.loc[z, 2025]), m(ZY.loc[z].sum())]
                                                      for z in ZY.sort_values(2025, ascending=False).index] + [
              ["All zones", m(ZY[2022].sum()), m(ZY[2025].sum()), m(SMF["total"])]])),

        T("2.10.2", "Why PivotTables Matter to Analysts",
          simple="The same zone × year summary built twice: once with formulas, once as a pivot.",
          grid=dict(label="Same answer, different effort", widths=[1.4, 2.2, 1.6], size=12.5, rows=[
              ["Approach", "Work needed", "Grand total"], ["SUMIFS grid", "16 SUMIFS + 9 SUM formulas", N(SMF["total"])],
              ["PivotTable", "3 drags", N(SMF["total"])]]),
          points=["Fast: rearrange by dragging", "Safe: no formulas to break", "Flexible: drill down with a double-click",
                  "Refreshable: new rows come in"],
          points_label="WHY", order=["grid", "points"], split=0.62),

        T("2.10.3", "PivotTable Data Requirements",
          simple="A pivot is only as good as its source. Check the source before you pivot.",
          grid=dict(label="Is tSM ready?", widths=[2.2, 2], size=12.5, rows=[
              ["Requirement", "tSM status"], ["One header row, every column named", f"{SMF['cols']} headers, none blank"],
              ["No blank rows or columns", "Cleaned in 2.5 and 2.6"], ["One record per row", "One sales line per row"],
              ["Consistent types per column", "Sales numeric, dates real dates"], ["Source is an Excel table", "Yes: tSM"]]),
          watch="One text value in Sales and the pivot offers Count of Sales instead of Sum (2.10.13)."),

        T("2.10.4", "The Four PivotTable Areas",
          simple="Drag fields into four areas. Rows and Columns decide the grouping; Values are aggregated; Filters limit the data.",
          diagram=("pivot", dict(filters=["Year = 2025"], columns=["Category"], rows=["Zone"], values=["Sum of Sales"],
                                 table=[["Sum of Sales", "Furniture", "Office Supplies", "Technology"]] + [
                                     [z, m(D.table("tZoneCat").set_index("Zone").loc[z, "Furniture"]),
                                      m(D.table("tZoneCat").set_index("Zone").loc[z, "Office Supplies"]),
                                      m(D.table("tZoneCat").set_index("Zone").loc[z, "Technology"])] for z in ZORD])),
          order=["diagram"]),

        T("2.10.5", "Creating a PivotTable",
          steps=["Click any cell in tSM", "Insert › PivotTable › From Table/Range", "Table/Range: tSM; New Worksheet › OK",
                 "Drag Zone to Rows and Sales to Values"],
          grid=dict(label="Result", widths=[1.6, 1.5], size=12.5, rows=[
              ["Row Labels", "Sum of Sales"]] + [[z, N(ZY.loc[z].sum())] for z in sorted(ZY.index)] + [
              ["Grand Total", N(SMF["total"])]]),
          order=["steps", "grid"], split=0.5,
          watch="Rows sort A to Z by default, not by size. Sort by value (2.10.16) to see the largest first."),

        T("2.10.6", "Choosing the Data Source: Range, Table or Data Model",
          simple="Build on a table, not a fixed range: a table grows with new rows and the pivot follows on Refresh.",
          grid=dict(label="A pivot built before the 2025 lines were added", widths=[2.2, 1, 1.2], size=12.5, rows=[
              ["Source", "Rows seen", "Total sales"], ["Fixed range (2022-2024 only)", I(SMF["fixed"][0]), m(SMF["fixed"][1])],
              ["Table tSM", I(SMF["rows"]), m(SMF["total"])]]),
          points=["Range: fixed edges; misses new rows", "Table: grows automatically",
                  "Data Model: several related tables, Distinct Count (Week 3)"],
          points_label="THREE SOURCES", order=["grid", "points"], split=0.58),

        T("2.10.7", "Placing Fields in the Rows Area",
          simple="Each field in Rows adds a level. The outer field groups the inner one.",
          grid=dict(label="FY2025: Zone, then Category, in Rows", widths=[2, 1.6], size=12.5, rows=[
              ["Row Labels", "Sum of Sales"], ["South-West", N(ZY.loc["South-West", 2025])]] + [
              [f"    {c}", N(D.table("tZoneCat").set_index("Zone").loc["South-West", c])]
              for c in ["Furniture", "Office Supplies", "Technology"]] + [["South-South", N(ZY.loc["South-South", 2025])]]),
          points=["Drag to reorder levels", "Click − or + to collapse or expand",
                  "Design › Report Layout › Tabular Form for a flat table"],
          points_label="TIPS", order=["grid", "points"], split=0.55),

        T("2.10.8", "Placing Fields in the Columns Area",
          simple="Columns spread a field across the top. Keep it short: few values read well; many make a wide, unreadable table.",
          grid=dict(label="Fields in Columns for tSM", widths=[2, 1.3, 1.5], size=12.5, rows=[
              ["Field in Columns", "Value columns", "Verdict"], ["Year", "4", "Easy to read"],
              ["Quarter within Year", "16", "Readable with care"], ["Month, all years", "48", "Too wide"]]),
          watch="Put the field with more items in Rows and the one with fewer in Columns."),

        T("2.10.9", "Adding Values and Choosing Aggregations",
          simple="Value Field Settings › Summarize Values By changes Sum to Count, Average, Max, Min and others.",
          grid=dict(label="Sales in Values four times, FY2025", widths=[2.4, 1.6], size=12.5, rows=[
              ["Summary", "Result"], ["Sum of Sales", N(CTRL["sales"])], ["Count of Sales (lines)", I(CTRL["rows"])],
              ["Average of Sales", N(CTRL["sales"] / CTRL["rows"])],
              ["Distinct Count of Order_ID", f"{SMF['orders']:,} orders"]]),
          watch="Distinct Count is available only when you tick 'Add this data to the Data Model' while creating the pivot."),

        T("2.10.10", "Using Report Filters",
          simple="A field in Filters limits the whole pivot to the items you choose, without changing its layout.",
          grid=dict(label="Store pivot, filter Year = 2025", widths=[2, 1.4, 1.4], size=12.5, rows=[
              ["Store", "Sales", "Profit"]] + [[s, N(r.Sales), N(r.Profit)] for s, r in SMF["store"].head(3).iterrows()]),
          points=["PivotTable Analyze › Options › Show Report Filter Pages: one sheet per item",
                  "Select multiple items for several years", "A slicer does the same, visibly (2.10.19)"],
          points_label="MORE", order=["grid", "points"], split=0.6),

        T("2.10.11", "Grouping Dates",
          simple="Drop a date field into Rows and Excel groups it by Years, Quarters and Months. Right-click › Group to change.",
          grid=dict(label="FY2025 sales by quarter", widths=[1.4, 1.6], size=12.5, rows=[
              ["2025 quarter", "Sales"]] + [[f"Q{q}", N(v)] for q, v in SMF["q25"].items()]),
          points=["Right-click a date › Group › Months, Quarters, Years", "Days with Number of days = 7 for weeks",
                  "Ungroup to see each date"],
          points_label="HOW", order=["grid", "points"], split=0.5,
          watch="Dates stored as text cannot be grouped. Fix them first (2.5.8, 2.6.6)."),

        T("2.10.12", "Grouping Numbers and Text",
          simple="Right-click a number field in Rows › Group to make bands. Select text items › Group to combine them.",
          grid=dict(label="FY2025 sales lines by size", widths=[1.8, 1], size=12.5, rows=[
              ["Line size", "Lines"]] + [[k.replace("-", " to "), I(v)] for k, v in SMF["sizes"].items()]),
          points=["Numbers: Starting at, Ending at, By (equal bands)", "Unequal bands: a helper column with IFS, then pivot",
                  "Text: select items › Group Selection"],
          points_label="OPTIONS", order=["grid", "points"], split=0.5),

        T("2.10.13", "The Default Summarisation Trap: Sum versus Count",
          simple="If any cell in a numeric field is text or blank, the pivot defaults to Count. The numbers look plausible, but "
                 "they are line counts.",
          grid=dict(label="Same field, two summaries, FY2025", widths=[1.5, 1.2, 1.6], size=12.5, rows=[
              ["Zone", "Count of Sales", "Sum of Sales"]] + [[z, I(r["count"]), N(r["sum"])] for z, r in SMF["zone_cnt"].iterrows()]),
          watch="In the raw dirty file 234 Sales values are text, so a pivot built on it defaults to Count of Sales. "
                "Always read the field caption before reading the numbers."),

        T("2.10.14", "Show Values As: Percentages and Differences",
          simple="Value Field Settings › Show Values As turns totals into % of total, % of row, difference from a previous item, "
                 "running total or rank.",
          grid=dict(label="Zone × Year pivot", widths=[1.5, 1.6, 1.4], size=12, rows=[
              ["Zone", "2025 as % of row total", "Difference from 2024"]] + [
              [z, P(SMF["share25"][z]), ("+" if SMF["diff25"][z] >= 0 else "") + Nm(SMF["diff25"][z])]
              for z in SMF["share25"].index]),
          points=["% of Row Total: each year's share of the zone's 4 years", "Difference From › Base field Year › (previous)",
                  "Add Sales twice to show the value and the %"],
          points_label="SETTINGS", order=["grid", "points"], split=0.6),

        T("2.10.15", "Number Formatting in PivotTables",
          simple="Format through Value Field Settings › Number Format, not the Home tab: the format then survives refresh and re-layout.",
          grid=dict(label="FY2025 store pivot for a board report", widths=[1.4, 1.2, 1.5], size=12.5, rows=[
              ["Field", "Format code", "Shows"], ["Sum of Sales", "₦#,##0", N(SMF["store"].Sales.iloc[0])],
              ["Margin", "0.0%", P(SMF["store"].Profit.iloc[0] / SMF["store"].Sales.iloc[0])], ["Lines", "#,##0", I(CTRL["rows"])],
              ["Sales in millions", "₦#,##0.0,,\"m\"", m(SMF["store"].Sales.iloc[0])]]),
          watch="Formatting cells directly is lost when the pivot refreshes or changes shape."),

        T("2.10.16", "Sorting and Filtering Within PivotTables",
          simple="Sort by a value to rank; use Value Filters › Top 10 for the top or bottom N.",
          grid=dict(label="FY2025 rankings", widths=[1.4, 2, 1], size=12.5, rows=[
              ["Ranking", "Store", "Value"], ["Top sales", SMF["store"].index[0], m(SMF["store"].Sales.iloc[0])],
              ["2nd sales", SMF["store"].index[1], m(SMF["store"].Sales.iloc[1])],
              ["Biggest loss", SMF["loss_stores"].index[0], m(SMF["loss_stores"].Profit.iloc[0])],
              ["2nd biggest loss", SMF["loss_stores"].index[1], m(SMF["loss_stores"].Profit.iloc[1])]]),
          points=["Right-click a value › Sort › Largest to Smallest", "Row filter › Value Filters › Top 10 › Bottom 2 by Profit",
                  "Label Filters for text rules"],
          points_label="HOW", order=["grid", "points"], split=0.58),

        T("2.10.17", "Refreshing PivotTables",
          simple="A pivot keeps a copy of its source (the pivot cache). It does not update until you refresh.",
          grid=dict(label="December's file arrives", widths=[2.2, 0.9, 1.6], size=12.5, rows=[
              ["State", "Lines", "Sales"], ["Jan-Nov (before refresh)", I(SMF["jan_nov"][0]), N(SMF["jan_nov"][1])],
              ["December added", I(SMF["dec"][0]), N(SMF["dec"][1])], ["After Refresh All", I(CTRL["rows"]), N(CTRL["sales"])]]),
          points=["Right-click › Refresh (Alt+F5)", "Data › Refresh All (Ctrl+Alt+F5)",
                  "PivotTable Options › Data › Refresh data when opening the file"],
          points_label="HOW", order=["grid", "points"], split=0.6,
          watch="Every pivot on the same source shares one cache: refreshing one refreshes them all."),

        T("2.10.18", "Calculated Fields and Calculated Items",
          defn="A calculated field is a new value field computed from the sums of other fields. A calculated item is a new item "
               "inside a row or column field.",
          grid=dict(label="FY2025 chain margin", widths=[3, 1.2], size=12.5, rows=[
              ["Method", "Margin"], ["Calculated field: = Profit / Sales (sum ÷ sum)", P(CTRL["profit"] / CTRL["sales"])],
              ["Average of each line's margin (wrong)", P(SMF["avg_line_margin"])]]),
          steps=["PivotTable Analyze › Fields, Items & Sets › Calculated Field", "Name: Margin; Formula: = Profit / Sales",
                 "Format as 0.0%"],
          order=["grid", "steps"], split=0.58,
          watch="Averaging line margins gives every line the same weight. A ₦444 line counts as much as a ₦14m one."),

        T("2.10.19", "Slicers and Timelines with PivotTables",
          defn="A slicer is a set of filter buttons for one field. A timeline is a slider for filtering a date field by year, "
               "quarter, month or day.",
          grid=dict(label="Live pivots with a Zone slicer and an Order_Date timeline", widths=[2.2, 1.6], size=12.5, rows=[
              ["Selection", "Sales"], ["Zone = South-West, 2025", N(ZY.loc["South-West", 2025])],
              ["All zones, Q4 2025", N(SMF["q25"][4])]]),
          steps=["PivotTable Analyze › Insert Slicer › Zone", "Insert Timeline › Order_Date",
                 "Slicer › Report Connections: tick every pivot to control"],
          order=["grid", "steps"], split=0.5),
    ],
    takeaways=["A pivot groups and aggregates a table without formulas.",
               "Build on an Excel table; check headers, blanks and types first.",
               "Rows, Columns, Values, Filters; few items in Columns.",
               "Read the caption: Count of Sales means text slipped in.",
               "Group dates, Show Values As, calculated fields for ratios; refresh after every change."],
)

M211 = dict(
    id="2.11", title="PivotCharts",
    subtitle="Charts tied to a PivotTable: they follow its fields, filters and refresh, so the picture always matches the numbers.",
    objectives=["Explain what a PivotChart is and when to use one.", "Create a PivotChart and name its parts.",
                "Understand the link with its PivotTable.", "Format the chart and choose the right chart type.",
                "Filter and sort a PivotChart."],
    topics=[
        T("2.11.1", "What Is a PivotChart?",
          defn="A PivotChart is a chart whose data comes from a PivotTable. Its axis, legend and values are the pivot's fields.",
          simple="Change the pivot and the chart changes; filter the chart and the pivot filters too.",
          grid=dict(label="Pivot 1 as a chart: sales by zone and year", widths=[1.6, 1, 1], size=12, rows=[
              ["Zone", "2022", "2025"]] + [[z, m(ZY.loc[z, 2022]), m(ZY.loc[z, 2025])] for z in ZORD])),

        T("2.11.2", "Why Use a PivotChart?",
          simple="Patterns that are hard to see in a table jump out of a chart.",
          grid=dict(label="Sales by year and quarter: is there a seasonal pattern?", widths=[0.9, 1.3, 1.6], size=12, rows=[
              ["Year", "Q4 sales", "Q4 share of the year"]] + [[str(y), m(r.Q4), P(r.share)] for y, r in SMF["q4"].iterrows()]),
          points=[f"Q4 is the biggest quarter every year, at {SMF['q4'].share.min() * 100:.0f}-{SMF['q4'].share.max() * 100:.0f}% of annual sales", "A column chart shows it at a glance",
                  "Plan stock and staff for the fourth quarter"],
          points_label="WHAT THE CHART SHOWS", order=["grid", "points"], split=0.55),

        T("2.11.3", "Difference Between Normal Charts and PivotCharts",
          simple="A normal chart is drawn from fixed cells. A PivotChart is drawn from a pivot, so it follows new data and filters.",
          grid=dict(label="FY2025 sales by month, after December is added", widths=[1.8, 1.4, 1.2], size=12.5, rows=[
              ["", "Normal chart", "PivotChart"], ["Bars shown", "11 (Jan-Nov)", "12"],
              ["Total behind the bars", m(SMF["jan_nov"][1]), m(CTRL["sales"])], ["Field buttons", "No", "Yes"],
              ["Scatter, histogram, map", "Available", "Not available"]])),

        T("2.11.4", "Creating a PivotChart",
          steps=["Click inside the pivot", "PivotTable Analyze › PivotChart", "Choose Clustered Column › OK",
                 "Or Insert › PivotChart from the table directly: pivot and chart together"],
          grid=dict(label="FY2025 sales by zone", widths=[1.6, 1.2], size=12.5, rows=[
              ["Zone (2025)", "Sales"]] + [[z, m(ZY.loc[z, 2025])] for z in ZORD]),
          order=["steps", "grid"], split=0.55),

        T("2.11.5", "Understanding the PivotChart Interface",
          simple="A PivotChart has the parts of a normal chart plus field buttons and its own ribbon tabs.",
          grid=dict(label="The zone × year chart", widths=[1.5, 3], size=12.5, rows=[
              ["Part", "What it does here"], ["PivotChart Analyze tab", "Refresh, Field List, Field Buttons"],
              ["Design and Format tabs", "Change Chart Type, styles, colours"], ["Field list", "Axis = Zone, Legend = Year"],
              ["Field buttons", "Zone and Year drop-downs on the chart"], ["Plot area", "16 bars, one colour per year"]])),

        T("2.11.6", "PivotChart Fields and the Linked PivotTable",
          simple="The chart's areas have chart names, but they are the pivot's areas.",
          grid=dict(label="Chart area ↔ pivot area", widths=[1.8, 1.2, 1.4], size=12.5, rows=[
              ["Chart area", "Pivot area", "Field used"], ["Axis (Categories)", "Rows", "Zone"], ["Legend (Series)", "Columns", "Year"],
              ["Values", "Values", "Sum of Sales"], ["Filters", "Filters", "(none)"]]),
          watch="Deleting the PivotTable turns the chart into a static chart with fixed values."),

        T("2.11.7", "Bi-Directional Sync Between PivotTable and PivotChart",
          simple="Filter on the chart and the pivot filters; filter the pivot and the chart redraws. They are one object seen twice.",
          grid=dict(label="Zone button on the chart › South-West only", widths=[1.2, 1.4], size=12.5, rows=[
              ["South-West", "Sales"]] + [[str(y), m(SMF["sw_year"][y])] for y in YEARS]),
          points=["The pivot now shows South-West only", "Clear the filter on either to restore both",
                  "A slicer connected to the pivot drives the chart too"],
          points_label="WHAT HAPPENS", order=["grid", "points"], split=0.45),

        T("2.11.8", "Basic PivotChart Formatting",
          simple="Format for the reader: a title that states the finding, a clean axis and no clutter.",
          grid=dict(label="FY2025 zone chart for a board pack", widths=[1.5, 2.5], size=12.5, rows=[
              ["Setting", "Value used"], ["Title", "South-West leads FY2025 sales"], ["Axis number format", "₦#,##0,,\"m\""],
              ["Axis minimum", "0 (never truncate a column chart)"], ["Field buttons", "Hidden for printing"]]),
          watch="An axis starting above 0 exaggerates differences. Column and bar charts start at 0."),

        T("2.11.9", "Changing Chart Types",
          simple="Design › Change Chart Type. Choose by the shape of the data, not by looks.",
          grid=dict(label="Match the chart to the data", widths=[1.9, 1.6, 2], size=12, rows=[
              ["Data shape", "Chart", "Why"], ["Zones over 4 years", "Clustered column", "Compare and see a trend"],
              [f"{len(DEPT)} departments ranked", "Bar", "Long labels, easy ranking"],
              ["5 grades of one whole", "Pie", "Few parts that add to 100%"], ["12 stores' share", "Sorted bar, not pie", "Too many slices"],
              ["Months in sequence", "Line", "Trend over time"]])),

        T("2.11.10", "Filtering and Sorting in PivotCharts",
          simple="Use the field buttons to filter and sort. Sorting the pivot sorts the chart.",
          grid=dict(label="FY2025 zone chart, sorted largest first", widths=[1.6, 1.2], size=12.5, rows=[
              ["Zone (2025)", "Sales"]] + [[z, m(ZY.loc[z, 2025])] for z in ZORD]),
          steps=["Click the Zone field button › More Sort Options", "Descending (Z to A) by Sum of Sales",
                 "Year button › tick 2025 only"],
          order=["grid", "steps"], split=0.45),

        T("2.11.11", "Practical PivotChart Examples Using the Three Datasets",
          simple="One PivotChart per dataset, each type chosen for its data.",
          grid=dict(label="Three charts", widths=[1.1, 0.9, 3.4], size=12, rows=[
              ["Dataset", "Chart", "What it shows"],
              ["SuperMart", "Column", f"South-West grew from {m(ZY.loc['South-West', 2022])} (2022) to {m(ZY.loc['South-West', 2025])} (2025)"],
              ["Mercyland", "Bar", f"{DEPT.iloc[0].Department}: {DEPT.iloc[0].Admissions:,} admissions in 2023, {DEPT.iloc[0].Admissions / DEPT.iloc[1].Admissions:.1f} times the next department"],
              ["Kingsway", "Pie", " · ".join(f"{g} {v:.1%}" for g, v in CVF["grades"].items())]])),
    ],
    takeaways=["A PivotChart is drawn from a pivot and stays in sync with it.",
               "It follows new data and filters; a normal chart does not.",
               "Axis = Rows, Legend = Columns, Values = Values.",
               "Pick the chart type by the shape of the data; start columns at 0.",
               "Filter and sort with the field buttons or a slicer."],
)

MODULES = [M29, M210, M211]
