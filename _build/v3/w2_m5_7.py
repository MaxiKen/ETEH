"""Week 2 edition-3 slide copy: Modules 2.5 (cleaning in Excel), 2.6 (Power Query transforms), 2.7 (Excel vs PQ).
Practice file throughout: SuperMart_Sales_2025_DIRTY.csv, loaded as table tDirty25. Figures from w2_fx."""
from v3kit import T
from w2_fx import DF, F, CTRL, N, I, P, Nm, S

CC = DF["cat_counts"]
BAND = F["bands"]
ZN = F["zone"]
FR = F["first_rows"]


def _cat(*names):
    return ", ".join(f"{n} ({CC[n]})" for n in names)


M25 = dict(
    id="2.5", title="Data Cleaning with Excel",
    subtitle="Measure every problem in the dirty 2025 sales export, fix it with Excel's own tools and formulas, then prove "
             "the result against the POS control totals.",
    objectives=["Describe the twelve common data quality problems.", "Remove duplicates and blank rows safely.",
                "Fix numbers and dates stored as text.", "Clean text and standardise categories.",
                "Verify a cleaned dataset against control totals."],
    topics=[
        T("2.5.1", "Understanding Dirty Data",
          defn="Dirty data is any value you cannot trust for the job: duplicates, numbers stored as text, one category "
               "spelt several ways, missing or impossible values.",
          simple="Cleaning is not making the sheet look nice. It is making every value clear, correctly typed and consistent, "
                 "so totals and counts come out true.",
          like="Like rice with small stones: fine in the bag, but cook it without picking and someone bites a stone.",
          grid=dict(label=f"Measured in SuperMart_Sales_2025_DIRTY.csv ({DF['rows']:,} lines)", widths=[2.6, 1.6, 0.8],
                    size=12, rows=[["Problem found", "Example", "Count"],
                                   ["Repeated records", "Row_ID 7026 twice", I(DF["dup_lines"])],
                                   ["Sales stored as text", DF["text_sales_example"], I(DF["sales_text"])],
                                   ["Order_Date stored as text", "14/05/2025", I(DF["text_dates"])],
                                   ["Category spellings for 3 categories", "Furnitre, Tech", "12"],
                                   ["Missing Quantity", "(blank)", I(DF["qty_blank"])]]),
          watch="Count each problem before fixing anything. The count is your proof that the cleaning worked."),

        T("2.5.2", "The Twelve Common Data Quality Problems",
          simple="Use a checklist, not guesswork. Write one detection formula per problem and run it over the whole column.",
          grid=dict(label="The twelve problems", widths=[1, 1], size=12, rows=[
              ["Problem", "Problem"], ["1  Duplicates", "7  Bad or mixed dates"], ["2  Missing values", "8  Out-of-range values"],
              ["3  Inconsistent case and spacing", "9  Combined values in one column"],
              ["4  Inconsistent categories", "10  Hidden characters"], ["5  Wrong data types", "11  Mixed units"],
              ["6  Numbers stored as text", "12  Structure: blank rows, extra headers"]]),
          fx=[["=COUNTBLANK(tDirty25[Order_ID])", I(DF["blank_rows"]), "Blank rows"],
              ["=SUMPRODUCT(--ISTEXT(tDirty25[Sales]))", I(DF["sales_text"]), "Sales stored as text"],
              ["=COUNTIF(tDirty25[Customer_Name],\"N/A\")", I(DF["na_cust"]), "Placeholder names"],
              ["=COUNTIF(tDirty25[Quantity],\"<0\")", I(DF["qty_neg"]), "Negative quantities"]],
          fx_label="ONE DETECTION FORMULA PER PROBLEM", order=["grid", "fx"], stack=True),

        T("2.5.3", "Removing Duplicates",
          defn="A duplicate is a record that appears more than once. Data › Remove Duplicates deletes every repeat of the "
               "columns you tick, keeping the first.",
          simple="Some POS lines were exported twice. Row_ID should be unique; tick all columns so only exact copies go.",
          steps=["Copy the sheet first (Remove Duplicates cannot be audited)", "Click in the table › Data › Remove Duplicates",
                 "Keep every column ticked › OK", "Read the message: values removed and remaining"],
          grid=dict(label="Result", widths=[2.4, 1.2], size=12.5, rows=[
              ["Check", "Value"], ["Rows before (blank rows removed)", I(DF["after_blank"])],
              ["Repeated records removed", I(DF["dup_lines"])], ["Rows after", I(DF["after_blank"] - DF["dup_lines"])],
              ["Sales the copies added", Nm(DF["dup_sales"])]]),
          order=["steps", "grid"], split=0.52,
          watch=f"Run it before deleting blank rows and Excel reports {DF['dup_all']} removed: the {DF['blank_rows'] - 1} "
                f"extra blank rows count as copies of the first."),

        T("2.5.4", "Finding and Replacing Values",
          simple="Find & Replace (Ctrl+H) changes every match at once. By default it matches part of a cell, which can "
                 "damage good values.",
          grid=dict(label=f"Replace 'Tech' with 'Technology' in Category ({CC['Tech']} cells say 'Tech')", widths=[1.8, 1, 2.4],
                    size=12, rows=[
                        ["Setting", "Cells changed", "Result"],
                        ["Default (part of cell)", I(DF["tech_partial"]), "'Technology' becomes 'Technologynology'"],
                        ["Match entire cell contents", I(CC["Tech"]), "Only the 'Tech' cells fixed"]]),
          steps=["Select the Category column only", "Ctrl+H › Find: Tech › Replace with: Technology",
                 "Options › tick Match entire cell contents", "Replace All; confirm the count"],
          order=["grid", "steps"], split=0.6,
          watch="Select the column first. With one cell selected, Replace All works on the whole sheet."),

        T("2.5.5", "Fixing Numbers Stored as Text",
          simple="A number typed with ₦ or commas arrives as text. It sits left-aligned and SUM ignores it.",
          fx=[["=VALUE(SUBSTITUTE(SUBSTITUTE(R2,\"₦\",\"\"),\",\",\"\"))", "239240",
               "Strip ₦ and commas, then convert"],
              ["=VALUE(SUBSTITUTE(T2,\"%\",\"\"))/100", "0.2", "'20%' text to 0.2"],
              ["=SUMPRODUCT(--ISTEXT(tDirty25[Sales]))", "0", "Re-count after the fix"]],
          grid=dict(label="Before and after", widths=[2.2, 1.1, 1.1], size=12, rows=[
              ["Measure", "Before", "After"], ["Sales cells stored as text", I(DF["sales_text"]), "0"],
              ["Value hidden from SUM", Nm(DF["text_sales_value"]), "₦0"],
              ["Discount cells stored as text", I(DF["disc_text"]), "0"]]),
          order=["fx", "grid"], stack=True,
          watch="Paste the converted column back as Values before deleting the original."),

        T("2.5.6", "Cleaning Text with TRIM and CLEAN",
          defn="TRIM removes leading and trailing spaces and reduces inner runs of spaces to one. CLEAN removes the "
               "non-printing characters with codes 0 to 31, such as the line feed CHAR(10).",
          simple="Invisible characters make 'South-West' and ' South-West' two different zones.",
          fx=[["=TRIM(L2)", "South-West", "Extra spaces removed"],
              ["=CLEAN(G2)", "Omolara Yusuf", "Line feed removed; CHAR(160) stays"],
              ["=TRIM(CLEAN(SUBSTITUTE(G2, CHAR(160), \" \")))", "Omolara Yusuf", "Every hidden character fixed"]],
          fx_widths=[3.3, 1.4, 1.7],
          grid=dict(label="Before and after", widths=[2.2, 0.8, 0.8], size=12, rows=[
              ["Check", "Before", "After"], ["Zone cells with extra spaces", I(DF["zone_space"]), "0"],
              ["Distinct Zone values", I(DF["zone_distinct"]), I(DF["zone_distinct_trim"])],
              ["Names with CHAR(160)", I(DF["nbsp"]), "0"], ["Names with a line feed", I(DF["lf"]), "0"]]),
          order=["fx", "grid"], split=0.62,
          watch="CLEAN does not remove CHAR(160). Substitute it with a normal space first, then TRIM."),

        T("2.5.7", "Standardising Categorical Values",
          simple="One category spelt several ways splits every count and total. Map each variant to one standard value.",
          grid=dict(label=f"Category: {DF['cat_bad']} cells in 9 non-standard spellings", widths=[3.4, 1.4], size=12, rows=[
              ["Variants (count)", "Standard"],
              [_cat("Office Supply", "Office-Supplies", "office supplies"), "Office Supplies"],
              [_cat("FURNITURE", "furniture", "Furnitre"), "Furniture"],
              [_cat("Technolgy", "Tech", "technology"), "Technology"],
              ["PivotTable rows", "12 before, 3 after"]]),
          fx=[["=XLOOKUP(TRIM(O2), tCatMap[Variant], tCatMap[Standard], TRIM(O2))", "Mapping table: one row per variant"],
              ["=PROPER(O2)", "Fixes case only, not spelling"]],
          order=["grid", "fx"], stack=True),

        T("2.5.8", "Fixing Incorrect Dates",
          simple="A date stored as text cannot be sorted, filtered by month or subtracted. Repair each pattern with its own formula.",
          fx=[[f"dd/mm/yyyy ({DF['d_slash']})", "14/05/2025", "=DATE(RIGHT(C2,4), MID(C2,4,2), LEFT(C2,2))"],
              [f"dd-Mon-yyyy ({DF['d_mon']})", "04-May-2025", "=DATEVALUE(C2)"],
              [f"yyyy.mm.dd ({DF['d_dots']})", "2025.06.10", "=DATEVALUE(SUBSTITUTE(C2, \".\", \"-\"))"]],
          fx_head=["Pattern (count)", "Example", "Fix"], fx_widths=[1.5, 1.2, 3.2],
          fx_label=f"{DF['text_dates']} ORDER_DATE VALUES ARE TEXT, IN THREE PATTERNS",
          points=["=ISNUMBER(C2) is TRUE for a real date", "Format the result as yyyy-mm-dd to check it",
                  f"After the fix: =SUMPRODUCT(--ISTEXT(…)) = 0"],
          points_label="CHECK", order=["fx", "points"], stack=True,
          watch="DATEVALUE reads dd/mm/yyyy using your PC's regional settings. On a US-set PC 03/04/2025 becomes 4 March. "
                "DATE with LEFT/MID/RIGHT is safe everywhere."),

        T("2.5.9", "Handling Missing Values",
          simple="A blank is not zero. Decide for each column whether to recover, flag or exclude, and write the decision down.",
          grid=dict(label="Decisions for the dirty sales file", widths=[1.8, 0.7, 2.4], size=12, rows=[
              ["Problem", "Count", "Decision"], ["Quantity blank", I(DF["qty_blank"]), "Flag; ask the store (do not type 0)"],
              ["Customer_Name = 'N/A'", I(DF["na_cust"]), "Set to blank; report to IT"],
              ["Negative Quantity", I(DF["qty_neg"]), "Flag as possible returns"],
              ["Ship_Date before Order_Date", I(DF["ship_before"]), "Flag; query the store"]]),
          fx=[["=IF(S2=\"\", \"Missing\", \"\")", "Flag column"],
              ["=AVERAGEIFS(tDirty25[Quantity], tDirty25[Quantity], \"<>\")", "Averages ignore blanks"]],
          order=["grid", "fx"], split=0.58,
          watch="Typing 0 into a blank quantity changes averages and makes Sales ÷ Quantity a #DIV/0! error."),

        T("2.5.10", "Removing Blank Rows and Empty Columns",
          simple="Blank rows break a table into pieces: sorting, filtering and Ctrl+A stop at the first gap.",
          grid=dict(label="Structure defects", widths=[1.6, 2.4, 1.3], size=12, rows=[
              ["Defect", "How to detect", "Fix"],
              [f"{DF['blank_rows']} blank rows", "Go To Special › Blanks on Order_ID", "Delete Sheet Rows"],
              ["Empty Column1 (M)", "=COUNTA(M:M) = 1 (header only)", "Delete the column"],
              ["Rows after fixing", f"{DF['rows']:,} → {DF['after_blank']:,}", "One continuous block"]]),
          steps=["Select the Order_ID column", "Home › Find & Select › Go To Special › Blanks",
                 "Home › Delete › Delete Sheet Rows"],
          order=["grid", "steps"], split=0.6,
          watch="Use a key column that is never blank in a real row. Blanks in the whole sheet would also catch real rows."),

        T("2.5.11", "Using Text to Columns",
          defn="Text to Columns splits one column into several at a delimiter or at fixed widths.",
          simple="Product_ID holds three codes in one. Split it so the manager can filter by category code.",
          grid=dict(label="Product_ID split at '-'", widths=[1.6, 0.9, 0.9, 1.1], size=12, rows=[
              ["Product_ID", "Cat_Code", "Sub_Code", "Item_No"], ["TEC-AC-10003614", "TEC", "AC", "10003614"],
              ["OFF-ST-10002485", "OFF", "ST", "10002485"], ["FUR-FU-10001940", "FUR", "FU", "10001940"]]),
          steps=["Insert 3 empty columns to the right", "Data › Text to Columns › Delimited",
                 "Other: - › Next", "Set Item_No to Text › Finish"],
          order=["grid", "steps"], split=0.55,
          watch="Text to Columns overwrites the cells to the right without asking. Make room first."),

        T("2.5.12", "Verifying a Clean Dataset",
          defn="Control totals are agreed figures from the source system (row count, totals) that a cleaned dataset must match.",
          simple="Cleaning is finished only when every control total matches.",
          grid=dict(label="POS control totals vs the cleaned table", widths=[2.2, 1.5, 1.5], size=12.5, rows=[
              ["Check", "POS control", "Clean data"], ["Rows", I(CTRL["rows"]), I(CTRL["rows"])],
              ["Total Sales", N(CTRL["sales"]), N(CTRL["sales"])], ["Total Profit", N(CTRL["profit"]), N(CTRL["profit"])],
              ["Total Quantity", I(CTRL["qty"]), I(CTRL["qty"])], ["Distinct orders", I(CTRL["orders"]), I(CTRL["orders"])]]),
          fx=[["=ROWS(tClean)", "Row count"], ["=SUM(tClean[Sales])", "Total Sales"],
              ["=COUNTA(UNIQUE(tClean[Order_ID]))", "Distinct orders"]],
          order=["grid", "fx"], split=0.55,
          watch=f"Some Sales values were keyed 100 times too large (largest {N(DF['top_sales'][0])}). Totals will not "
                f"match until they are corrected from the receipts: find them with a sort or MAX."),
    ],
    takeaways=["Measure each problem first; the count proves the fix.",
               "Remove duplicates on all columns, after removing blank rows.",
               "Convert text numbers with SUBSTITUTE + VALUE; repair text dates by pattern.",
               "TRIM, CLEAN and CHAR(160) for hidden characters; a mapping table for categories.",
               "Done means every control total matches."],
)

M26 = dict(
    id="2.6", title="Cleaning and Transforming in Power Query",
    subtitle="The same clean-up as Module 2.5, built as recorded steps: remove, replace, type, split, add columns, group, "
             "pivot and refresh.",
    objectives=["Remove rows, columns, blanks and duplicates as steps.", "Replace values and set types with the right locale.",
                "Split, merge, trim and extract text.", "Add conditional, custom and index columns.",
                "Group, pivot and unpivot; manage steps and refresh."],
    topics=[
        T("2.6.1", "Removing Rows and Columns",
          simple="Remove what you do not need early: every later step then runs on less data.",
          grid=dict(label="Remove the empty Column1", widths=[1.6, 2.6], size=12.5, rows=[
              ["Item", "Value"], ["Columns before", f"{DF['cols']} (including Column1)"], ["Columns after", str(DF['cols'] - 1)],
              ["Step created", "Removed Columns"], ["M code", "Table.RemoveColumns(Source, {\"Column1\"})"]]),
          points=["Home › Remove Columns, or Choose Columns to keep", "Home › Remove Rows › Top, Bottom, Alternate, Errors",
                  "Remove Other Columns keeps only the selected ones"],
          points_label="WHERE", order=["grid", "points"], split=0.58,
          watch="Prefer Remove Other Columns: new columns added to the source later are then ignored, not loaded."),

        T("2.6.2", "Removing Blank Rows",
          simple="Home › Remove Rows › Remove Blank Rows deletes rows where every value is empty or null.",
          grid=dict(label="SuperMart_Sales_2025_DIRTY.csv", widths=[2.2, 1], size=12.5, rows=[
              ["Stage", "Rows"], ["After Promoted Headers", I(DF["rows"])], ["After Remove Blank Rows", I(DF["after_blank"])],
              ["Rows removed", I(DF["blank_rows"])]]),
          fx=[["Table.SelectRows(#\"Removed Columns\", each not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), "
               "{\"\", null})))", "The step Excel writes"]],
          fx_head=["Step formula", "Meaning"], fx_widths=[4.2, 1.4], order=["grid", "fx"], stack=True),

        T("2.6.3", "Removing Duplicate Records",
          simple="Home › Remove Rows › Remove Duplicates keeps the first row for each combination of the selected columns. "
                 "The columns you select are the key.",
          grid=dict(label=f"{DF['after_blank']:,} rows after removing blanks: which key?", widths=[2.4, 1, 2.2], size=12,
                    rows=[["Key selected", "Rows left", "Correct?"],
                          ["All columns (exact copies)", I(DF["after_blank"] - DF["dup_lines"]), "Yes: matches the POS control"],
                          ["Order_ID only", I(CTRL["orders"]), f"No: {CTRL['rows'] - CTRL['orders']:,} real lines lost"]]),
          points=["Select all columns (Ctrl+A in the preview) first", "One order has several lines: Order_ID is not unique",
                  "Unlike Excel, Power Query treats 'abc' and 'ABC' as different"],
          points_label="REMEMBER", order=["grid", "points"], stack=True),

        T("2.6.4", "Filtering Rows",
          simple="Filtering in Power Query removes the rows from the result; it does not just hide them.",
          grid=dict(label="Keep the loss-making 2025 lines (Profit < 0)", widths=[2.2, 1.4], size=12.5, rows=[
              ["Measure", "Value"], ["Lines before the filter", I(CTRL["rows"])], ["Lines with Profit < 0", I(F["loss_lines"])],
              ["Sales on those lines", Nm(F["loss_sales"])], ["Loss on those lines", Nm(F["loss_profit"])]]),
          steps=["Click the Profit filter arrow", "Number Filters › Less Than… › 0",
                 "Step: Table.SelectRows(…, each [Profit] < 0)"],
          order=["grid", "steps"], split=0.5,
          watch="Filtering by ticking values from the list hard-codes those values. Use Number or Text Filters for rules."),

        T("2.6.5", "Replacing Values",
          simple="Transform › Replace Values swaps one value for another in the selected columns. It matches the whole "
                 "cell unless you choose otherwise.",
          grid=dict(label="Replacements in the dirty sales file", widths=[1.4, 1.2, 1.4, 0.6], size=12, rows=[
              ["Column", "Find", "Replace with", "Cells"], ["Category", "Furnitre", "Furniture", I(CC["Furnitre"])],
              ["Category", "FURNITURE", "Furniture", I(CC["FURNITURE"])], ["Customer_Name", "N/A", "null", I(DF["na_cust"])]]),
          points=["Case-sensitive: FURNITURE and furniture need separate steps", "Advanced options › Match entire cell contents",
                  "For many variants, merge with a mapping table instead"],
          points_label="REMEMBER", order=["grid", "points"], split=0.58),

        T("2.6.6", "Changing Data Types with Locale",
          defn="Change Type › Using Locale converts text to a type using another region's rules, e.g. dates in dd/mm/yyyy "
               "order or numbers with a comma as the decimal mark.",
          simple="The same text, 03/04/2025, is 3 April in Nigeria and 4 March in the USA. Tell Power Query which.",
          grid=dict(label=f"{DF['d_slash']} Order_Date values written dd/mm/yyyy", widths=[1.9, 0.8, 1.2, 0.9], size=12, rows=[
              ["Locale used", "Errors", "Silently wrong", "Correct"],
              ["English (United States)", I(DF["us_err"]), I(DF["us_wrong"]), I(DF["us_ok"])],
              ["English (United Kingdom)", "0", "0", I(DF["d_slash"])]]),
          steps=["Right-click Order_Date › Change Type › Using Locale…", "Data Type: Date; Locale: English (United Kingdom)",
                 "Check column quality: 0% Error"],
          order=["grid", "steps"], split=0.58,
          watch=f"The {DF['us_wrong']} silently wrong dates raise no error. Only the right locale prevents them."),

        T("2.6.7", "Splitting Columns by Delimiter",
          simple="Transform › Split Column › By Delimiter splits text at a character, at each occurrence or only the first or last.",
          grid=dict(label="Store_Location split at ', '", widths=[2, 1.4, 1.2], size=12.5, rows=[
              ["Store_Location", "City", "State"], ["Benin City, Edo", "Benin City", "Edo"],
              ["Port Harcourt, Rivers", "Port Harcourt", "Rivers"], ["Uyo, Akwa Ibom", "Uyo", "Akwa Ibom"]]),
          points=["Delimiter: Custom ', ' (comma and space)", "Split at: Left-most delimiter",
                  "Advanced: split into Rows for lists"],
          points_label="SETTINGS", order=["grid", "points"], split=0.6,
          watch="Split at each occurrence fixes the number of output columns. A value with an extra comma later spills into a new column that never loads."),

        T("2.6.8", "Merging Columns",
          simple="Transform › Merge Columns joins several text columns with a separator into one.",
          grid=dict(label="Store_ID and Store_Name merged with ' · '", widths=[0.9, 2, 2.6], size=12, rows=[
              ["Store_ID", "Store_Name", "Store_Label"], ["ST05", "Benin City - Ring Road", "ST05 · Benin City - Ring Road"],
              ["ST01", "Lagos - Ikeja", "ST01 · Lagos - Ikeja"], ["ST12", "Owerri - New Owerri", "ST12 · Owerri - New Owerri"]]),
          points=["Transform › Merge Columns replaces the originals", "Add Column › Merge Columns keeps them",
                  "Numbers must be text first"],
          points_label="TWO PLACES", order=["grid", "points"], split=0.64),

        T("2.6.9", "Trimming and Cleaning Text",
          simple="Transform › Format › Trim and Clean do what TRIM and CLEAN do in Excel, as steps.",
          grid=dict(label="Before and after", widths=[2.2, 1.5, 1.2], size=12, rows=[
              ["Check", "Before", "After"], ["Zone cells with extra spaces", I(DF["zone_space"]), "0"],
              ["Distinct Zone values", I(DF["zone_distinct"]), I(DF["zone_distinct_trim"])],
              ["'Omolara Yusuf' + line feed", "A different customer", "Same customer"]]),
          points=["Text.Trim removes leading and trailing spaces only", "Inner double spaces stay (Excel's TRIM removes them)",
                  "Replace CHAR(160) with Text.Replace(_, \"#(00A0)\", \" \")"],
          points_label="DIFFERENCES FROM EXCEL", order=["grid", "points"], split=0.55),

        T("2.6.10", "Extracting Text from Columns",
          simple="Add Column › Extract takes part of the text: first or last characters, a range, or text before or after a delimiter.",
          grid=dict(label="Category code from Product_ID", widths=[1.8, 1.2], size=12.5, rows=[
              ["Product_ID (kept)", "Cat_Code (new)"], ["FUR-FU-10001940", "FUR"], ["OFF-ST-10002485", "OFF"],
              ["TEC-AC-10003614", "TEC"]]),
          fx=[["Text.Start([Product_ID], 3)", "First Characters: 3"], ["Text.BeforeDelimiter([Product_ID], \"-\")", "Text Before Delimiter"],
              ["Text.AfterDelimiter([Product_ID], \"-\", 1)", "After the 2nd '-': 10001940"]],
          fx_head=["M formula", "Menu option"], order=["grid", "fx"], split=0.4),

        T("2.6.11", "Adding Conditional Columns",
          simple="Add Column › Conditional Column builds an if-then-else rule from a dialog: no formula typing.",
          grid=dict(label="Disc_Band on the clean 2025 lines", widths=[1.2, 0.8, 1.1, 1.1], size=12, rows=[
              ["Disc_Band", "Lines", "Sales", "Profit"]] + [
              [b, I(r.Lines), Nm(r.Sales), Nm(r.Profit)] for b, r in BAND.iterrows()]),
          steps=["Add Column › Conditional Column", "If Discount equals 0 › None", "Else if Discount is less than or equal "
                 "to 0.2 › 1-20%", "Else › Above 20%"],
          order=["grid", "steps"], split=0.52,
          watch="Rules are tested top to bottom; the first match wins. Put the narrowest rule first."),

        T("2.6.12", "Adding Custom Columns",
          simple="Add Column › Custom Column takes an M formula that runs once per row. Column names go in square brackets.",
          grid=dict(label="Margin = [Profit] / [Sales]", widths=[0.9, 1.2, 1.2, 0.9], size=12, rows=[
              ["Row_ID", "Sales", "Profit", "Margin"]] + [
              [str(r["Row_ID"]), N(r["Sales"]), N(r["Profit"]), P(r["Profit"] / r["Sales"])] for r in FR]),
          fx=[["[Profit] / [Sales]", "Margin"], ["Date.Year([Order_Date])", "Year"],
              ["if [Quantity] = null then \"Missing\" else \"OK\"", "Flag"]],
          fx_head=["Custom column formula", "Result"], order=["grid", "fx"], split=0.5,
          watch="M is case-sensitive: [profit] fails if the column is Profit. Set the new column's type afterwards."),

        T("2.6.13", "Creating Index Columns",
          simple="Add Column › Index Column numbers the rows 0, 1, 2… or 1, 2, 3… in their current order.",
          grid=dict(label="Line_No on the clean 2025 lines, in date order", widths=[0.9, 1.3, 0.9], size=12.5, rows=[
              ["Line_No", "Order_Date", "Row_ID"], ["1", "2025-01-01", str(FR[0]["Row_ID"])], ["2", "2025-01-01", str(FR[1]["Row_ID"])],
              [I(CTRL["rows"]), F["last_row"]["Order_Date"].strftime("%Y-%m-%d"), str(F["last_row"]["Row_ID"])]]),
          points=["Sort first: the index follows the current order", "Use it to keep the original order before merges",
                  "Or to number rows within groups (Group By › All Rows)"],
          points_label="USES", order=["grid", "points"], split=0.5),

        T("2.6.14", "Grouping Data with Group By",
          defn="Group By collapses rows that share the same values in the chosen columns into one row each, with aggregations "
               "(sum, count, average, max) for the other columns.",
          grid=dict(label="Group By Zone: Sum of Sales, Sum of Profit", widths=[1.4, 1, 1, 0.8], size=12, rows=[
              ["Zone", "Sales", "Profit", "Margin"]] + [[z, Nm(r.Sales), Nm(r.Profit), P(r.Margin)] for z, r in ZN.iterrows()]),
          steps=["Home › Group By › Advanced", "Group by: Zone", "New column Sales = Sum of Sales; Profit = Sum of Profit",
                 "Add Column › Custom: [Profit] / [Sales]"],
          order=["grid", "steps"], split=0.55,
          watch="Group By replaces the detail rows. Do it in a Reference query so the detailed query stays available."),

        T("2.6.15", "Pivoting and Unpivoting Columns",
          simple="Pivot turns the values of one column into column headers. Unpivot does the reverse: headers back into rows.",
          ex=dict(t="tZoneCat", cols=["Zone", "Furniture", "Office Supplies", "Technology"], n=4, size=11.5, rows=lambda d: d.head(4),
                  note="Pivot Category, values Sales (Sum): the Zone × Category grid."),
          points=["Pivot: select Category › Transform › Pivot Column › Values: Sales", "Unpivot: select Zone › Unpivot Other Columns",
                  "Analysis wants the long form; reports want the wide form"],
          points_label="HOW", order=["ex", "points"], split=0.6),

        T("2.6.16", "Handling Missing and Null Values",
          defn="null is Power Query's empty value: no value at all, which is different from zero or an empty text string.",
          simple="Decide what a null means before replacing it.",
          grid=dict(label=f"{DF['qty_blank']} empty Quantity values; Unit_Price = [Sales] / [Quantity]", widths=[1.7, 2, 0.8],
                    size=12, rows=[
                        ["Option", "Effect", "Verdict"], ["Replace null with 0", "Division by zero gives infinity", "Wrong"],
                        ["Leave null and flag", "Unit_Price null; row flagged", "Right"],
                        ["Remove the rows", "Their sales vanish from totals", "Wrong"]]),
          points=["Replace Values › null for blanks", "Fill Down for report-style gaps", "Keep Errors / Remove Errors to inspect"],
          points_label="TOOLS", order=["grid", "points"], split=0.64),

        T("2.6.17", "Sorting and Filtering in Power Query",
          simple="Sort decides which row comes first; Remove Duplicates then keeps that first row.",
          grid=dict(label="One row per 2025 customer, showing their latest order", widths=[2.6, 0.8, 1.4], size=12, rows=[
              ["Approach", "Rows", "SMC-0388 shows"],
              ["Sort newest first, remove duplicates on Customer_ID", I(CTRL["customers"]),
               F["smc0388"][2].strftime("%Y-%m-%d")],
              ["Sort oldest first, remove duplicates", I(CTRL["customers"]), F["smc0388"][1].strftime("%Y-%m-%d")]]),
          points=[f"SMC-0388 (Zuwaira Salami) has {F['smc0388'][0]} lines in 2025", "Add Table.Buffer after the sort so it is kept",
                  "Filter before sorting: less data to sort"],
          points_label="NOTE", order=["grid", "points"], split=0.62),

        T("2.6.18", "Reordering and Renaming Columns",
          simple="Drag headers to reorder; double-click a header to rename. Clear names make every later step readable.",
          grid=dict(label="Transactions sheet of SuperMart_Online_RAW_Export.xlsx", widths=[1.6, 1.6, 2], size=12, rows=[
              ["Raw name", "Renamed", "Why"], ["transaction_id", "Transaction_ID", "Consistent style"],
              ["customer_id", "Customer_ID", "Matches the Customers sheet"], ["total_price", "Total_Price_NGN", "Shows the currency"]]),
          watch="Later steps refer to columns by name. Rename early, before steps that use the column."),

        T("2.6.19", "Applied Steps and Transformation History",
          simple="Each step can be renamed, moved, deleted or edited with its gear icon. Rename steps so the list reads like a report.",
          grid=dict(label="Before handing the query over", widths=[1.8, 2.6], size=12, rows=[
              ["Default name", "Renamed to"], ["Removed Blank Rows", f"Removed {DF['blank_rows']} spacer rows"],
              ["Removed Duplicates", f"Removed {DF['dup_lines']} duplicate lines"],
              ["Replaced Value", "Fixed Furnitre spelling"], ["Changed Type1", "Typed dates as UK (dd/mm)"]]),
          points=["Right-click a step › Rename, Delete, Move", "Delete Until End removes all later steps",
                  "Right-click › Properties to add a description"],
          points_label="ACTIONS", order=["grid", "points"], split=0.6),

        T("2.6.20", "Refreshing Queries After Source Changes",
          simple="Refresh reruns every step on the new source. Most changes are absorbed; renamed columns are not.",
          grid=dict(label="What happens at Refresh", widths=[1.6, 1.8, 2], size=12, rows=[
              ["Source change", "Effect", "Prevention"], ["New rows or new month file", "Rows increase; works", "Check the counts"],
              ["Column renamed", "Steps fail: column not found", "Agree fixed names; rename early"],
              ["File moved", "Source step fails", "Parameter for the path"],
              ["New column added", "Ignored if you used Remove Other Columns", "Choose columns to keep"]]),
          watch="After every refresh, re-check the control totals. A query can run without errors and still be wrong."),
    ],
    takeaways=["Remove columns, blanks and duplicates first; choose the duplicate key deliberately.",
               "Change Type with the right locale; wrong locales create silent errors.",
               "Split, merge, trim and extract work like their Excel counterparts, as steps.",
               "Conditional, custom and index columns add logic without formulas in cells.",
               "Group By and Pivot summarise; Unpivot tidies; rename steps and check totals after refresh."],
)

M27 = dict(
    id="2.7", title="Excel vs Power Query Cleaning",
    subtitle="When manual cleaning is right, when automation pays, and how to prove how a number was produced.",
    objectives=["Say when manual Excel cleaning is the right choice.", "List Power Query's strengths.",
                "Weigh the cost of repeating manual work.", "Show how a number was produced.",
                "Choose the right tool for each task."],
    topics=[
        T("2.7.1", "Manual Cleaning in Excel: What It Is Good For",
          simple="Manual cleaning is fine for one-off, small, visible fixes, where building a query would take longer than the job.",
          grid=dict(label="Four requests in one week", widths=[2.6, 1.2, 1.2], size=12.5, rows=[
              ["Request", "Repeats?", "Tool"], ["Tidy 40 names for a one-off letter", "No", "Excel"],
              ["Fix one typo in a finished report", "No", "Excel"], ["Clean the monthly POS export", "Every month", "Power Query"],
              ["Combine 12 monthly files", "Every year", "Power Query"]])),

        T("2.7.2", "Automated Cleaning in Power Query: Strengths",
          simple="The same steps, every time, by anyone: one answer instead of several.",
          grid=dict(label="Three analysts clean the dirty file by hand", widths=[2.5, 0.9, 1.6], size=12, rows=[
              ["Who cleaned it", "Rows left", "Effect on Sales"],
              ["Analyst A (forgot duplicates)", I(DF["after_blank"]), f"{Nm(DF['dup_sales'])} too high"],
              ["Analyst B (forgot text Sales)", I(CTRL["rows"]), f"{Nm(DF['text_sales_clean'])} missing"],
              ["Analyst C (did everything)", I(CTRL["rows"]), "Correct"], ["Power Query (every refresh)", I(CTRL["rows"]), "Correct"]]),
          points=["Repeatable and consistent", "Every step visible", "Handles more rows than a sheet", "Source never changed"],
          points_label="STRENGTHS", order=["grid", "points"], split=0.66),

        T("2.7.3", "Reproducibility and the Cost of Manual Cleaning",
          defn="A result is reproducible when someone else, following the same recorded steps on the same data, gets the same answer.",
          simple="Manual work costs the same every month. A query costs more once, then almost nothing.",
          grid=dict(label="Example figures: replace them with your own", widths=[2, 1.2, 1.4], size=12.5, rows=[
              ["Approach", "First month", "Whole year"], ["Manual cleaning, 2 hours a month", "2 hours", "24 hours"],
              ["Power Query, 3 hours to build", "3 hours", "About 4 hours"]]),
          watch="The cost of manual cleaning is not only time: every repeat is another chance to make a different mistake."),

        T("2.7.4", "Auditability: Proving How a Number Was Produced",
          defn="Auditability is the ability to show where a figure came from and every change made on the way.",
          simple="An auditor asks why reported 2025 sales are lower than the raw export suggests. The query answers.",
          grid=dict(label="Audit questions", widths=[1.8, 1.3, 2.3], size=12, rows=[
              ["Question", "Manual answer", "Power Query answer"], ["Where did the data come from?", "'An export'", "Source step shows the file"],
              ["Why fewer rows?", "'I removed some'", f"{DF['dup_lines']} duplicates, {DF['blank_rows']} blank rows"],
              ["What was the effect?", "Unknown", f"Duplicates added {Nm(DF['dup_sales'])}"]])),

        T("2.7.5", "Choosing the Right Tool for Each Task",
          simple="Ask two questions: will it repeat, and is it data preparation or presentation?",
          grid=dict(label="A monthly management pack", widths=[2.4, 1.3, 1.6], size=12.5, rows=[
              ["Task", "Tool", "Reason"], [f"Clean {CTRL['rows']:,} POS lines", "Power Query", "Monthly and large"],
              ["Combine the 12 monthly files", "Power Query", "Repeats every month"],
              ["Write the 3-line summary", "Excel (typed)", "Needs human judgement"], ["Format the final table", "Excel", "Presentation only"]])),

        T("2.7.6", "Practical Comparison Using the Same Dataset",
          simple="The dirty sales file cleaned twice: once by hand in Excel, once in Power Query. Count the actions.",
          grid=dict(label="Actions needed", widths=[2.2, 1.3, 1.5], size=12.5, rows=[
              ["Task", "Excel actions", "Power Query steps"], ["Blank rows, duplicates", "2", "2"],
              ["Text Sales, text dates", "7", "3"], ["Categories, split, trim", "6", "4"],
              ["Next month's file", "Repeat all 15", "Refresh: 0"]]),
          watch="Both routes must end at the same control totals. Check them whichever tool you use."),
    ],
    takeaways=["Manual cleaning suits one-off, small, visible fixes.",
               "Power Query gives one repeatable answer, with every step visible.",
               "Reproducible means anyone following the steps gets the same result.",
               "Auditable means you can show the source and every change.",
               "Repeating preparation → Power Query; judgement and presentation → Excel."],
)

MODULES = [M25, M26, M27]
