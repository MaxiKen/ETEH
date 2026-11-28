"""Week 1 edition-3 slide copy: Modules 1.7 (formulas), 1.8 (Excel tables) and 1.9 (tidy data laws).
Formula results are computed from table tStores25 (FY2025 store summary) in _build/w1_data.py."""
from v3kit import T

M17 = dict(
    id="1.7", title="Excel Formulas and Core Functions",
    subtitle="Writing formulas, order of operations, cell references, the essential functions and fixing errors.",
    objectives=["Tell a formula from a function and write both with correct structure.",
                "Apply BODMAS and the Excel operators correctly.",
                "Use relative, absolute and mixed references, switching with F4.",
                "Use AutoSum, Insert Function and the six essential functions.",
                "Recognise and fix formula errors, and copy formulas safely across a dataset."],
    topics=[
        T("1.7.1", "What is a Formula?",
          defn="A formula is an expression that calculates a value. In Excel every formula starts with an equal sign "
               "(=) and can combine values, cell references, operators and functions.",
          simple="A formula is a calculation you write once. Excel shows the result in the cell and recalculates it "
                 "whenever the inputs change.",
          like="Like a recipe card: write the method once and it works for any quantity of ingredients.",
          fx=[["=146388345-29366453", "117,021,892", "Typed numbers: works, but never updates"],
              ["=E2-F2", "117,021,892", "Cell references: updates when E2 or F2 changes"],
              ["=F2/E2", "20.1%", "Margin for Lagos - Ikeja (Profit / Sales)"]],
          fx_label="FORMULAS ON tStores25 (TABLE STARTING IN A1; ROW 2 = LAGOS - IKEJA)",
          watch="Refer to cells, not typed numbers. A typed number is hidden inside the formula and goes out of date."),

        T("1.7.2", "What is a Function?",
          defn="A function is a predefined formula built into Excel that performs a specific calculation using the "
               "values (arguments) you give it, in the form =NAME(argument1, argument2, …).",
          simple="A function is a ready-made shortcut. Instead of =E2+E3+E4+… for twelve stores, you write "
                 "=SUM(E2:E13).",
          like="A function is a kitchen blender: you put ingredients in (arguments), press one button (the name) and "
               "get the result out.",
          fx=[["=SUM(tStores25[Sales])", "₦733,215,255", "Adds all values"],
              ["=MAX(tStores25[Sales])", "₦146,388,345", "Largest value"],
              ["=ROUND(F2/E2, 3)", "0.201", "Margin rounded to 3 decimals: two arguments"]],
          fx_label="FUNCTIONS AND ARGUMENTS"),

        T("1.7.3", "Difference Between a Formula and a Function",
          simple="Every function is used inside a formula, but not every formula uses a function. A formula is the "
                 "whole calculation; a function is one built-in tool inside it.",
          grid=dict(label="Formula vs function", widths=[1.2, 2.4, 2.4], size=13, rows=[
              ["", "Formula", "Function"],
              ["What it is", "Any calculation you write", "A built-in, named calculation"],
              ["Starts with", "= (always)", "Its name, inside a formula"],
              ["Example", "=F2/E2", "SUM(E2:E13)"],
              ["Both together", "=SUM(F2:F13)/SUM(E2:E13)", "Chain margin: 12.7%"]])),

        T("1.7.4", "The Equal Sign and Formula Structure",
          simple="Start with =, then the calculation. Excel colours each reference in the formula and outlines the same "
                 "cells on the sheet, which helps you check it.",
          diagram=("flow", dict(steps=["=", "SUM", "(", "E2:E13", ")", "/", "12"],
                                sub=["starts every formula", "function name", "opens arguments", "argument: a range",
                                     "closes", "operator", "constant"])),
          diagram_label="Anatomy of =SUM(E2:E13)/12",
          points=["Without =, Excel stores the text 'SUM(E2:E13)'", "Brackets must be in pairs; Excel colours each pair",
                  "Commas separate arguments (some regions use semicolons)", "Press Enter to finish; F2 to edit"],
          points_label="RULES"),

        T("1.7.5", "Arithmetic Operators and Order of Operations (BODMAS)",
          defn="Excel evaluates a formula in a fixed order: Brackets, then powers (Orders), then Division and "
               "Multiplication (left to right), then Addition and Subtraction (left to right).",
          simple="Excel does not read left to right. Use brackets to make the order you want certain.",
          grid=dict(label="Operators", widths=[1, 1.5, 1.4], size=12.5, rows=[
              ["Operator", "Meaning", "Example"], ["+  -", "Add, subtract", "=E2-F2"], ["*  /", "Multiply, divide", "=F2/E2"],
              ["^", "Power", "=2^3 gives 8"], ["%", "Percent", "=50000*7.5%"], ["&", "Join text", "=B2&\" (\"&C2&\")\""],
              ["=  <>  >  <", "Compare", "=F2>0 gives TRUE"]]),
          fx=[["=10+20*3", "70", "Multiply first"], ["=(10+20)*3", "90", "Brackets first"],
              ["=100-20/4", "95", "Divide first"], ["=50000*(1+7.5%)", "53,750", "Price plus 7.5% VAT"]],
          fx_label="BODMAS IN ACTION", order=["grid", "fx"], split=0.5),

        T("1.7.6", "Cell References and Reference Types",
          defn="A cell reference identifies a cell or range by its column letter and row number (C4, E2:E13), so a "
               "formula uses the cell's current value.",
          simple="References make formulas reusable. When copied, a reference either moves with the formula (relative) "
                 "or stays fixed (absolute).",
          grid=dict(label="Reference types", widths=[1.2, 1, 3], size=12.5, rows=[
              ["Type", "Written", "When the formula is copied"],
              ["Relative", "E2", "Column and row both shift"],
              ["Absolute", "$E$2", "Never shifts: a fixed input such as a VAT rate"],
              ["Mixed", "E$2 or $E2", "Only the part without $ shifts"],
              ["Range", "E2:E13", "A block of cells from E2 to E13"],
              ["Other sheet", "'Inputs'!B1", "A cell on another worksheet"],
              ["Structured", "tStores25[Sales]", "A table column by name (Module 1.8)"]])),

        T("1.7.7", "Relative, Absolute and Mixed References (F4)",
          simple="Put fixed inputs in their own cell and lock them with $. Press F4 while the cursor is on a reference "
                 "to cycle E2 → $E$2 → E$2 → $E2.",
          like="An absolute reference is like your home address on every letter you send: the letters change, the "
               "address stays the same.",
          grid=dict(label="Copying =E2*$B$1 down (B1 holds the VAT rate 7.5%)", widths=[1, 1.8, 2.2], size=12.5, rows=[
              ["Row", "Formula after copying", "What happened"],
              ["2", "=E2*$B$1", "Original"], ["3", "=E3*$B$1", "E moved down; $B$1 stayed"],
              ["4", "=E4*$B$1", "Same pattern for every row"]]),
          watch="Without the $, the copy in row 3 becomes =E3*B2: the rate cell slides away and the answers are wrong."),

        T("1.7.8", "Basic Math Formulas",
          simple="Most business measures are simple arithmetic on two columns. Write it once, then copy it down.",
          fx=[["Cost = Sales − Profit", "=E2-F2", "₦117,021,892"],
              ["Margin = Profit ÷ Sales", "=F2/E2", "20.1%"],
              ["Attainment = Sales ÷ Target", "=E2/H2", "103.1%"],
              ["Average order value = Sales ÷ Orders", "=E2/G2", "₦425,548"],
              ["Growth = (This year − Last year) ÷ Last year", "=(733215255-609205599)/609205599", "20.4%"]],
          fx_label="tStores25 STARTING IN A1: ROW 2 = LAGOS - IKEJA (E Sales, F Profit, G Orders, H Target)",
          fx_widths=[2.8, 2.6, 1.3], fx_head=["Measure", "Formula", "Result (Ikeja)"], order=["fx"],
          watch="Margin and attainment are ratios. Format them as %, and never add percentages across stores."),

        T("1.7.9", "Core Built-in Functions",
          simple="Excel has over 450 functions, grouped by category on the Formulas tab. A small core covers most "
                 "analyst work; later weeks add more.",
          grid=dict(label="Function categories and the core functions", widths=[1.3, 2.4, 2.4], size=12, rows=[
              ["Category", "Core functions", "Example"],
              ["Math", "SUM, ROUND, SUMIF, SUMIFS", "=SUMIF(tStores25[Zone], \"South-West\", tStores25[Sales])"],
              ["Statistical", "AVERAGE, COUNT, COUNTA, MAX, MIN, MEDIAN, COUNTIF", "=COUNTIF(tStores25[Profit],\"<0\") gives 5"],
              ["Text", "LEFT, RIGHT, UPPER, PROPER, TRIM, LEN", "=PROPER(\"LAGOS - IKEJA\")"],
              ["Date", "TODAY, YEAR, MONTH, DATEDIF", "=YEAR(\"2025-01-01\") gives 2025"],
              ["Logical", "IF, AND, OR, IFERROR", "=IF(F2<0,\"Loss\",\"Profit\")"],
              ["Lookup", "XLOOKUP, VLOOKUP (Week 2)", "Bring Target into the sales lines"]],
              note="SUMIF for South-West gives ₦250,128,366 (Ikeja + Lekki + Ibadan).")),

        T("1.7.10", "Using Insert Function (fx)",
          simple="If you don't know a function's name or its arguments, let Excel guide you. The fx button opens a "
                 "searchable list and a form for the arguments.",
          steps=["Click the cell for the result", "Click fx beside the formula bar (or press Shift+F3)",
                 "Search: type 'count cells that meet a condition', pick COUNTIF, OK",
                 "Fill the Function Arguments form: Range = tStores25[Profit], Criteria = \"<0\"",
                 "Read the preview (= 5) and the help text, then click OK"],
          steps_label="INSERT COUNTIF WITH THE fx DIALOG",
          points=["Typing =COUNTIF( shows a tooltip with the arguments", "Bold argument = the one you are typing",
                  "[Square brackets] = optional argument", "Tab accepts the suggested function name"],
          points_label="WHILE TYPING", order=["steps", "points"], split=0.58),

        T("1.7.11", "AutoSum and the Six Essential Functions",
          simple="AutoSum (Alt+=) writes a SUM for the numbers above or to the left. Its drop-down gives the other "
                 "everyday functions.",
          fx=[["=SUM(tStores25[Sales])", "₦733,215,255", "Total"],
              ["=AVERAGE(tStores25[Sales])", "₦61,101,271", "Mean per store"],
              ["=COUNT(tStores25[Sales])", "12", "Counts numbers"],
              ["=COUNTA(tStores25[Store_Name])", "12", "Counts non-empty cells (text too)"],
              ["=MAX(tStores25[Sales])", "₦146,388,345", "Largest (Lagos - Ikeja)"],
              ["=MIN(tStores25[Sales])", "₦31,585,710", "Smallest (Owerri - New Owerri)"]],
          fx_label="THE SIX ESSENTIALS ON tStores25",
          watch="COUNT ignores text. If COUNT is lower than COUNTA on a number column, some numbers are stored as text."),

        T("1.7.12", "Understanding Formula Errors",
          simple="An error value tells you what went wrong. Read it before you change anything.",
          grid=dict(label="Error values and their fixes", widths=[1.1, 2.4, 2.6], size=12, rows=[
              ["Error", "Meaning", "Usual fix"],
              ["#DIV/0!", "Dividing by zero or a blank", "Check the divisor; =IFERROR(F2/E2,\"\")"],
              ["#N/A", "A lookup found no match", "Check spelling and extra spaces"],
              ["#NAME?", "Unknown function or name", "Fix the spelling: =SUMM → =SUM"],
              ["#VALUE!", "Wrong type, e.g. text in maths", "Convert text-numbers to numbers"],
              ["#REF!", "Reference to a deleted cell", "Undo, or rewrite the reference"],
              ["#SPILL!", "Spill range is blocked", "Clear the cells in the way"],
              ["####", "Column too narrow (not an error)", "Widen the column"]]),
          watch="IFERROR hides every error, including real mistakes. Fix the cause first; use IFERROR only for "
                "expected cases."),

        T("1.7.13", "Copying and Extending Formulas Across Data",
          simple="Write the formula once in the first row, then extend it. Relative references adjust for each row.",
          grid=dict(label="Ways to extend a formula", widths=[1.8, 3], size=12.5, rows=[
              ["Method", "How"],
              ["Double-click fill handle", "Fills down to the last row of the neighbouring data"],
              ["Drag fill handle", "Fills as far as you drag"],
              ["Ctrl+D / Ctrl+R", "Fill down / right from the first cell of a selection"],
              ["Ctrl+Enter", "Enter the same formula into all selected cells at once"],
              ["Excel table", "Type once: the whole column fills automatically"]]),
          points=["Click a few cells down and read the formula bar", "Check that fixed inputs stayed fixed ($)",
                  "Check the last row: did the fill reach the end?"],
          points_label="AFTER COPYING, CHECK", order=["grid", "points"], split=0.6),

        T("1.7.14", "Practical Multi-Sector Calculations",
          simple="The same few formula patterns (ratio, average, rate) answer questions in every sector.",
          grid=dict(label="One pattern, five sectors (results from our datasets)", widths=[1.6, 1.8, 2.6, 1.3], size=12, rows=[
              ["Sector", "Measure", "Formula pattern", "Result"],
              ["Retail", "FY2025 margin", "=SUM(Profit)/SUM(Sales)", "12.7%"],
              ["University", "At-risk rate", "=COUNTIF(At_Risk,\"Yes\") / COUNTA(Matric_No)", "27.2%"],
              ["Hospital", "Average length of stay, 2023", "=AVERAGE(LOS_Days)", "5.0 days"],
              ["Hotel", "Cancellation rate, 2023-2024", "=COUNTIF(Reservation_Status,\"Cancelled\") / COUNTA(Reservation_ID)", "33.5%"],
              ["E-commerce", "Average order value", "=SUM(total_price)/COUNTA(order_id)", "₦126,792"]])),
    ],
    takeaways=["A formula starts with the = sign; a function is a named tool used inside a formula.",
               "BODMAS decides the order; brackets make your intention certain.",
               "Relative references move when copied; $ locks them; F4 cycles the types.",
               "SUM, AVERAGE, COUNT, COUNTA, MAX and MIN answer most first questions; AutoSum is Alt+=.",
               "Read error values before fixing; check copied formulas in several rows."],
)

M18 = dict(
    id="1.8", title="Organising Data with Excel Tables",
    subtitle="Datasets, records and variables, and why Excel tables (Ctrl+T) are the foundation of every analysis.",
    objectives=["Describe a dataset in terms of records (rows) and variables (columns).",
                "Convert a range to an Excel table and format, name and total it.",
                "Write structured references and calculated columns.",
                "Sort, filter and remove duplicates correctly.",
                "Explain why tables beat plain ranges."],
    topics=[
        T("1.8.1", "What is a Dataset?",
          defn="A dataset is a structured collection of related data, usually arranged as a table in which each row is "
               "one record and each column is one variable.",
          simple="A dataset is all the data about one subject in one place: every sale, every student, every admission.",
          like="A class register is a dataset: one line per pupil, one column per detail (name, age, attendance).",
          grid=dict(label="The datasets you will use", widths=[2.2, 1.4, 2.4], size=12, rows=[
              ["Dataset", "Records", "One row is…"],
              ["SuperMart sales (tSales)", "9,994", "one product line on an order"],
              ["Kingsway students (tStudents)", "5,840", "one student"],
              ["Mercyland admissions, 2023 (tAdmissions23)", "12,626", "one hospital stay"],
              ["Atlantic Pearl bookings (tBookings)", "38,400", "one room booking"],
              ["SuperMart Online orders (tOrdersOnline)", "11,000", "one order line"]])),

        T("1.8.2", "Rows as Records, Columns as Variables",
          simple="Read across a row and you learn everything about one record. Read down a column and you see one "
                 "variable for every record.",
          ex=dict(t="tSales", cols=["Row_ID", "Order_ID", "Order_Date", "Store_Name", "Category", "Sales", "Profit"], n=4,
                  note="Each row: one record (a sales line). Each column: one variable. 9,994 records × 11 variables."),
          grid=dict(label="Terms that mean the same", widths=[1, 1, 1], size=12.5, rows=[
              ["Excel", "Statistics", "Databases"], ["Row", "Observation", "Record"], ["Column", "Variable", "Field"],
              ["Header", "Variable name", "Field name"]]),
          order=["ex", "grid"], stack=True),

        T("1.8.3", "Normal Range vs Excel Table",
          defn="A range is any block of cells. An Excel table is a range converted into a named, structured object that "
               "Excel manages: it has headers, styles, filters, and grows with the data.",
          simple="Data in a plain range is just cells. Turning it into a table tells Excel 'this is one dataset', and "
                 "Excel starts helping.",
          grid=dict(label="Range vs table", widths=[1.6, 2, 2.4], size=12.5, rows=[
              ["", "Normal range", "Excel table"],
              ["Name", "A1:K9995", "tSales"],
              ["New rows", "Formulas and charts miss them", "Included automatically"],
              ["Formulas", "Copied down by hand", "Fill the column by themselves"],
              ["Filters and style", "Added by hand", "Built in, banded rows"],
              ["References", "=SUM(H2:H9995)", "=SUM(tSales[Sales])"]])),

        T("1.8.4", "Converting a Range to a Table with Ctrl+T",
          simple="Click any cell inside the data and press Ctrl+T. Excel selects the whole block, provided it has no "
                 "blank rows or columns.",
          steps=["Click one cell inside the data", "Press Ctrl+T (or Insert › Table)",
                 "Check the range Excel proposes covers all rows and columns",
                 "Keep 'My table has headers' ticked, then OK",
                 "Table Design › Table Name: type a name such as tSales"],
          steps_label="CREATE A TABLE",
          points=["Exactly one header row", "No blank rows or columns", "No merged cells", "One record per row"],
          points_label="BEFORE YOU PRESS CTRL+T", order=["steps", "points"], split=0.6,
          watch="If Excel's proposed range stops early, there is a blank row or column. Delete it and try again."),

        T("1.8.5", "Creating and Formatting Excel Tables",
          simple="A new table gets a style straight away. Table Design lets you change the style and turn parts on or "
                 "off without touching the cells.",
          grid=dict(label="Table Design › Table Style Options", widths=[1.7, 3], size=12.5, rows=[
              ["Option", "Effect"], ["Header Row", "Show or hide the header"], ["Banded Rows", "Alternate row shading for reading"],
              ["First / Last Column", "Bold the first or last column"], ["Total Row", "Adds a totals row (see 1.8.11)"],
              ["Filter Button", "Show or hide the drop-down arrows"]]),
          points=["Pick a light, simple style from Table Styles", "Format whole columns: Sales as ₦ #,##0, Discount as %",
                  "Formats extend to new rows automatically", "Convert to Range removes the table but keeps the look"],
          points_label="GOOD PRACTICE", order=["grid", "points"], split=0.55),

        T("1.8.6", "Table Headers and Structured References",
          defn="A structured reference refers to a table or its columns by name instead of by cell address, for "
               "example tStores25[Sales] for the whole Sales column.",
          simple="Names are easier to read than addresses, and they grow with the table.",
          fx=[["=SUM(tStores25[Sales])", "₦733,215,255", "Whole column"],
              ["=[@Profit]/[@Sales]", "20.1% (row 1)", "This row's values: used in a calculated column"],
              ["=COUNTIF(tStores25[Attainment],\">=1\")", "5", "Stores at or above target"],
              ["=tStores25[[#Totals],[Sales]]", "₦733,215,255", "The total row's Sales cell"]],
          fx_label="STRUCTURED REFERENCES ON tStores25",
          watch="Header names must be unique and should have no spaces. 'Store_Name' is easier to type in a formula "
                "than 'Store Name'."),

        T("1.8.7", "Sorting Data",
          simple="Sorting reorders whole rows by one or more columns. In a table, every column moves together, so "
                 "records stay intact.",
          ex=dict(t="tStores25", cols=["Store_Name", "Zone", "Sales"], n=5,
                  rows=lambda d: d.sort_values(["Zone", "Sales"], ascending=[True, False]),
                  note="Sorted by Zone (A to Z), then Sales (largest first): first 5 of 12 rows."),
          steps=["Quick sort: header arrow › Sort Largest to Smallest", "Several levels: Data › Sort (Alt, A, S, S)",
                 "Add Level: Zone A to Z, then Sales Largest to Smallest", "OK"],
          steps_label="MULTI-LEVEL SORT", order=["ex", "steps"], split=0.5,
          watch="In a plain range, sorting one selected column alone scrambles the records. Always sort the whole table."),

        T("1.8.8", "Filtering Data",
          simple="A filter hides the rows you don't need, without deleting them. Each column's arrow offers text, "
                 "number or date filters.",
          ex=dict(t="tStores25", cols=["Store_Name", "Zone", "Profit"], rows=lambda d: d[d.Profit < 0], n=5,
                  note="Number Filters › Less Than 0 on Profit: 5 of 12 stores shown."),
          grid=dict(label="Filter types", widths=[1.2, 2.4], size=12.5, rows=[
              ["Type", "Example"], ["Text", "Zone equals South-West"], ["Number", "Profit less than 0"],
              ["Date", "Order_Date in 2025, or Last Month"], ["Top 10", "Top 3 stores by Sales"]]),
          order=["ex", "grid"],
          watch="SUM still counts hidden rows. Use the table's Total Row (SUBTOTAL) to total only the visible rows."),

        T("1.8.9", "Removing Duplicates",
          simple="Duplicate records inflate counts and totals. Remove Duplicates deletes rows whose chosen columns "
                 "repeat an earlier row.",
          grid=dict(label="Duplicates in the dirty SuperMart file (tDirtySales)", widths=[1.9, 2.4], size=12.5, rows=[
              ["Check", "Result"], ["Rows before", "3,387"], ["Repeated records", "60 (e.g. Row_ID 6748 appears twice)"],
              ["Empty rows (also removed)", "15"], ["Rows after cleaning", "3,312: matches the clean FY2025 file"]]),
          steps=["Copy the sheet first (it cannot be undone after saving)", "Click in the table: Data › Remove Duplicates",
                 "Tick the columns that define a duplicate (here: all, or just Row_ID)",
                 "OK: Excel reports how many it removed"],
          steps_label="HOW", order=["grid", "steps"], split=0.48,
          watch="Choosing too few columns (e.g. only Customer_Name) deletes real records. Decide what makes a row unique first."),

        T("1.8.10", "Auto-Expansion and Calculated Columns",
          defn="A calculated column is a table column whose formula, typed once, fills every row automatically and "
               "continues into new rows.",
          simple="Type a new row under a table and the table grows to include it. Type a formula in an empty column "
                 "and it fills to the bottom.",
          steps=["In the first empty column to the right of the table, type the header Margin",
                 "In the first row type =[@Profit]/[@Sales] and press Enter",
                 "Every row fills; format the column as %",
                 "Add a new row at the bottom: Margin calculates itself"],
          steps_label="ADD A MARGIN COLUMN TO tStores25",
          points=["Charts and PivotTables based on the table include new rows", "Formulas like =SUM(tStores25[Sales]) "
                  "include them too", "Undo (Ctrl+Z) once to stop a formula filling the column"],
          points_label="WHY IT MATTERS", order=["steps", "points"], split=0.58),

        T("1.8.11", "Table Design, Total Row and Naming Tables",
          simple="Give every table a meaningful name, and use the Total Row for quick totals that respect filters.",
          grid=dict(label="Total Row (Table Design › Total Row, or Ctrl+Shift+T)", widths=[1.4, 2.6], size=12.5, rows=[
              ["Choose", "Excel writes"], ["Sum", "=SUBTOTAL(109,[Sales])"], ["Average", "=SUBTOTAL(101,[Sales])"],
              ["Count", "=SUBTOTAL(103,[Store_Name])"], ["Max / Min", "=SUBTOTAL(104 / 105, …)"]],
              note="SUBTOTAL with 101-111 ignores rows hidden by a filter."),
          points=["Start with t, then a clear noun: tSales, tStores25", "No spaces; letters, numbers and _ only",
                  "Rename in Table Design › Table Name", "All names: Formulas › Name Manager (Ctrl+F3)"],
          points_label="NAMING TABLES", order=["grid", "points"], split=0.55),

        T("1.8.12", "Why Excel Tables Are Better Than Normal Ranges",
          simple="Tables remove a whole class of mistakes: missed rows, formulas not copied, hard-to-read references.",
          grid=dict(label="Six reasons to use tables", widths=[1.8, 3.4], size=12.5, rows=[
              ["Benefit", "What it prevents"],
              ["Grows automatically", "New rows missed by formulas, charts and PivotTables"],
              ["Calculated columns", "Formulas copied to only some rows"],
              ["Structured references", "Unreadable formulas such as =SUM(H2:H9995)"],
              ["Built-in sort and filter", "Sorting one column and scrambling records"],
              ["Total Row with SUBTOTAL", "Totals that include hidden rows"],
              ["Named object", "Confusion about which data a report uses: Power Query and PivotTables use the name"]])),
    ],
    takeaways=["A dataset is rows (records) × columns (variables), with one header row.",
               "Ctrl+T turns a range into a table; name it at once (tSales, tStores25).",
               "Structured references (tStores25[Sales], [@Profit]) are readable and grow with the data.",
               "Sort the whole table; filters hide, never delete; copy the sheet before Remove Duplicates.",
               "Calculated columns and the Total Row remove whole classes of mistakes."],
)

M19 = dict(
    id="1.9", title="Twelve Tidy Data Laws",
    subtitle="Twelve rules that keep data in the one shape every Excel, Power Query, SQL and Power BI tool expects.",
    objectives=["Define tidy data and explain why tools depend on it.",
                "Recognise breaches of each of the twelve laws in real data.",
                "Fix untidy data: split, unpivot, standardise, separate raw from report.",
                "Document a dataset with a simple data dictionary."],
    topics=[
        T("1.9.1", "The Tidy Data Foundation",
          defn="Tidy data is data in which each variable forms a column, each observation forms a row, and each type of "
               "observational unit forms a table (Hadley Wickham, 2014).",
          simple="One shape for all data: columns = variables, rows = records, cells = single values. Tables, "
                 "PivotTables, Power Query, SQL and Power BI all expect it.",
          like="Like a well-packed shop shelf: one product per slot, every slot labelled. You find anything at once.",
          grid=dict(label="The twelve laws", widths=[1, 1], size=12, rows=[
              ["Laws 1-6", "Laws 7-12"],
              ["1  One record per row", "7  Blank and zero mean different things"],
              ["2  One variable per column", "8  Keep raw data separate from reports"],
              ["3  Atomic values: one value per cell", "9  Control your vocabulary"],
              ["4  One header row", "10  Store calculations, not typed results"],
              ["5  No spacer rows or blank columns", "11  No merged cells"],
              ["6  ISO dates (yyyy-mm-dd)", "12  Document everything"]])),

        T("1.9.2", "Law 1: One Record per Row",
          defn_label="THE LAW", defn="Each row describes exactly one record (one sale, one admission), and every record "
                                     "appears exactly once.",
          simple="Duplicate rows, subtotal rows and rows holding two records all break counts and totals, because every "
                 "row is counted as one record.",
          grid=dict(label="Breaches and fixes", widths=[1.7, 2.6, 1.8], size=12.5, rows=[
              ["Breach", "Example", "Fix"],
              ["Duplicate row", "Row_ID 6748 twice in tDirtySales (60 repeats)", "Remove Duplicates"],
              ["Subtotal row inside data", "'Lagos total' typed between sales lines", "Delete; total in a report"],
              ["Two records in one row", "Two products typed in one order line", "One row per product"]])),

        T("1.9.3", "Law 2: One Variable per Column",
          defn_label="THE LAW", defn="Each column holds one variable, in one form. A variable is never spread across "
                                     "several columns, and two variables never share a column.",
          simple="Months spread across columns (Jan, Feb, Mar…) is the most common breach. It looks like a report, but "
                 "it cannot be filtered by month or summed easily.",
          grid=dict(label="Wide (untidy) vs long (tidy)", widths=[1.4, 1, 1, 0.3, 1.4, 1, 1], size=12, rows=[
              ["Store", "Q1", "Q2", "", "Store", "Quarter", "Sales"],
              ["Lagos - Ikeja", "₦…", "₦…", "→", "Lagos - Ikeja", "Q1", "₦…"],
              ["Lagos - Lekki", "₦…", "₦…", "", "Lagos - Ikeja", "Q2", "₦…"],
              ["", "", "", "", "Lagos - Lekki", "Q1", "₦…"]],
              note="Fix: unpivot the quarter columns into Quarter + Sales (Power Query, Week 2)."),
          watch="A Discount column mixing 0.2 and '20%' is also a breach: one variable, two forms."),

        T("1.9.4", "Law 3: Atomic Values (One Value per Cell)",
          defn_label="THE LAW", defn="Each cell holds a single value that cannot sensibly be split further for the "
                                     "analysis.",
          simple="Combined values are hard to filter, sort and count. Split them into separate columns.",
          grid=dict(label="Combined → atomic", widths=[2, 2.8], size=12.5, rows=[
              ["Combined (one cell)", "Atomic (separate columns)"],
              ["Abuja, FCT (Store_Location)", "City: Abuja  |  State: FCT"],
              ["chuKwuemEKa f. odili", "First: Chukwuemeka  |  Initial: F.  |  Surname: Odili"],
              ["₦1,399,930", "Sales: 1399930 (number; ₦ is a format)"]]),
          steps=["Data › Text to Columns › Delimited › Comma", "Or Flash Fill (Ctrl+E) for name parts",
                 "Or Power Query › Split Column (Week 2)"],
          steps_label="HOW TO SPLIT", order=["grid", "steps"], split=0.62),

        T("1.9.5", "Law 4: One Header Row",
          defn_label="THE LAW", defn="Data starts with exactly one row of unique column names, directly above the "
                                     "first record.",
          simple="Titles, logos and dates above the data, or two stacked header rows, confuse Tables, PivotTables and "
                 "Power Query.",
          grid=dict(label="Two stacked header rows → one header row", widths=[1.5, 1.2, 1.2, 1.2], size=12.5, rows=[
              ["", "Column A", "Column B", "Column C"],
              ["Untidy: row 1", "(blank)", "Q1", "(merged over C)"],
              ["Untidy: row 2", "Store", "Sales", "Profit"],
              ["Tidy: one row", "Store", "Q1_Sales", "Q1_Profit"]]),
          points=["Short, unique names: Order_Date, Store_Name", "No spaces (use _), no line breaks",
                  "Put titles and notes on another sheet", "Never leave a header blank"],
          points_label="GOOD HEADERS", order=["grid", "points"], split=0.55),

        T("1.9.6", "Law 5: No Spacer Rows or Blank Columns",
          defn_label="THE LAW", defn="A dataset is one continuous block: no empty rows between records and no empty "
                                     "columns between variables.",
          simple="People add blank rows to make sheets look nicer, but Excel treats a blank row as the end of the data.",
          grid=dict(label="What a blank row breaks", widths=[1.6, 3], size=12.5, rows=[
              ["Tool", "Effect of one blank row"], ["Ctrl+Down", "Stops at the gap"],
              ["Ctrl+T / Ctrl+A", "Select only the part above the gap"],
              ["Sort", "Rows below the gap are left behind"], ["Filter", "Rows below the gap are ignored"]]),
          steps=["Select a column, F5 › Special › Blanks", "Check the selected rows are truly empty",
                 "Home › Delete › Delete Sheet Rows"],
          steps_label="FIND AND REMOVE BLANK ROWS", order=["grid", "steps"], split=0.55,
          watch="tDirtySales has 15 completely empty rows. Use formatting (borders, banding) for spacing, never empty rows."),

        T("1.9.7", "Law 6: ISO Dates (yyyy-mm-dd)",
          defn_label="THE LAW", defn="Dates are stored as real Excel dates, in one column, displayed in the ISO 8601 "
                                     "format yyyy-mm-dd.",
          simple="04/05/2025 is 4 May in Nigeria and 5 April in the USA. 2025-05-04 means the same everywhere.",
          grid=dict(label="Date formats found in tDirtySales (335 non-ISO rows)", widths=[1.5, 1.9, 1.6], size=12.5, rows=[
              ["As exported", "Problem", "ISO"], ["14/05/2025", "Day-first: may be read as text", "2025-05-14"],
              ["04-May-2025", "Text, not a date", "2025-05-04"], ["2025.06.10", "Dots: text, not a date", "2025-06-10"],
              ["2025-03-18", "Already correct", "2025-03-18"]]),
          points=["Test: =ISNUMBER(A2) is TRUE for a real date", "Format: Ctrl+1 › Custom › yyyy-mm-dd",
                  "Convert text dates with DATE, Text to Columns or Power Query"],
          points_label="CHECK AND FIX", order=["grid", "points"], split=0.6),

        T("1.9.8", "Law 7: Missing Values: Blank vs Zero",
          defn_label="THE LAW", defn="A blank means unknown or not applicable; a zero means the value was measured and "
                                     "is nothing. Never use one for the other.",
          simple="AVERAGE ignores blanks but includes zeros, so typing 0 for 'not yet known' drags averages down.",
          grid=dict(label="Same four scores, two ways to record the missing one", widths=[1.6, 1.4, 1.4], size=12.5, rows=[
              ["", "Blank used", "Zero used"], ["Scores", "70, 80, 90, (blank)", "70, 80, 90, 0"],
              ["=AVERAGE()", "80", "60"], ["=COUNT()", "3", "4"]]),
          points=["Discount 0 = no discount given (a real zero)", "Discount blank = not recorded",
                  "Never type 'N/A', '-' or 'none' in a number column", "Count blanks: =COUNTBLANK(range)"],
          points_label="IN PRACTICE", order=["grid", "points"], split=0.5),

        T("1.9.9", "Law 8: Separate Raw Data from Reports",
          defn_label="THE LAW", defn="Keep the original data exactly as received, do the work in a separate layer, and "
                                     "build summaries on their own sheets.",
          simple="If something goes wrong, you can always start again from the raw data.",
          diagram=("flow", dict(steps=["RAW: data as received", "CLEAN: fixed copy or Power Query", "REPORT: summaries and charts"],
                                sub=["Never edited", "Every change repeatable", "Reads only from the clean layer"])),
          diagram_label="Three layers",
          grid=dict(label="In a workbook", widths=[1.2, 3], size=12.5, rows=[
              ["Sheet", "Contents"], ["Raw", "The import, untouched (e.g. tDirtySales)"],
              ["Clean", "The fixed table (e.g. tSales)"], ["Report", "Totals, PivotTables, charts"]])),

        T("1.9.10", "Law 9: Control Your Vocabulary",
          defn_label="THE LAW", defn="Each category value is written exactly one way, from an agreed list.",
          simple="To Excel, 'Lagos - Ikeja', 'LAGOS - IKEJA' and 'lagos - ikeja' are three different stores. Counts and "
                 "PivotTables split one group into many.",
          grid=dict(label="Store_Name in tDirtySales: one store, three spellings", widths=[2, 1], size=12.5, rows=[
              ["Value", "Rows"], ["Lagos - Ikeja", "642"], ["lagos - ikeja", "21"], ["LAGOS - IKEJA", "19"]]),
          points=["Fix existing values: TRIM, PROPER, Find & Replace", "Prevent new ones: Data Validation lists",
                  "Watch hidden spaces: ' North-Central' is not 'North-Central'"],
          points_label="FIX AND PREVENT", order=["grid", "points"], split=0.45),

        T("1.9.11", "Law 10: Store Calculations, Not Typed Results",
          defn_label="THE LAW", defn="Derived values are calculated with formulas from the data, never typed in as "
                                     "fixed numbers.",
          simple="A typed 103% never changes and nobody can see how it was worked out. =Sales/Target updates and shows "
                 "its logic in the formula bar.",
          fx=[["Typed: 103%", "Wrong when sales are corrected", "Hidden logic"],
              ["=E2/H2", "103.1%, updates automatically", "Logic visible"],
              ["=[@Sales]/[@Target]", "Same, in a table: fills every row", "Readable"]],
          fx_label="ATTAINMENT FOR LAGOS - IKEJA", fx_head=["In the cell", "Result", "Why"],
          watch="Paste Special › Values is fine for a final snapshot you send out, never inside your working data."),

        T("1.9.12", "Laws 11 and 12: No Merged Cells; Document Everything",
          defn_label="THE LAWS", defn="Law 11: never merge cells inside data. Law 12: document what each file and column "
                                      "means, its unit, its source and every change made.",
          simple="Merged cells break sorting, filtering and PivotTables. A data dictionary lets anyone, including you in "
                 "six months, understand the data.",
          grid=dict(label="A data dictionary for tStores25 (first rows)", widths=[1.3, 3, 1.4], size=12, rows=[
              ["Column", "Meaning", "Unit / type"], ["Store_ID", "Store code, ST01-ST12", "Text"],
              ["Sales", "Total FY2025 sales for the store", "₦, number"],
              ["Target", "Annual sales target for 2025", "₦, number"],
              ["Attainment", "Sales ÷ Target", "%, formula"]]),
          order=["grid"],
          watch="Use Center Across Selection for titles instead of merging (see 1.6.10)."),
    ],
    takeaways=["Tidy = one variable per column, one record per row, one value per cell.",
               "One header row, no blank rows or columns, no merged cells.",
               "ISO dates; blank is not zero; one spelling per category.",
               "Raw → clean → report; formulas, not typed results.",
               "Document every dataset with a data dictionary."],
)

MODULES = [M17, M18, M19]
