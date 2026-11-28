"""Week 2 Section Examples (Edition 3.1): questions in teaching order.

Progression rule for Week 2 (Week 1 skills are assumed)
    2.1 - 2.4  importing files and Power Query: read counts from what Excel / Power Query shows; no new functions
    2.5        cleaning in Excel with the functions shown in 2.5 (TRIM, CLEAN, PROPER, COUNTBLANK, ROWS)
    2.6 - 2.7  Power Query transforms; choosing Excel or Power Query
    2.8A/B     the 65 functions: one simple use each, on the 12-row tStores table where possible
    2.9        M language basics
    2.10-2.11  PivotTables and PivotCharts on tSM
"""
from __future__ import annotations

import pandas as pd

import w2_data as W
from common import CORE, W2D

DELIMS = ["Comma", "Semicolon", "Pipe", "Tab"]


def build(b):
    st = W.table("tStores")
    s25 = W.table("tSales25")
    dirty = W.table("tDirty25")
    sm = W.sales()

    mods = {}

    def M(code, sheet, title, learned):
        mods[code] = b.module(code, sheet, title, learned)
        return mods[code]

    m21 = M("2.1", "2.1 External Formats", "Understanding External Data Sources",
            "External data, file formats (.xlsx, .csv, .txt, .json), delimiters and text qualifiers, importing safely "
            "with Data › From Text/CSV, and locale and encoding traps.")
    m22 = M("2.2", "2.2 Power Query Intro", "Introduction to Power Query",
            "What Power Query is, the ETL cycle (Extract, Transform, Load), the Power Query Editor, Applied Steps, "
            "column quality and distribution, managing queries and loading results.")
    m23 = M("2.3", "2.3 Connect & Combine", "Connecting and Importing with Power Query",
            "Connecting to workbooks, CSV and text files, importing several sheets, combining every file in a folder, "
            "file paths and parameters, and refreshing.")
    m24 = M("2.4", "2.4 Web & PDF Sources", "Web and PDF Data Sources",
            "Reading a web page before importing, extracting a PDF table, unpivoting a cross-tab, reconciling to the "
            "ledger, and an audit-ready refresh.")
    m25 = M("2.5", "2.5 Excel Cleaning", "Data Cleaning with Excel",
            "Dirty data, the twelve common problems, Remove Duplicates, Find & Replace, numbers stored as text, TRIM and "
            "CLEAN, standardising categories, fixing dates, missing values, Text to Columns and control totals.")
    m26 = M("2.6", "2.6 PQ Transforms", "Cleaning and Transforming in Power Query",
            "Removing rows and columns, duplicates, filtering, replacing values, data types with locale, splitting and "
            "merging columns, trimming, conditional and custom columns, Group By, pivot / unpivot, Applied Steps.")
    m27 = M("2.7", "2.7 Excel vs PQ", "Excel vs Power Query Cleaning",
            "When manual Excel cleaning is fine, when Power Query is better, reproducibility and auditability.")
    m28a = M("2.8A", "2.8A Functions 1-33", "The 65 Analytical Excel Functions (1-33)",
             "SUM to TEXTJOIN: conditional totals and counts (SUMIF(S), COUNTIF(S), AVERAGEIFS), MEDIAN, LARGE, SMALL, "
             "ROUND, ABS, IF, IFS, AND, OR, IFERROR, TEXT and joining text.")
    m28b = M("2.8B", "2.8B Functions 34-65", "Text, Date, Lookup and Dynamic Array Functions",
             "LEFT, RIGHT, MID, LEN, FIND, SUBSTITUTE, UPPER, PROPER, TEXTBEFORE, dates (DATE, YEAR, EOMONTH, "
             "NETWORKDAYS), lookups (XLOOKUP, VLOOKUP, INDEX/MATCH) and dynamic arrays (UNIQUE, FILTER, SORT).")
    m29 = M("2.9", "2.9 M Language", "M Language",
            "What M is, why it is case-sensitive, let ... in, data types, the formula bar, the Advanced Editor and "
            "common M functions.")
    m210 = M("2.10", "2.10 PivotTables", "PivotTables",
             "Pivot data requirements, the four areas, creating a pivot, aggregations, filters, grouping dates, "
             "Show Values As, refreshing, slicers and timelines.")
    m211 = M("2.11", "2.11 PivotCharts", "PivotCharts",
             "PivotCharts, how they differ from normal charts, field buttons, the link to the PivotTable, chart types "
             "and filtering.")

    # ------------------------------------------------------------------ data sheets
    sales_fmt = {"Row_ID": "gen", "Order_Date": "date", "Ship_Date": "date", "Sales": "naira", "Quantity": "int",
                 "Discount": "pct0", "Profit": "naira", "Year": "gen"}
    r_st = b.data("Stores", st, "One SuperMart store with its FY2025 target and sales (12 stores)",
                  fmts={"Opened": "date", "Floor_Area_sqm": "int", "Target_2025": "naira", "Sales_2025": "naira"},
                  widths={"Store_Name": 27, "Store_Manager": 18}, used_in="2.8A, 2.8B")
    r_s25 = b.data("Sales25", s25, "One clean SuperMart sales line in FY2025", fmts=sales_fmt,
                   widths={"Product_Name": 40, "Store_Name": 26}, used_in="2.8A")
    r_d = b.data("Dirty25", dirty, "One line of SuperMart_Sales_2025_DIRTY.csv exactly as Excel opens it (untidy on purpose)",
                 fmts={"Order_Date": "date", "Ship_Date": "date", "Row_ID": "gen"}, used_in="2.5")
    r_sm = b.data("SM", sm, "One clean SuperMart sales line, 2022-2025 (the PivotTable source)", fmts=sales_fmt,
                  widths={"Product_Name": 40, "Store_Name": 26}, used_in="2.10, 2.11")

    # files next to the workbook
    f21 = m21.link_files([W2D / "SuperMart_Store_Targets_EU_semicolon.csv", W2D / "Mercyland_Episodes_2023_pipe.txt",
                          W2D / "SuperMart_Sales_2025_DIRTY.csv", W2D / "SuperMart_Online_Sellers.json"])
    f22 = m22.link_files([W2D / "SuperMart_Sales_2025_DIRTY.csv"])
    f23 = m23.link_files([W2D / "SuperMart_Monthly_Sales_2025", W2D / "SuperMart_Online_RAW_Export.xlsx",
                          CORE / "SuperMart_Stores_Sales_2022_2025.csv"])
    f24 = m24.link_files([W2D / "SuperMart_Store_Directory.html", W2D / "SuperMart_Monthly_Sales_Bulletin_2025.pdf"])
    f26 = m26.link_files([W2D / "SuperMart_Sales_2025_DIRTY.csv"])
    m25.set_refs([r_d])
    m28a.set_refs([r_st, r_s25])
    m28b.set_refs([r_st])
    m210.set_refs([r_sm])
    m211.set_refs([r_sm])

    raw = pd.read_csv(W2D / "SuperMart_Sales_2025_DIRTY.csv", dtype=str, keep_default_na=False, encoding="utf-8-sig")
    nonblank = raw[~(raw == "").all(axis=1)]
    n_blank = len(raw) - len(nonblank)
    n_after = len(nonblank.drop_duplicates())
    pipe_n = len(pd.read_csv(W2D / "Mercyland_Episodes_2023_pipe.txt", sep="|", dtype=str))
    months = sorted((W2D / "SuperMart_Monthly_Sales_2025").glob("*.csv"))
    month_rows = [len(pd.read_csv(f)) for f in months]
    stores_csv_n = len(pd.read_csv(CORE / "SuperMart_Stores_Sales_2022_2025.csv"))

    # ================================================================== 2.1
    m = m21
    m.part("Part A · Formats and delimiters")
    m.mc("What is external data?", ["Data typed straight into the current workbook",
                                    "Data stored outside the workbook (files, databases, web pages) that you bring in",
                                    "Only data from the internet", "A chart copied from PowerPoint"], "B",
         "External = it lives outside the workbook and is imported or connected.")
    m.pick(f"Which delimiter separates the values in {f21} › SuperMart_Store_Targets_EU_semicolon.csv?", DELIMS,
           "Semicolon", "Open it in Notepad: the values are separated by ; (common in European files).")
    m.pick(f"Which delimiter separates the values in {f21} › Mercyland_Episodes_2023_pipe.txt?", DELIMS, "Pipe",
           "The | character (called a pipe) separates the values.")
    m.pick(f"Which delimiter separates the values in {f21} › SuperMart_Sales_2025_DIRTY.csv?", DELIMS, "Comma",
           "CSV = comma-separated values.")
    m.mc('In the CSV, a value is written as "Abuja, FCT" (in double quotes). Why are the quotes there?',
         ["To make the text bold", "So the comma inside the value is not read as a delimiter",
          "To mark the value as a number", "They are a mistake"], "B",
         "The double quote is the text qualifier: everything inside it is one value.")
    m.mc("In the European targets file, floor area is written 2400,0. What does it mean?",
         ["24,000", "2,400.0 (the comma is the decimal mark)", "Text that cannot be used", "240"], "B",
         "In many European locales the comma is the decimal mark. Tell Excel the file's locale when you import.")
    m.mc("After importing, you see â‚¦ where the file had ₦. What went wrong?",
         ["The numbers are too big", "The file was read with the wrong encoding", "The delimiter is wrong", "Nothing"], "B",
         "A UTF-8 file read as Windows-1252 shows â‚¦. Pick the right File Origin (encoding) when importing.")
    m.yn("Is double-clicking a CSV the safest way to open it for analysis?", "No",
         "Use Data › From Text/CSV so you control the delimiter, encoding and data types.")
    m.part("Part B · Try it in Excel")
    m.num(f"Data › From Text/CSV › {f21} › SuperMart_Store_Targets_EU_semicolon.csv (delimiter: Semicolon). "
          "How many store rows load (not counting the header)?", None, 12, "12 stores, one row each.", do=True)
    m.num(f"Import {f21} › Mercyland_Episodes_2023_pipe.txt with the delimiter set to Pipe (or Custom: |). "
          "How many rows load (not counting the header)?", None, pipe_n,
          f"{pipe_n:,} admission episodes. The count shows in the Queries & Connections pane after loading.", do=True)
    m.mc("SuperMart_Online_Sellers.json comes from a web service (API). What kind of format is JSON?",
         ["A picture format", "A text format that stores records as nested name : value pairs",
          "A compressed Excel workbook", "A PDF"], "B", "JSON is text organised as {name: value} records.")

    # ================================================================== 2.2
    m = m22
    m.part("Part A · Power Query basics")
    m.mc("What is Power Query?",
         ["A chart type", "Excel's tool for importing, cleaning and shaping data with repeatable steps",
          "A macro language", "A type of PivotTable"], "B", "Power Query records each step so it can be replayed.")
    m.mc("What does ETL stand for?", ["Enter, Type, Leave", "Extract, Transform, Load", "Excel, Table, List",
                                     "Edit, Test, Link"], "B", "Get the data, clean and shape it, then load it.")
    m.mc("Where does Power Query list every step you have applied?",
         ["Name Box", "Applied Steps (in Query Settings)", "Status Bar", "Formula Bar only"], "B",
         "Applied Steps is on the right of the Power Query Editor.")
    m.yn("Does Power Query change the original source file?", "No",
         "Power Query reads the source; the original file stays exactly as it was.")
    m.tf("When the source file changes, clicking Refresh replays all the recorded steps.", True,
         "That is the main reason to use Power Query.")
    m.mc("What is the difference between Duplicate and Reference for a query?",
         ["There is no difference",
          "Duplicate copies all the steps; Reference starts a new query from the output of the first",
          "Reference deletes the query", "Duplicate only works for CSV files"], "B",
         "Reference keeps one source of truth; Duplicate makes an independent copy.")
    m.mc("Which Load option lets you choose a table on a sheet, a PivotTable, or only a connection?",
         ["Close & Load To...", "Save As", "Refresh All", "Undo"], "A", "Home › Close & Load › Close & Load To...")
    m.part("Part B · Try it in Excel")
    m.do(f"Data › From Text/CSV › {f22} › SuperMart_Sales_2025_DIRTY.csv › Transform Data. Turn on View › Column quality. "
         "What percentage 'Empty' does the column called Column1 show? (Type the number with %, e.g. 50%.)", "100%",
         "Column1 has a header but no values at all: 100% empty. It can be removed.", alts=("100", "1"))

    # ================================================================== 2.3
    m = m23
    m.part("Part A · Connecting and combining")
    m.mc("What does Get Data › From Folder › Combine Files do?",
         ["Deletes old files", "Applies the same steps to every file in the folder and stacks the results into one table",
          "Zips the folder", "Opens each file in a new window"], "B", "One query handles every file in the folder.")
    m.mc("Next month a new file is saved into the folder. What do you do to include it?",
         ["Rebuild the query", "Click Refresh", "Copy and paste it in", "Nothing: it cannot be added"], "B",
         "Refresh re-reads the folder and picks up the new file.")
    m.mc("Why store a folder path in a parameter?",
         ["It looks nicer", "So you change the path in one place when the files move", "It makes queries run faster",
          "Excel requires it"], "B", "A parameter is a named value used by many queries.")
    m.tf("A workbook connection can list several sheets so you can choose which one to load.", True,
         "The Navigator lists every sheet and table in the workbook.")
    m.part("Part B · Try it in Excel")
    m.num(f"How many CSV files are in {f23} › SuperMart_Monthly_Sales_2025?", None, len(months),
          "One file per month of 2025.", do=True)
    m.num(f"Open {f23} › SuperMart_Monthly_Sales_2025 › 2025-01_SuperMart_Sales.csv with Data › From Text/CSV. "
          "How many rows does it have (not counting the header)?", None, month_rows[0],
          f"January 2025 has {month_rows[0]} sales lines.", do=True)
    m.num(f"Get Data › From File › From Folder › {f23} › SuperMart_Monthly_Sales_2025 › Combine & Load. "
          "How many rows does the combined table have?", None, sum(month_rows),
          f"All 12 months stacked: {sum(month_rows):,} rows (the FY2025 sales lines).", do=True)
    m.num(f"Get Data › From File › From Excel Workbook › {f23} › SuperMart_Online_RAW_Export.xlsx. "
          "How many sheets does the Navigator list?", None, 5,
          "Transactions, Sellers, Products, Deliveries and Customers.", do=True)
    m.num(f"Connect to {f23} › SuperMart_Stores_Sales_2022_2025.csv. How many rows load?", None, stores_csv_n,
          f"{stores_csv_n:,} sales lines from 2022 to 2025.", do=True)

    # ================================================================== 2.4
    m = m24
    m.part("Part A · Web pages and PDFs")
    m.num(f"Open {f24} › SuperMart_Store_Directory.html in your browser. How many tables are on the page?", None, 2,
          "Table 1: Store directory and Table 2: Opening hours. Read a page before you import it.", do=True)
    m.num("How many stores are listed in Table 1 of the web page?", None, 12, "ST01 to ST12.", do=True)
    m.do("In Table 1 of the web page, which city is store ST04 in?", "Port Harcourt", "ST04 = Port Harcourt - GRA.")
    m.num(f"Get Data › From File › From PDF › {f24} › SuperMart_Monthly_Sales_Bulletin_2025.pdf. Open Table001 and ignore "
          "the header and the Total row. How many store rows are there?", None, 12, "One row per store.", do=True)
    m.num("The bulletin has 12 stores × 12 month columns. After Unpivot Other Columns, how many rows will the tidy table "
          "have?", "=12*12", 144, "Unpivot makes one row per store per month: 12 × 12 = 144.")
    m.mc("What is reconciliation?", ["Making a chart", "Checking that figures from one source agree with the system of "
                                     "record and explaining every difference", "Deleting duplicates", "Saving as PDF"], "B",
         "Compare the PDF with the ledger, store by store.")
    m.yn("Adding up the rounded monthly figures in the PDF gives ₦732.9m, but the ledger says ₦733.2m. Does this prove the "
         "PDF is wrong?", "No", "The PDF shows figures rounded to ₦0.1m; adding rounded figures causes small differences.")

    # ================================================================== 2.5
    m = m25
    m.example("detection formulas on tDirty25",
              ["Each formula below runs over a whole column of tDirty25 and counts one kind of problem.",
               "Write the count down before you change anything, so you can prove what you fixed."],
              [("Rows received", "=ROWS(tDirty25)", "=ROWS(tDirty25)", len(dirty), "int"),
               ("Blank Order_ID cells", "=COUNTBLANK(tDirty25[Order_ID])", "=COUNTBLANK(tDirty25[Order_ID])",
                int((dirty.Order_ID.isna() | (dirty.Order_ID.astype(str).str.strip() == "")).sum()), "int")])
    m.part("Part A · Cleaning functions")
    m.mc("Which function removes spaces at the start and end of text (and doubled spaces inside)?",
         ["CLEAN", "TRIM", "PROPER", "LEN"], "B", "TRIM removes extra spaces.")
    m.mc("Which function removes non-printing characters such as line breaks?", ["CLEAN", "TRIM", "UPPER", "SUM"], "A",
         "CLEAN removes characters with codes 0 to 31.")
    m.short('What does =TRIM("   South-West  ") return? (Type the result.)', "South-West",
            "The spaces at both ends are removed.")
    m.short('What does =PROPER("LAGOS - IKEJA") return?', "Lagos - Ikeja", "PROPER capitalises the first letter of each word.")
    m.mc("Numbers stored as text usually ...",
         ["sit on the right of the cell and add up normally", "sit on the left of the cell and are skipped by SUM",
          "show #DIV/0!", "turn red automatically"], "B", "Text sits left by default and SUM ignores it.")
    m.mc("Which tool splits 'Abuja, FCT' into 'Abuja' and 'FCT'?",
         ["Text to Columns (Delimited, comma)", "Freeze Panes", "Remove Duplicates", "AutoSum"], "A",
         "Data › Text to Columns › Delimited › Comma.")
    m.part("Part B · Numbers from the dirty file")
    m.num("How many rows does tDirty25 have? (=ROWS(tDirty25))", "=ROWS(tDirty25)", len(dirty), "ROWS counts the data rows.")
    m.num("The clean file must have 3,312 rows (the till system's control total). How many rows must cleaning remove "
          "from tDirty25?", "=ROWS(tDirty25)-3312", len(dirty) - 3312,
          f"{len(dirty):,} − 3,312 = {len(dirty) - 3312}: the 15 blank rows and the 60 repeated rows.")
    m.yn("Is cleaning finished when the row count and totals match the control totals?", "Yes",
         "Control totals prove the clean data is complete and correct.")
    m.yn("A Quantity cell is blank. Should you type 0 into it?", "No",
         "Blank means unknown; 0 means 'none sold'. Flag it and ask the store.")

    # ================================================================== 2.6
    m = m26
    m.part("Part A · Which transform?")
    T = ["Unpivot Columns", "Group By", "Split Column", "Conditional Column"]
    m.pick("Month columns Jan to Dec need to become two columns, Month and Sales. Which transform?", T, "Unpivot Columns",
           "Unpivot turns many columns into Attribute / Value pairs.")
    m.pick("You want total Sales per Zone in Power Query. Which transform?", T, "Group By",
           "Group By summarises rows, like a small PivotTable.")
    m.pick("'Abuja, FCT' must become City and State. Which transform?", T, "Split Column",
           "Split Column › By Delimiter › comma.")
    m.pick("Add a column that says 'Loss' when Profit < 0, otherwise 'Profit'. Which transform?", T, "Conditional Column",
           "Add Column › Conditional Column builds an if-then rule without writing code.")
    m.mc("European numbers such as 2400,0 must be read as numbers. Which option?",
         ["Change Type › Using Locale", "Remove Columns", "Sort Ascending", "Keep Top Rows"], "A",
         "Change Type › Using Locale lets you choose the country's number format.")
    m.tf("Renaming a step in Applied Steps changes the data.", False, "It only changes the label; the data is the same.")
    m.part("Part B · Try it in Power Query")
    m.num(f"Load {f26} › SuperMart_Sales_2025_DIRTY.csv into Power Query. Use Home › Remove Rows › Remove Blank Rows. "
          "How many rows are left?", None, len(nonblank), f"{len(raw):,} − {n_blank} blank rows = {len(nonblank):,}.",
          do=True)
    m.num("Now select all columns and use Home › Remove Rows › Remove Duplicates. How many rows are left?", None, n_after,
          f"The exact repeats are removed: {n_after:,} rows, the same as the till system's control total.", do=True)
    m.num("Select the Zone column, use Transform › Format › Trim, then Group By Zone (Count Rows). How many zones are there?",
          None, 4, "After trimming the extra spaces there are 4 zones (without Trim there are 12 'different' zones).",
          do=True)

    # ================================================================== 2.7
    m = m27
    m.part("Excel or Power Query?")
    m.pick("Fixing one typo in a file you will never receive again.", ["Excel", "Power Query"], "Excel",
           "A one-off manual fix is quickest in Excel.")
    m.pick("Cleaning the same monthly file every month.", ["Excel", "Power Query"], "Power Query",
           "Record the steps once and refresh every month.")
    m.pick("You must be able to show an auditor every step that changed the data.", ["Excel", "Power Query"],
           "Power Query", "Applied Steps is a written record of every change.")
    m.pick("Combining 12 monthly files from one folder.", ["Excel", "Power Query"], "Power Query",
           "The Folder connector does it in one query.")
    m.yn("Is manual cleaning in Excel easy to repeat exactly the same way next month?", "No",
         "Manual steps are easy to forget or do differently: that is the reproducibility problem.")
    m.tf("Power Query and Excel cleaning can give the same clean result; the difference is how repeatable it is.", True,
         "Same answer, different effort and audit trail.")

    # ================================================================== 2.8A
    m = m28a
    sw = int(st.loc[st.Zone == "South-West", "Sales_2025"].sum())
    m.example("conditional totals on tStores",
              ["SUMIF(range, criteria, sum_range) adds the sum_range where range meets the criteria.",
               "COUNTIF(range, criteria) counts the cells that meet the criteria."],
              [("Sales of the South-West stores", '=SUMIF(tStores[Zone],"South-West",tStores[Sales_2025])',
                '=SUMIF(tStores[Zone],"South-West",tStores[Sales_2025])', sw, "naira"),
               ("Number of South-West stores", '=COUNTIF(tStores[Zone],"South-West")',
                '=COUNTIF(tStores[Zone],"South-West")', int((st.Zone == "South-West").sum()), "int")])
    m.part("Part A · Totals, counts and averages")
    m.num("How many stores are in the South-East zone? (COUNTIF on tStores[Zone])", '=COUNTIF(tStores[Zone],"South-East")',
          int((st.Zone == "South-East").sum()), '=COUNTIF(tStores[Zone],"South-East")')
    m.num("What are the total 2025 sales of the South-South stores? (SUMIF)",
          '=SUMIF(tStores[Zone],"South-South",tStores[Sales_2025])', int(st.loc[st.Zone == "South-South", "Sales_2025"].sum()),
          '=SUMIF(tStores[Zone],"South-South",tStores[Sales_2025])', fmt="naira")
    m.num("How many stores have Sales_2025 above ₦60,000,000? (COUNTIF with \">60000000\")",
          '=COUNTIF(tStores[Sales_2025],">60000000")', int((st.Sales_2025 > 60_000_000).sum()),
          '=COUNTIF(tStores[Sales_2025],">60000000")')
    m.num("How many FY2025 sales lines (tSales25) are Technology lines in the South-West zone? (COUNTIFS)",
          '=COUNTIFS(tSales25[Category],"Technology",tSales25[Zone],"South-West")',
          int(((s25.Category == "Technology") & (s25.Zone == "South-West")).sum()),
          "COUNTIFS needs every condition to be true.")
    m.num("What is the median Sales_2025 of the 12 stores? (MEDIAN)", "=MEDIAN(tStores[Sales_2025])",
          float(st.Sales_2025.median()), "The middle value: the average of the 6th and 7th largest.", fmt="naira")
    m.num("What is the second-largest Sales_2025? (LARGE with k = 2)", "=LARGE(tStores[Sales_2025],2)",
          int(st.Sales_2025.nlargest(2).iloc[-1]), "=LARGE(tStores[Sales_2025],2): Port Harcourt - GRA.", fmt="naira")
    m.part("Part B · Rounding, logic and text")
    m.num("What does =ROUND(1234.567,1) return?", "=ROUND(1234.567,1)", 1234.6, "Rounded to 1 decimal place.", fmt="dec1")
    m.num("What does =ABS(-8838505) return?", "=ABS(-8838505)", 8838505, "ABS removes the minus sign.")
    m.short("Ibadan - Ring Road (row 4 of Data_Stores): what does "
            '=IF(Data_Stores!J4>=Data_Stores!I4,"Met","Missed") return?', "Missed",
            "Sales ₦32.9m is below the target of ₦37.0m, so the IF returns 'Missed'.")
    m.short('What does =IFERROR(10/0,"n/a") return?', "n/a", "10/0 is an error (#DIV/0!), so IFERROR shows 'n/a'.")
    m.yn('Does =AND(5>3, 2>4) return TRUE?', "No", "AND needs every test to be TRUE; 2>4 is FALSE.")
    m.short('What does =TEXT(0.256,"0%") return?', "26%", "TEXT formats a number as text: 0.256 shown as 26%.")

    # ================================================================== 2.8B
    m = m28b
    m.part("Part A · Text functions")
    m.short('What does =LEFT("ST07",2) return?', "ST", "The first 2 characters.")
    m.short('What does =RIGHT("SM-2025-119438",6) return?', "119438", "The last 6 characters.")
    m.num('What does =LEN("Lagos - Ikeja") return?', '=LEN("Lagos - Ikeja")', 13,
          "LEN counts every character, including the spaces and the dash: 13.")
    m.short('What does =UPPER("uyo - ewet") return?', "UYO - EWET", "UPPER turns every letter into a capital.")
    m.short('What does =TEXTBEFORE("Abuja - Wuse"," - ") return?', "Abuja", "Everything before the first ' - '.")
    m.short('What does =SUBSTITUTE("Office-Supplies","-"," ") return?', "Office Supplies", "Each '-' becomes a space.")
    m.part("Part B · Dates, lookups and dynamic arrays")
    m.num("What does =DATE(2025,12,31)-DATE(2025,1,1) return?", "=DATE(2025,12,31)-DATE(2025,1,1)", 364,
          "Subtracting two dates gives the number of days between them.")
    m.num("In which year did store ST01 open? (=YEAR of the Opened date in row 2 of Data_Stores)", "=YEAR(Data_Stores!F2)",
          int(st.Opened.iloc[0].year), "=YEAR(Data_Stores!F2). Lagos - Ikeja opened on 2015-03-02.")
    m.short("Who manages store ST09? Use =XLOOKUP(\"ST09\",tStores[Store_ID],tStores[Store_Manager]).",
            st.loc[st.Store_ID == "ST09", "Store_Manager"].iloc[0], "XLOOKUP finds ST09 and returns the manager on the same row.")
    m.short("Which zone is store ST11 in? Use VLOOKUP(\"ST11\",tStores,5,FALSE) (Zone is the 5th column).",
            st.loc[st.Store_ID == "ST11", "Zone"].iloc[0], "VLOOKUP looks in the first column and returns column 5.")
    m.num("How many different zones are there? (=COUNTA(UNIQUE(tStores[Zone])))", "=COUNTA(UNIQUE(tStores[Zone]))", 4,
          "UNIQUE lists each zone once; COUNTA counts them.", dynamic=True)
    m.num('How many stores does =FILTER(tStores[Store_Name],tStores[Zone]="North-Central") return? '
          "(Count the spilled rows, or wrap it in ROWS.)", '=ROWS(FILTER(tStores[Store_Name],tStores[Zone]="North-Central"))',
          int((st.Zone == "North-Central").sum()), "FILTER returns the 3 North-Central stores.", dynamic=True)

    # ================================================================== 2.9
    m = m29
    m.part("M language basics")
    m.yn("Is the M language case-sensitive (Text.Upper works, text.upper does not)?", "Yes",
         "M is case-sensitive: function names and column names must match exactly.")
    m.mc("Every M query is built from which structure?", ["IF ... THEN", "let ... in", "SELECT ... FROM", "BEGIN ... END"],
         "B", "let lists the steps; in names the step that is returned.")
    m.short('What does Text.Upper("lagos") return in M?', "LAGOS", "Text.Upper capitalises every letter.")
    m.num("What does Number.Round(2.567, 1) return in M?", None, 2.6, "Rounded to 1 decimal place.", fmt="dec1")
    m.mc("Where can you see and edit the whole M script of a query?",
         ["Advanced Editor", "Name Box", "Status Bar", "Page Layout"], "A", "Home › Advanced Editor.")
    m.mc("Which M function makes the first row the column headers?",
         ["Table.PromoteHeaders", "Table.RemoveRows", "Text.Trim", "List.Sum"], "A", "Table.PromoteHeaders.")
    m.mc("In M, how do you start a comment that runs to the end of the line?", ["#", "//", "--", "'"], "B",
         "// starts a single-line comment; /* ... */ wraps a longer one.")

    # ================================================================== 2.10
    m = m210
    z = sm.groupby("Zone").Sales.sum()
    m.example("a first PivotTable",
              ["Click any cell in tSM, then Insert › PivotTable › From Table/Range › New Worksheet › OK.",
               "Drag Zone to Rows and Sales to Values. The Values area shows 'Sum of Sales'.",
               "The same South-West figure with a formula: SUMIFS adds Sales where Zone is South-West."],
              [("South-West sales, 2022-2025", '=SUMIFS(tSM[Sales],tSM[Zone],"South-West")',
                '=SUMIFS(tSM[Sales],tSM[Zone],"South-West")', int(z["South-West"]), "naira")])
    m.part("Part A · PivotTable ideas")
    m.mc("Which are the four PivotTable areas?", ["Rows, Columns, Values, Filters", "Top, Bottom, Left, Right",
                                                 "Sum, Count, Average, Max", "Sheet, Table, Chart, Slicer"], "A",
         "Rows and Columns group; Values aggregate; Filters limit.")
    m.mc("When a numeric field is dragged to Values, what does Excel use by default?", ["Count", "Sum", "Average", "Max"],
         "B", "Sum, unless the column contains blanks or text, then Count. Always check the label.")
    m.yn("After the source data changes, does the PivotTable update by itself?", "No",
         "Right-click › Refresh (or Data › Refresh All).")
    m.mc("Which tool gives clickable buttons that filter a PivotTable?", ["Slicer", "Sparkline", "Freeze Panes", "Name Box"],
         "A", "PivotTable Analyze › Insert Slicer.")
    m.part("Part B · Build it on tSM")
    m.do("Build the pivot: Zone in Rows, Sum of Sales in Values. Which zone has the highest total sales?",
         z.idxmax(), f"{z.idxmax()}: ₦{z.max():,.0f}.")
    m.num("What is the Sales total for South-East?", '=SUMIFS(tSM[Sales],tSM[Zone],"South-East")', int(z["South-East"]),
          "Read it from the pivot (or check with SUMIFS).", fmt="naira", do=True)
    m.num("Move Year to the Filters area and choose 2025. What is the Grand Total?", '=SUMIFS(tSM[Sales],tSM[Year],2025)',
          int(sm.loc[sm.Year == 2025, "Sales"].sum()), "FY2025 sales: the same ₦733,215,255 as the till system.",
          fmt="naira", do=True)
    m.num("Clear the filter. Change the value to Show Values As › % of Grand Total. What share is South-West? "
          "(1 decimal place)", '=SUMIFS(tSM[Sales],tSM[Zone],"South-West")/SUM(tSM[Sales])',
          float(z["South-West"] / sm.Sales.sum()), "South-West sales ÷ all sales.", fmt="pct", do=True)
    m.num("Put Category in Rows and Count of Row_ID in Values. How many Furniture lines are there?",
          '=COUNTIFS(tSM[Category],"Furniture")', int((sm.Category == "Furniture").sum()),
          "Count of Row_ID counts the lines.", do=True)

    # ================================================================== 2.11
    m = m211
    tech = sm[sm.Category == "Technology"].groupby("Zone").Sales.sum()
    m.part("PivotCharts")
    m.mc("What is a PivotChart?", ["A picture pasted from the web", "A chart linked to a PivotTable that changes when the "
                                   "pivot changes", "A chart that cannot be filtered", "A table of numbers"], "B",
         "PivotChart and PivotTable stay in sync both ways.")
    m.yn("If you filter the PivotChart with its field buttons, does the linked PivotTable change too?", "Yes",
         "They share the same fields and filters.")
    m.mc("Which chart type is best for comparing sales across four zones?", ["Pie with 20 slices", "Clustered column",
                                                                           "Scatter", "Radar"], "B",
         "Columns make it easy to compare a few categories.")
    m.mc("Which chart type is best for a trend over 12 months?", ["Line", "Pie", "Doughnut", "Treemap"], "A",
         "A line shows change over time.")
    m.do("Build a PivotChart from the Zone pivot (Sum of Sales). Filter Category to Technology. Which zone has the tallest "
         "column?", tech.idxmax(), f"{tech.idxmax()}: ₦{tech.max():,.0f} of Technology sales.")
    m.tf("Changing the chart type (e.g. column to bar) changes the numbers in the PivotTable.", False,
         "Only the picture changes; the data is the same.")
