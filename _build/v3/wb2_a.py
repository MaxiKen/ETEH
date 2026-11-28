"""Week 2 workbook, modules 2.1 - 2.4 (edition 3): external files, Power Query basics, connecting, web and PDF."""
from __future__ import annotations

import pandas as pd

from w1x import ModuleSheet
from wb2_common import N, code, lit_grid, serial
from w2_fx import DF

DATA = "01_Section_Examples\\Week2_Data_Files\\"


# =====================================================================================================
def m21(b, D):
    S, DR, A, FILES = D.S, D.DR, D.A, D.FILES
    m = ModuleSheet(b, "2.1 External Formats", "2.1  Understanding External Data Sources",
                    "Tutor walk-through after Module 2.1: file formats, importing workbooks, CSV and TXT, delimiters and "
                    "qualifiers, safe import, locale and encoding", "2.1.1 - 2.1.8")
    m.dataset([("tFiles", "On this sheet, below", "The nine Week 2 practice files: format, encoding, data rows and what each holds. "
                                                  "The files themselves are in " + DATA),
               ("tDirty25", "Sheet Data_Dirty25 (starts in A1)", "SuperMart_Sales_2025_DIRTY.csv as Excel opens it: 3,387 rows, 21 columns, "
                                                                 "defects left in on purpose."),
               ("tAdm23", "Sheet Data_Adm23 (starts in A1)", "Mercyland_Episodes_2023_pipe.txt imported: 12,626 admissions (14 of the 22 columns).")])
    m.h("tFiles")
    m.s.table(FILES, "tFiles", fmts={"Rows": "int"}, widths={"File": 40, "Format": 18, "Encoding": 15, "What it holds": 44})

    # ---------------------------------------------------------------- 2.1.1-2
    csv_rows = int(FILES[FILES.Format.str.startswith("CSV")].Rows.sum())
    m.topic("2.1.1-2", "What external data is, and the common file formats",
            "External data lives outside your workbook: a CSV export, a text file from another system, a web page, a PDF or another "
            "workbook. Each format stores rows differently, so each needs the right import route.")
    m.steps(["Open the folder " + DATA + " in File Explorer and switch the view to Details.",
             "Match each file to a row of tFiles: the extension tells you the format (.csv, .txt, .xlsx, .json, .html, .pdf).",
             "Note the Encoding column: UTF-8 keeps the ₦ sign; Windows-1252 is an older Western encoding.",
             "Sort tFiles by Rows (largest to smallest) to see which files are the big ones."])
    m.worked([("Practice files listed", "=ROWS(tFiles)", len(FILES), "int", "One row per file or folder."),
              ("Data rows in all CSV files", '=SUMIFS(tFiles[Rows],tFiles[Format],"CSV*")', csv_rows, "int",
               "The * wildcard matches CSV (comma) and CSV (semicolon)."),
              ("Largest file (rows)", "=MAX(tFiles[Rows])", int(FILES.Rows.max()), "int", ""),
              ("Which file is it?", "=INDEX(tFiles[File],MATCH(MAX(tFiles[Rows]),tFiles[Rows],0))",
               FILES.loc[FILES.Rows.idxmax(), "File"], None, "INDEX-MATCH returns the name of the largest file.")])
    m.turn()
    m.ex("How many data rows are in the two Mercyland files together? (Hint: SUMIFS with \"Mercyland*\".)",
         '=SUMIFS(tFiles[Rows],tFiles[File],"Mercyland*")', int(FILES[FILES.File.str.startswith("Mercyland")].Rows.sum()), "int",
         short="Rows in the Mercyland files")
    m.ex_text("Which format stores data as plain text with one record per line and commas between fields?", "CSV", alts=(".csv",))
    m.ex_text("Which format can hold several sheets, data types and formulas?", "xlsx", alts=(".xlsx", "Excel workbook", "workbook"))

    # ---------------------------------------------------------------- 2.1.3
    m.topic("2.1.3", "Importing from .xlsx and .xls workbooks",
            "Data › Get Data › From File › From Excel Workbook connects to another workbook. The Navigator lists every sheet and table; "
            "you tick the ones you need and either Load them or Transform Data first.")
    m.do([("Data › Get Data › From File › From Excel Workbook › pick SuperMart_Online_RAW_Export.xlsx › Import.",
           "The Navigator lists 5 sheets: Transactions, Sellers, Products, Deliveries, Customers."),
          ("Click Transactions once (do not tick it yet) and read the preview.", "15 columns, from transaction_id to order_date."),
          ("Tick Transactions › Transform Data.", "The Power Query Editor opens with the query 'Transactions'."),
          ("Home › Close & Load.", "A new sheet with an 11,000-row table called Transactions.")])
    m.turn()
    m.ex_text("How many rows does the Customers sheet of SuperMart_Online_RAW_Export.xlsx hold?", "2500", alts=("2,500",))
    m.ex_text("Which sheet of the workbook has the most rows?", "Transactions")
    m.ex_text("How many columns does the Deliveries sheet have?", "13")

    # ---------------------------------------------------------------- 2.1.4
    m.topic("2.1.4", "Importing from CSV files",
            "From Text/CSV reads a comma-separated file. Check three settings in the preview: File Origin (encoding), Delimiter and "
            "Data Type Detection, then Transform Data to review the columns.")
    m.steps(["Data › From Text/CSV › SuperMart_Sales_2025_DIRTY.csv › Import.",
             "File Origin: 65001: Unicode (UTF-8), so the ₦ in the text Sales values displays correctly.",
             "Delimiter: Comma. Data Type Detection: Based on first 200 rows.",
             "Transform Data. The status bar shows the column count; Close & Load shows the row count."])
    m.worked([("Rows Excel received", "=ROWS(tDirty25)", len(DR), "int", "The data rows in the file, header excluded."),
              ("Columns", "=COLUMNS(tDirty25)", DR.shape[1], "int", "Includes the empty Column1 (fixed in 2.5.10)."),
              ("Sales cells that arrived as text", "=SUMPRODUCT(--ISTEXT(tDirty25[Sales]))", DF["sales_text"], "int",
               "Typed with ₦ and commas, so Excel kept them as text.")])
    m.turn()
    m.ex("How many rows does SuperMart_Sales_2025_DIRTY.csv give you after import?", "=ROWS(tDirty25)", len(DR), "int",
         short="Rows in the dirty CSV")
    m.ex("How many Discount cells arrived as text (e.g. '20%')?", "=SUMPRODUCT(--ISTEXT(tDirty25[Discount]))", DF["disc_text"], "int",
         short="Discount cells stored as text")
    m.ex_text("Which File Origin keeps the ₦ sign readable in this file? Type the code number.", "65001",
              alts=("65001: unicode (utf-8)", "utf-8", "utf8"))

    # ---------------------------------------------------------------- 2.1.5
    m.topic("2.1.5", "Importing from TXT files",
            "A .txt file is read the same way as a CSV, but its delimiter is often something else: a pipe |, a tab or a semicolon. "
            "Choose Delimiter › Custom and type the character.")
    m.steps(["Data › From Text/CSV › Mercyland_Episodes_2023_pipe.txt.",
             "Delimiter › --Custom-- › type | (Shift + the key above Enter).",
             "Check the preview: 22 columns, one admission per row.",
             "Transform Data, keep the 14 columns used on Data_Adm23, then Close & Load."])
    m.worked([("Admissions imported", "=ROWS(tAdm23)", len(A), "int", "Matches the 12,626 records in the file."),
              ("Columns kept on Data_Adm23", "=COLUMNS(tAdm23)", A.shape[1], "int", "8 of the 22 were not needed for Week 2."),
              ("Admissions in Internal Medicine", '=COUNTIF(tAdm23[Department],"Internal Medicine")',
               int((A.Department == "Internal Medicine").sum()), "int", "A first question answered from the imported file.")])
    m.turn()
    m.ex("How many 2023 admissions were Emergency admissions?", '=COUNTIF(tAdm23[Admission_Route],"Emergency")',
         int((A.Admission_Route == "Emergency").sum()), "int", short="Emergency admissions 2023")
    m.ex_text("Which character separates the fields in Mercyland_Episodes_2023_pipe.txt?", "|", alts=("pipe",))

    # ---------------------------------------------------------------- 2.1.6
    line = ('6683,SM-2025-107503,2025-01-01,2025-01-06,Standard Class,SMC-0294,Chijioke Obasi,Consumer,ST06,Uyo - Ewet,Uyo,'
            'Akwa Ibom,South-South,FUR-FU-10003878,Furniture,Furnishings,"Linden 10"" Round Wall Clock, Black",48896,4,0.2,8557')
    m.topic("2.1.6", "Delimiters and text qualifiers",
            "The delimiter separates fields. A text qualifier (usually the double quote) wraps a field that contains the delimiter, so "
            "the comma in \"Linden 10\"\" Round Wall Clock, Black\" stays part of the name. Two quotes inside a quoted field mean one quote.")
    m.h("A real line from 2025-01_SuperMart_Sales.csv (Row_ID 6683)")
    inp = m.s.inputs([("Line as stored in the file", line, "text", "Paste any other line here to test it")])
    cell = inp["Line as stored in the file"]
    naive = line.count(",") + 1
    w = m.worked([("Commas in the line", f'=LEN({cell})-LEN(SUBSTITUTE({cell},",",""))', line.count(","), "int", ""),
                  ("Fields if every comma split", f'=LEN({cell})-LEN(SUBSTITUTE({cell},",",""))+1', naive, "int",
                   "One field too many: the comma inside the quotes."),
                  ("Commas inside the quoted name", f'=LEN(MID({cell},FIND("""",{cell}),FIND(""",",{cell})-FIND("""",{cell})))'
                                                    f'-LEN(SUBSTITUTE(MID({cell},FIND("""",{cell}),FIND(""",",{cell})-FIND("""",{cell})),",",""))',
                   1, "int", "From the first quote to the closing quote.")])
    m.worked([("Real fields (qualifier respected)", f"={w['Commas in the line']}-{w['Commas inside the quoted name']}+1", naive - 1, "int",
               "21 columns, as in the file header.")], label="Worked solution (continued)")
    m.turn()
    m.ex_text("Without a text qualifier, how many columns would this line split into?", str(naive))
    m.ex_text("Inside a quoted field, how is one double-quote character written?", '""', alts=("two quotes", "2 quotes", "two double quotes"))

    # ---------------------------------------------------------------- 2.1.7
    m.topic("2.1.7", "Importing data into Excel safely",
            "Double-clicking a CSV lets Excel guess every type: leading zeros vanish from phone numbers, long IDs lose digits after the "
            "15th and codes such as 1-2 become dates. Import through Data › From Text/CSV and set those columns to Text.")
    m.steps(["Never double-click the CSV. Use Data › From Text/CSV.",
             "Transform Data. Select phone, account and ID columns › Data Type › Text.",
             "Only then set the number and date columns.",
             "Compare a few values with the file opened in Notepad."])
    m.worked([("Phone as Excel stores it after a careless open", "=7053939110", 7053939110, "0", "The leading 0 is gone."),
              ("Digits left", "=LEN(7053939110)", 10, "int", "Nigerian mobile numbers have 11."),
              ("Repaired for display", '=TEXT(7053939110,"00000000000")', "07053939110", None, "Pads back to 11 digits, as text.")])
    m.note("Excel keeps only 15 significant digits in a number: a 19-digit ID typed as a number loses its last 4 digits for good. "
           "Import such IDs as Text.")
    m.turn()
    m.ex_text("What data type should a phone-number column get when you import it?", "Text")
    m.ex("A seller phone number arrived as 8031234567. Write a formula that shows it with its leading zero (11 digits), then type "
         "the number of characters in the result.", '=LEN(TEXT(8031234567,"00000000000"))', 11, "int", short="Length after padding")

    # ---------------------------------------------------------------- 2.1.8
    m.topic("2.1.8", "Locale and encoding traps (UTF-8, Windows-1252)",
            "Encoding decides which bytes mean which characters: open a UTF-8 file as Windows-1252 and ₦ becomes â‚¦. Locale decides how "
            "numbers and dates are read: in France 2400,0 is two thousand four hundred and the thousands separator is a space.")
    m.steps(["Data › From Text/CSV › SuperMart_Store_Targets_EU_semicolon.csv.",
             "File Origin: 1252: Western European (Windows). Delimiter: Semicolon.",
             "Transform Data. Right-click Annual_Target_2025 › Change Type › Using Locale › Decimal Number, French (France).",
             "In a worksheet, NUMBERVALUE does the same job: you tell it the decimal and group separators."])
    m.worked([("Floor area '2400,0'", '=NUMBERVALUE("2400,0",",",".")', 2400, "#,##0.0", "Comma is the decimal mark."),
              ("Ikeja target '142 000 000,00'", '=NUMBERVALUE("142 000 000,00",","," ")', 142000000, N, "Space is the group separator."),
              ("Checked against tStores", '=XLOOKUP("ST01",tStores[Store_ID],tStores[Target_2025])', int(D.ST.Target_2025[0]), N,
               "Same value: the locale-aware reading is right.")])
    m.turn()
    m.ex("Convert the Port Harcourt - GRA target as written in the EU file, '97 000 000,00', to a number.",
         '=NUMBERVALUE("97 000 000,00",","," ")', 97000000, N, short="NUMBERVALUE of the ST04 target")
    m.ex_text("A UTF-8 file opened as Windows-1252 shows â‚¦ instead of which character?", "₦", alts=("naira", "naira sign"))
    m.finish()


# =====================================================================================================
def m22(b, D):
    S, DR, ST = D.S, D.DR, D.ST
    m = ModuleSheet(b, "2.2 Power Query Intro", "2.2  Introduction to Power Query",
                    "Tutor walk-through after Module 2.2: what Power Query is, ETL, the editor, applied steps, profiling, managing and "
                    "loading queries", "2.2.1 - 2.2.10")
    m.dataset([("tDirty25", "Sheet Data_Dirty25 (starts in A1)", "The raw file as received. Its counts are what Power Query should show "
                                                                 "before cleaning."),
               ("tSales25", "Sheet Data_Sales25 (starts in A1)", "The same lines after cleaning (3,312 rows). Its counts are what your query "
                                                                 "should reach."),
               ("tStores", "Sheet Data_Stores (starts in A1)", "The 12 SuperMart stores; used here to open the editor From Table/Range.")])

    # ---------------------------------------------------------------- 2.2.1-3
    m.topic("2.2.1-3", "What Power Query is, why automate, and how it differs from cleaning in cells",
            "Power Query (Data › Get & Transform) connects to data, records every cleaning action as a step and replays all the steps "
            "when you click Refresh. The source file is never changed; the cleaned result is loaded where you choose.")
    m.steps(["Compare the two data sheets: Data_Dirty25 is what arrives, Data_Sales25 is what the report needs.",
             "Count what has to happen every month to get from one to the other (worked solution).",
             "Ask: will the same file arrive again next month? If yes, the steps belong in Power Query."])
    gap = len(DR) - len(S)
    m.worked([("Rows received", "=ROWS(tDirty25)", len(DR), "int", ""),
              ("Rows after cleaning", "=ROWS(tSales25)", len(S), "int", ""),
              ("Rows removed every month", "=ROWS(tDirty25)-ROWS(tSales25)", gap, "int", "15 blank rows + 60 duplicates."),
              ("Text values to convert", "=SUMPRODUCT(--ISTEXT(tDirty25[Sales]))+SUMPRODUCT(--ISTEXT(tDirty25[Discount]))",
               DF["sales_text"] + DF["disc_text"], "int", "By hand each month, or once as a recorded step.")])
    m.turn()
    m.ex_text("True or false: Power Query changes the original source file.", "False", alts=("f", "no"))
    m.ex_text("Next month's file has arrived in the same place. Which one command reruns every cleaning step?", "Refresh",
              alts=("refresh all", "data > refresh all"))
    m.ex("How many rows does the cleaning remove from this file? (Use the two tables.)", "=ROWS(tDirty25)-ROWS(tSales25)", gap, "int",
         short="Rows removed by cleaning")

    # ---------------------------------------------------------------- 2.2.4
    m.topic("2.2.4", "The ETL cycle in Power Query",
            "Extract (connect and read the source), Transform (clean and reshape in recorded steps), Load (send the result to a table, a "
            "PivotTable or the Data Model). Refresh runs the whole cycle again.")
    lit_grid(m, [["Source = Csv.Document(File.Contents(…))", "Extract"], ["Removed Duplicates", "Transform"],
                 ["Changed Type with Locale", "Transform"], ["Close & Load To… › Table", "Load"]],
             header=["Step in the dirty-file query", "Stage"])
    m.turn()
    m.ex_text("'Promoted Headers' (first row becomes the column names): E, T or L?", "T", alts=("transform",))
    m.ex_text("Close & Load To… › PivotTable Report: E, T or L?", "L", alts=("load",))
    m.ex_text("Navigator › tick the Transactions sheet: E, T or L?", "E", alts=("extract",))

    # ---------------------------------------------------------------- 2.2.5-6
    m.topic("2.2.5-6", "Opening the Power Query Editor, and its parts",
            "Open the editor from Get Data (a new source) or Data › From Table/Range (a table already in this workbook). The editor has "
            "the ribbon, the Queries pane (left), the formula bar, the data preview, Query Settings with Applied Steps (right) and the "
            "status bar (column and row counts).")
    m.do([("Click any cell on Data_Stores › Data › From Table/Range.", "The editor opens with a query called tStores."),
          ("Find the five parts: ribbon, Queries pane, formula bar, preview, Query Settings.", "If the formula bar is hidden: View › Formula Bar."),
          ("Read the status bar at the bottom left.", "10 COLUMNS, 12 ROWS."),
          ("Home › Close & Load To… › Only Create Connection.", "The query is saved without a second copy of the table.")])
    m.turn()
    m.ex("How many columns does the editor show for tStores?", "=COLUMNS(tStores)", ST.shape[1], "int", short="Columns of tStores")
    m.ex_text("Which pane on the right lists Applied Steps?", "Query Settings", alts=("query settings pane",))

    # ---------------------------------------------------------------- 2.2.7
    m.topic("2.2.7", "Queries, Applied Steps and Query Settings",
            "Each action you take becomes one Applied Step. Click a step to see the data as it was at that point; the gear opens its "
            "settings; the X deletes it. Steps run top to bottom every refresh.")
    m.do([("Data › From Text/CSV › SuperMart_Sales_2025_DIRTY.csv › Transform Data.", "Applied Steps: Source, Promoted Headers, Changed Type."),
          ("Click Source, then Promoted Headers.", "At Source the first row holds the column names; at Promoted Headers they are headers."),
          ("Right-click Changed Type › Delete.", "Columns go back to text (ABC icon). Undo with a new Changed Type if you need it."),
          ("Right-click the query in the Queries pane › Rename › qSales_Raw.", "Clear names make the Applied Steps list readable.")])
    m.turn()
    m.ex_text("What is the name of the step that turns the first row into column names?", "Promoted Headers")
    m.ex("How many data rows does the status bar show after Promoted Headers?", "=ROWS(tDirty25)", len(DR), "int", short="Rows after Promoted Headers")

    # ---------------------------------------------------------------- 2.2.8
    qb = sum(1 for v in DR.Quantity if v is None)
    m.topic("2.2.8", "Data preview: column quality, distribution and profile",
            "View › Column quality shows % valid, error and empty for each column; Column distribution shows distinct and unique counts; "
            "Column profile shows full statistics. By default they use the first 1,000 rows: switch the status bar to 'Column profiling "
            "based on entire data set'.")
    m.steps(["In qSales_Raw: View › tick Column quality, Column distribution and Column profile.",
             "Click 'Column profiling based on top 1000 rows' on the status bar › choose entire data set.",
             "Click the Quantity column and read Empty in the Column statistics.",
             "Click Zone and read the Distinct count in the value distribution."])
    m.worked([("Empty Quantity cells (whole file)", "=COUNTBLANK(tDirty25[Quantity])", qb, "int", "15 blank rows + 31 real lines."),
              ("As a share of rows", "=COUNTBLANK(tDirty25[Quantity])/ROWS(tDirty25)", qb / len(DR), "pct", "Column quality shows about 1%."),
              ("Distinct Zone values", '=ROWS(UNIQUE(FILTER(tDirty25[Zone],tDirty25[Zone]<>"")))', DF["zone_distinct"], "int",
               "Spaces make 12 'zones' out of 4.")])
    m.turn()
    m.ex("How many Quantity values are empty when profiling uses the entire data set?", "=COUNTBLANK(tDirty25[Quantity])", qb, "int",
         short="Empty Quantity (entire data set)")
    m.ex("How many distinct Zone values does Column distribution report (blank rows excluded)?",
         '=ROWS(UNIQUE(FILTER(tDirty25[Zone],tDirty25[Zone]<>"")))', DF["zone_distinct"], "int", short="Distinct Zone values")
    m.note("Excel's UNIQUE ignores case but Power Query does not: Category has 12 spellings in Power Query and 8 with UNIQUE. "
           "Zone differs only by spaces, so both give 12.")

    # ---------------------------------------------------------------- 2.2.9
    m.topic("2.2.9", "Managing queries: groups, Duplicate and Reference",
            "Duplicate copies every step into an independent query. Reference starts a new query from another query's result, so a fix "
            "in the cleaning query flows into every query that references it. Groups are folders in the Queries pane.")
    m.do([("Right-click qSales_Raw › Reference. Rename the new query qSales_Clean.", "qSales_Clean has one step: Source = qSales_Raw."),
          ("Right-click qSales_Clean › Reference again › rename qZoneSummary.", "Three queries, each starting from the one above."),
          ("Select all three › right-click › Move To Group › New Group › 'SuperMart'.", "The Queries pane shows a SuperMart folder."),
          ("Right-click qSales_Raw › Enable load: untick it.", "Staging query: it feeds others but loads nothing.")])
    m.turn()
    m.ex_text("Three summary queries must all pick up any fix made to the cleaning query. Duplicate or Reference?", "Reference")
    m.ex_text("You want a one-off variation that must NOT change when the original changes. Duplicate or Reference?", "Duplicate")

    # ---------------------------------------------------------------- 2.2.10
    m.topic("2.2.10", "Loading data into Excel and the Data Model",
            "Home › Close & Load To… offers Table, PivotTable Report, PivotChart and Only Create Connection, plus 'Add this data to the "
            "Data Model'. A sheet holds at most 1,048,576 rows; the Data Model has no such limit and enables Distinct Count.")
    m.do([("Right-click qSales_Clean › Load To… › Table › New worksheet.", "A table of the query's rows on a new sheet."),
          ("Right-click qZoneSummary › Load To… › PivotTable Report, tick 'Add this data to the Data Model'.", "An empty pivot whose field list "
                                                                                                             "shows the query."),
          ("Data › Queries & Connections.", "The pane lists each query with its row count or 'Connection only'.")])
    m.turn()
    m.ex_text("Which Load To option saves the query but puts no rows on a sheet?", "Only Create Connection",
              alts=("connection only", "create connection"))
    m.ex_text("How many rows can one worksheet hold? Type the number.", "1048576", alts=("1,048,576",))
    m.finish()


# =====================================================================================================
def m23(b, D):
    S, A, MF = D.S, D.A, D.MF
    m = ModuleSheet(b, "2.3 Connect & Combine", "2.3  Connecting and Importing with Power Query",
                    "Tutor walk-through after Module 2.3: connecting to workbooks, CSV and text, several sheets, a folder of files, "
                    "paths, parameters and refresh", "2.3.1 - 2.3.8")
    m.dataset([("tMonthFiles", "On this sheet, below", "The 12 monthly files in SuperMart_Monthly_Sales_2025: rows and sales in each. "
                                                       "Your combined query must reproduce these."),
               ("tSales25", "Sheet Data_Sales25 (starts in A1)", "The clean FY2025 lines: the combined monthly files must add up to it."),
               ("tAdm23", "Sheet Data_Adm23 (starts in A1)", "The Mercyland pipe file, connected in 2.3.3.")])
    m.h("tMonthFiles")
    m.s.table(MF, "tMonthFiles", fmts={"Rows": "int", "Sales": N}, widths={"File": 30, "Sales": 15})
    m.h("Folder path used by the M code on this sheet (change it to where you saved the pack)")
    fold = m.s.inputs([("DataFolder", "C:\\ETEH\\ETEH_Bootcamp_Pack\\01_Section_Examples\\Week2_Data_Files\\", "text",
                        "End with a backslash. The Power Query parameter of the same name is set up in 2.3.7.")])
    b.wb.define_name("DataFolder", f"='2.3 Connect & Combine'!{fold['DataFolder']}")

    # ---------------------------------------------------------------- 2.3.1
    m.topic("2.3.1", "Connecting to Excel workbooks",
            "Excel.Workbook lists every sheet, table and named range in a file. Navigating to one item adds a step that picks it by name "
            "and kind, so the query survives the sheet being moved.")
    code(m, '''let
    Source       = Excel.Workbook(File.Contents(DataFolder & "SuperMart_Online_RAW_Export.xlsx"), null, true),
    Transactions = Source{[Item = "Transactions", Kind = "Sheet"]}[Data],
    Promoted     = Table.PromoteHeaders(Transactions, [PromoteAllScalars = true])
in
    Promoted''', "M behind the Navigator (Home › Advanced Editor shows it)")
    m.do([("Data › Get Data › From File › From Excel Workbook › SuperMart_Online_RAW_Export.xlsx.", "The Navigator lists 5 sheets."),
          ("Tick Sellers › Transform Data.", "Source, Navigation, Promoted Headers, Changed Type."),
          ("Click the Source step.", "A table of 5 rows: Name, Data, Item, Kind, Hidden."),
          ("Close & Load To… › Only Create Connection.", "Saved for 2.3.4.")])
    m.turn()
    m.ex_text("How many rows does the Sellers sheet load?", "200")
    m.ex_text("In the Source step, which column holds each sheet's contents?", "Data")

    # ---------------------------------------------------------------- 2.3.2
    m.topic("2.3.2", "Connecting to CSV files",
            "Csv.Document reads a text file with the delimiter, encoding and quote handling you give it. QuoteStyle.Csv honours text "
            "qualifiers, so commas inside quotes stay in the field.")
    code(m, '''let
    Source   = Csv.Document(File.Contents(DataFolder & "SuperMart_Sales_2025_DIRTY.csv"),
                   [Delimiter = ",", Columns = 21, Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true])
in
    Promoted''', "Connect to the dirty sales file")
    m.turn()
    m.ex_text("Which Encoding number in Csv.Document means UTF-8?", "65001")
    m.ex_text("Which QuoteStyle keeps \"Round Wall Clock, Black\" in one field?", "QuoteStyle.Csv", alts=("csv",))
    m.ex("How many rows should the query show after Promoted Headers?", "=ROWS(tDirty25)", len(D.DR), "int", short="Rows after Promoted Headers")

    # ---------------------------------------------------------------- 2.3.3
    m.topic("2.3.3", "Connecting to text files",
            "Text files use the same Csv.Document function with a different delimiter. Fixed-width files have no delimiter: split them "
            "at character positions.")
    code(m, '''let
    Source   = Csv.Document(File.Contents(DataFolder & "Mercyland_Episodes_2023_pipe.txt"), [Delimiter = "|", Columns = 22, Encoding = 65001]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed    = Table.TransformColumnTypes(Promoted, {{"Admit_Date", type date}, {"Discharge_Date", type date},
                   {"Age", Int64.Type}, {"Bill_Amount_NGN", type number}})
in
    Typed''', "Connect to the Mercyland pipe file")
    m.worked([("Admissions loaded", "=ROWS(tAdm23)", len(A), "int", ""),
              ("Total billed in 2023", "=SUM(tAdm23[Bill_Amount_NGN])", float(A.Bill_Amount_NGN.sum()), N, "Only correct once Bill_Amount_NGN is a number."),
              ("Average length of stay (days)", "=AVERAGE(tAdm23[LOS_Days])", float(A.LOS_Days.mean()), "dec2", "")])
    m.turn()
    m.ex("How many 2023 admissions were readmitted within 30 days (Readmitted_30d = Yes)?", '=COUNTIF(tAdm23[Readmitted_30d],"Yes")',
         int((A.Readmitted_30d == "Yes").sum()), "int", short="Readmissions 2023")
    m.ex_text("What Delimiter value goes in Csv.Document for this file?", "|", alts=('"|"', "pipe"))

    # ---------------------------------------------------------------- 2.3.4
    m.topic("2.3.4", "Importing multiple worksheets from one workbook",
            "Tick Select multiple items in the Navigator to bring in several sheets as separate queries. Sheets with the same columns "
            "can be stacked by keeping the Source step's Data column and expanding it.")
    lit_grid(m, [["Transactions", 11000, 15, "seller_id, product_id, customer_id, delivery_id"], ["Sellers", 200, 14, "seller_id"],
                 ["Products", 800, 15, "product_id, seller_id"], ["Deliveries", 10430, 13, "delivery_id, order_id, customer_id"],
                 ["Customers", 2500, 15, "customer_id"]],
             header=["Sheet", "Rows", "Columns", "Key columns (join on these in Week 3)"], merge_last=True)
    m.do([("From Excel Workbook › SuperMart_Online_RAW_Export.xlsx › tick Select multiple items.", "A tick box appears beside every sheet."),
          ("Tick Transactions, Sellers and Products › Transform Data.", "Three queries in the Queries pane."),
          ("Group the three queries as 'Online'.", "Ready to be merged in Week 3.")])
    m.turn()
    m.ex_text("Which column links Transactions to Deliveries?", "delivery_id")
    m.ex_text("How many sheets does SuperMart_Online_RAW_Export.xlsx contain?", "5", alts=("five",))

    # ---------------------------------------------------------------- 2.3.5
    m.topic("2.3.5", "Combining files from a folder",
            "From Folder lists every file in a folder; Combine applies the steps built on one sample file to all of them and stacks the "
            "results. A new month's file dropped into the folder is picked up at the next refresh.")
    m.steps(["Data › Get Data › From File › From Folder › SuperMart_Monthly_Sales_2025 › OK.",
             "Combine › Combine & Transform Data. Pick the first file as the sample › OK.",
             "Keep the Source.Name column: it records which file each row came from.",
             "Check the result against tMonthFiles and against tSales25 (worked solution)."])
    m.worked([("Rows in the 12 files", "=SUM(tMonthFiles[Rows])", int(MF.Rows.sum()), "int", "Expected rows after Combine."),
              ("Sales in the 12 files", "=SUM(tMonthFiles[Sales])", int(MF.Sales.sum()), N, ""),
              ("Clean FY2025 sales (tSales25)", "=SUM(tSales25[Sales])", int(S.Sales.sum()), N, ""),
              ("Difference", "=SUM(tMonthFiles[Sales])-SUM(tSales25[Sales])", int(MF.Sales.sum() - S.Sales.sum()), N, "0: the folder holds "
                                                                                                                       "the whole year.")])
    big = MF.loc[MF.Rows.idxmax()]
    m.turn()
    m.ex("How many rows does the December file add?", "=INDEX(tMonthFiles[Rows],12)", int(MF.Rows.iloc[11]), "int", short="Rows in December")
    m.ex_text("Which file has the most rows?", big.File, formula="=INDEX(tMonthFiles[File],MATCH(MAX(tMonthFiles[Rows]),tMonthFiles[Rows],0))",
              short="File with most rows")
    m.ex_text("Which column added by Combine tells you the file each row came from?", "Source.Name")

    # ---------------------------------------------------------------- 2.3.6-7
    m.topic("2.3.6-7", "File paths, source settings and parameters",
            "A query stores the full path to its file. Move the file and refresh fails with DataSource.Error. Change it once in Data "
            "source settings, or better, keep the folder in a parameter and build every path from it.")
    m.steps(["Home › Manage Parameters › New. Name: DataFolder; Type: Text; Current Value: your folder path (ending in \\).",
             "In each query's Source step, replace the typed folder with DataFolder & \"file name\".",
             "Moving the pack? Change DataFolder once; every query follows.",
             "Data source settings (Home tab) lists every path the workbook uses."])
    m.worked([("Full path to the dirty file", '=DataFolder&"SuperMart_Sales_2025_DIRTY.csv"',
               "C:\\ETEH\\ETEH_Bootcamp_Pack\\01_Section_Examples\\Week2_Data_Files\\SuperMart_Sales_2025_DIRTY.csv", None,
               "Built from the DataFolder cell above, the same way the parameter builds it in M."),
              ("Characters in the path", '=LEN(DataFolder&"SuperMart_Sales_2025_DIRTY.csv")',
               len("C:\\ETEH\\ETEH_Bootcamp_Pack\\01_Section_Examples\\Week2_Data_Files\\SuperMart_Sales_2025_DIRTY.csv"), "int",
               "Long paths are easy to mistype: another reason to type them once.")])
    m.turn()
    m.ex_text("The file was moved and Refresh fails. The error starts with which word?", "DataSource.Error", alts=("datasource",))
    m.ex_text("Which Home-tab command lists and changes every file path in the workbook?", "Data source settings",
              alts=("data source settings",))

    # ---------------------------------------------------------------- 2.3.8
    m.topic("2.3.8", "Refreshing queries after source changes",
            "Refresh All (Ctrl+Alt+F5) reruns every query and PivotTable; Alt+F5 refreshes the selected one. Query Properties can refresh "
            "on open or every n minutes. Check the row count after each refresh.")
    m.do([("Data › Queries & Connections › right-click qSales_Raw › Properties.", "Usage tab: Refresh data when opening the file."),
          ("Add a copy of 2025-12_SuperMart_Sales.csv to the monthly folder as 2026-01_test.csv › Refresh All.",
           "The combined query grows by the December row count. Delete the test file and refresh again."),
          ("Compare the loaded rows with =SUM(tMonthFiles[Rows]).", "Equal again after removing the test file.")])
    m.turn()
    m.ex_text("Which shortcut refreshes every query and PivotTable in the workbook?", "Ctrl+Alt+F5", alts=("ctrl alt f5",))
    m.ex("After the test file is added, how many rows should the combined query show?", "=SUM(tMonthFiles[Rows])+INDEX(tMonthFiles[Rows],12)",
         int(MF.Rows.sum() + MF.Rows.iloc[11]), "int", short="Rows with the test copy added")
    m.finish()


# =====================================================================================================
def m24(b, D):
    S, ST, WEB, BUL = D.S, D.ST, D.WEB, D.BUL
    m = ModuleSheet(b, "2.4 Web & PDF Sources", "2.4  Web and PDF Data Sources (Practice Files)",
                    "Tutor walk-through after Module 2.4: a saved web page, a PDF table, unpivoting, reconciling and audit checks",
                    "2.4.1 - 2.4.5")
    m.dataset([("tWebStores", "On this sheet, below", "Table 1 of SuperMart_Store_Directory.html as From Web returns it: every value is text."),
               ("tBulletin", "On this sheet, below", "The table in SuperMart_Monthly_Sales_Bulletin_2025.pdf: one row per store, ₦ millions by "
                                                     "month, totals row removed."),
               ("tSales25", "Sheet Data_Sales25 (starts in A1)", "The ledger the bulletin is reconciled to.")])

    # ---------------------------------------------------------------- 2.4.1
    m.topic("2.4.1", "Beginner: inspect the store directory web page",
            "From Web reads the HTML tables on a page. A saved page works the same way as a live one: paste its file:/// path. "
            "Web tables arrive as text, so dates and numbers need types (with the right locale).")
    m.steps(["Open SuperMart_Store_Directory.html in a browser: 2 tables (stores, opening hours) and a footnote.",
             "Data › From Web › paste file:///C:/…/SuperMart_Store_Directory.html › OK.",
             "Navigator › Table 1 › Transform Data.",
             "Opened is text such as 02 Mar 2015: Change Type › Using Locale › Date, English (United Kingdom)."])
    m.h("tWebStores (as imported) with the type fixes as calculated columns")
    m.note("Opened_Date rebuilds the date with DATE(year, month, day), which works on any PC. DATEVALUE(\"02 Mar 2015\") also works, "
           "but only where Windows is set to English month names.")
    opened = pd.to_datetime(WEB.Opened, format="%d %b %Y")
    m.s.table(WEB, "tWebStores", widths={"Store": 26, "Manager": 18, "Opened": 12, "Floor area (sq m)": 15},
              calc={"Opened_Date": ('=DATE(RIGHT([@Opened],4),MATCH(MID([@Opened],4,3),{"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep",'
                                    '"Oct","Nov","Dec"},0),LEFT([@Opened],2))', "date", [serial(t) for t in opened]),
                    "Floor_sqm": ('=VALUE(SUBSTITUTE([@[Floor area (sq m)]],",",""))', "int",
                                  [int(v.replace(",", "")) for v in WEB["Floor area (sq m)"]])})
    first = WEB.Store[opened.idxmin()]
    m.turn()
    m.ex("What is the total floor area of the 12 stores (sq m)?", "=SUM(tWebStores[Floor_sqm])",
         int(sum(int(v.replace(",", "")) for v in WEB["Floor area (sq m)"])), "int", short="Total floor area")
    m.ex_text("Which store opened first?", first, formula="=INDEX(tWebStores[Store],MATCH(MIN(tWebStores[Opened_Date]),tWebStores[Opened_Date],0))",
              short="First store opened")
    m.ex_text("How many tables does From Web find on the page?", "2", alts=("two",))

    # ---------------------------------------------------------------- 2.4.2
    m.topic("2.4.2", "Beginner: import and extract the PDF bulletin",
            "Get Data › From File › From PDF (Microsoft 365) finds the tables on each page. Pick the one that looks like the printed "
            "table, promote headers and remove the printed Total row. Always spot-check a few cells against the PDF.")
    m.steps(["Data › Get Data › From File › From PDF › SuperMart_Monthly_Sales_Bulletin_2025.pdf.",
             "Navigator: choose Table001 (page 1) › Transform Data.",
             "Use First Row as Headers; filter Store to remove 'Total'.",
             "Spot-check three cells against the printed PDF (worked solution)."])
    m.h("tBulletin (₦ millions)")
    m.s.table(BUL, "tBulletin", fmts={c: "dec1" for c in BUL.columns[1:]}, widths={"Store": 26})
    b1 = BUL.set_index("Store")
    m.worked([("Lagos - Ikeja, Dec", '=INDEX(tBulletin[Dec],MATCH("Lagos - Ikeja",tBulletin[Store],0))', float(b1.loc["Lagos - Ikeja", "Dec"]),
               "dec1", "Matches the printed 19.3."),
              ("Port Harcourt - GRA, Oct", '=INDEX(tBulletin[Oct],MATCH("Port Harcourt - GRA",tBulletin[Store],0))',
               float(b1.loc["Port Harcourt - GRA", "Oct"]), "dec1", ""),
              ("Stores extracted", "=ROWS(tBulletin)", len(BUL), "int", "12: the Total row is gone.")])
    m.turn()
    m.ex("What does the bulletin show for Jos - Rayfield in June (₦ millions)?", '=INDEX(tBulletin[Jun],MATCH("Jos - Rayfield",tBulletin[Store],0))',
         float(b1.loc["Jos - Rayfield", "Jun"]), "dec1", short="Jos - Rayfield, Jun")
    m.ex("What is the bulletin's FY 2025 total for all stores (₦ millions, 1 decimal place)?", "=ROUND(SUM(tBulletin[FY_2025]),1)",
         round(float(BUL.FY_2025.sum()), 1), "dec1", short="Bulletin FY total")

    # ---------------------------------------------------------------- 2.4.3
    nm = len(BUL.columns) - 2
    m.topic("2.4.3", "Intermediate: tidy and normalise the bulletin",
            "The bulletin is a cross-tab: months across the top. Analysis needs it tidy: one row per store per month. Unpivot Other "
            "Columns turns the 12 month columns into Month and Sales_m rows, and keeps working if a column is added.")
    code(m, '''let
    Source    = Pdf.Tables(File.Contents(DataFolder & "SuperMart_Monthly_Sales_Bulletin_2025.pdf")),
    Table001  = Source{[Id = "Table001"]}[Data],
    Promoted  = Table.PromoteHeaders(Table001, [PromoteAllScalars = true]),
    NoTotal   = Table.SelectRows(Promoted, each [Store] <> "Total"),          // never import a total row as data
    NoFY      = Table.RemoveColumns(NoTotal, {"FY_2025"}),
    Unpivoted = Table.UnpivotOtherColumns(NoFY, {"Store"}, "Month", "Sales_m"),
    Typed     = Table.TransformColumnTypes(Unpivoted, {{"Sales_m", type number}}, "en-GB")
in
    Typed''', "The tidy query (paste into Advanced Editor)")
    m.worked([("Stores", "=ROWS(tBulletin)", len(BUL), "int", ""),
              ("Month columns", "=COLUMNS(tBulletin)-2", nm, "int", "All columns except Store and FY_2025."),
              ("Rows after unpivot", "=ROWS(tBulletin)*(COLUMNS(tBulletin)-2)", len(BUL) * nm, "int", "One row per store per month.")])
    m.turn()
    m.ex("How many rows would the unpivot give if the bulletin covered only the first 6 months?", "=ROWS(tBulletin)*6", len(BUL) * 6, "int",
         short="Rows for 6 months")
    m.ex_text("Which command keeps Store and turns every other column into rows?", "Unpivot Other Columns", alts=("unpivot other",))

    # ---------------------------------------------------------------- 2.4.4
    led = S.groupby("Store_Name").Sales.sum().reindex(BUL.Store) / 1e6
    led_r = [round(float(v), 1) for v in led]
    diff = [round(float(a) - b_, 1) for a, b_ in zip(BUL.FY_2025, led_r)]
    m.topic("2.4.4", "Intermediate: reconcile the PDF to the ledger",
            "A PDF is a report, not a source of truth. Before using it, compare each store's total with the ledger (tSales25). "
            "Every difference must be explained.")
    rec = BUL[["Store", "FY_2025"]].copy()
    m.h("tRecon: bulletin FY 2025 vs the ledger, ₦ millions")
    m.s.table(rec, "tRecon", fmts={"FY_2025": "dec1"}, widths={"Store": 26},
              calc={"Ledger_m": ("=ROUND(SUMIFS(tSales25[Sales],tSales25[Store_Name],[@Store])/1000000,1)", "dec1", led_r),
                    "Diff": ("=ROUND([@FY_2025]-[@Ledger_m],1)", "dec1", diff)})
    m.turn()
    m.ex("How many stores differ between the PDF and the ledger?", '=COUNTIF(tRecon[Diff],"<>0")', sum(1 for v in diff if v != 0), "int",
         short="Stores that differ")
    m.ex("What is Lagos - Ikeja's FY 2025 ledger total in ₦ millions?", '=XLOOKUP("Lagos - Ikeja",tRecon[Store],tRecon[Ledger_m])',
         led_r[0], "dec1", short="Ikeja ledger total (₦m)")

    # ---------------------------------------------------------------- 2.4.5
    m.topic("2.4.5", "Advanced practical: audit-ready refresh",
            "A query that refreshes automatically needs automatic checks. Put a 99_Checks sheet beside it: each check compares a "
            "refreshed figure with what it must be and shows PASS or FAIL.")
    m.h("99_Checks (live)")
    tot_b = round(float(BUL.FY_2025.sum()), 1)
    tot_l = round(float(S.Sales.sum()) / 1e6, 1)
    assert not any(diff) and abs(tot_b - tot_l) < 0.5, (diff, tot_b, tot_l)     # the four checks below all PASS
    chk = m.worked([("Stores in Table001 = 12", '=IF(ROWS(tBulletin)=12,"PASS","FAIL")', "PASS", None, ""),
              ("Rows after unpivot = 144", '=IF(ROWS(tBulletin)*(COLUMNS(tBulletin)-2)=144,"PASS","FAIL")', "PASS", None, ""),
              ("Stores that differ = 0", '=IF(COUNTIF(tRecon[Diff],"<>0")=0,"PASS","FAIL")', "PASS" if not any(diff) else "FAIL", None, ""),
              ("PDF total within ₦0.5m of the ledger", '=IF(ABS(ROUND(SUM(tBulletin[FY_2025]),1)-ROUND(SUM(tSales25[Sales])/1000000,1))<0.5,"PASS","FAIL")',
               "PASS" if abs(tot_b - tot_l) < 0.5 else "FAIL", None, "12 store totals each rounded to 0.1m: here 733.0 vs 733.2.")])
    cells = list(chk.values())
    m.turn()
    m.ex("How many of the four checks show FAIL right now?", f'=COUNTIF({cells[0]}:{cells[-1]},"FAIL")', 0, "int", short="Failing checks")
    m.ex_open("Next year the PDF adds a 13th store. Which check fails, and what should you do before changing it?",
              "'Stores in Table001 = 12' fails. Confirm the new store exists in the ledger and tStores, then update the expected count.")
    m.finish()


def build(b, D):
    m21(b, D)
    m22(b, D)
    m23(b, D)
    m24(b, D)
