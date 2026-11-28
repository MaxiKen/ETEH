"""Week 2 edition-3 slide copy: Modules 2.1 (external data), 2.2 (Power Query), 2.3 (connectors), 2.4 (web and PDF).
Every figure is computed in w2_fx.py / w2_data.py from the practice files."""
from v3kit import T
from w2_fx import DF, LF, N, I, D, S

MF = D.table("tMonthFiles")
BUL = D.table("tBulletin").set_index("Store")
_ben = BUL.loc["Benin City - Ring Road"]
BEN_SUM = round(sum(_ben[m] for m in D.MONTHS), 1)

M21 = dict(
    id="2.1", title="Understanding External Data Sources",
    subtitle="Where analytical data lives, the file formats it arrives in, and the delimiter, locale and encoding traps.",
    objectives=["Explain what external data is and list the sources an analyst meets.",
                "Recognise the common file formats and what each can and cannot hold.",
                "Import workbooks, CSV and TXT files with the right options.",
                "Use delimiters and text qualifiers correctly.",
                "Import safely and avoid locale and encoding mistakes."],
    topics=[
        T("2.1.1", "What Is External Data?",
          defn="External data is any data stored outside the current workbook (in files, databases, web pages or online "
               "services) that is brought into Excel by importing or connecting to it.",
          simple="Most of the data you analyse is produced by other systems: tills, hospital records, websites. Your first "
                 "job is to bring it in without damaging it.",
          like="Like foodstuff from the market: it comes from outside your kitchen, so you check and wash it before cooking.",
          ex=dict(t="tFiles", cols=["File", "Format", "Rows"], n=5, size=11,
                  note="5 of the 9 Week 2 practice files (folder Week2_Data_Files): all external to the workbook."),
          order=["ex"]),

        T("2.1.2", "Common Data File Formats",
          defn="A file format is the agreed way data is laid out inside a file, shown by its extension (.xlsx, .csv, .txt, "
               ".json, .html, .pdf).",
          simple="The format decides what survives the trip: a CSV keeps only values; a workbook keeps types, formulas "
                 "and several sheets.",
          grid=dict(label="Formats you will meet this week", widths=[1.2, 2.6, 2.4], size=12, rows=[
              ["Format", "What it holds", "Our example"],
              [".xlsx / .xls", "Sheets, tables, types, formulas", "SuperMart_Online_RAW_Export.xlsx (5 sheets)"],
              [".csv", "Plain text, one record per line, comma-separated", "SuperMart_Sales_2025_DIRTY.csv"],
              [".txt", "Plain text with another delimiter, e.g. | or Tab", "Mercyland_Episodes_2023_pipe.txt"],
              [".json", "Nested records from web services (APIs)", "SuperMart_Online_Sellers.json (200 sellers)"],
              [".html", "Web page; tables inside <table> tags", "SuperMart_Store_Directory.html"],
              [".pdf", "Printed layout; tables must be extracted", "SuperMart_Monthly_Sales_Bulletin_2025.pdf"]]),
          watch="A CSV has no types: Excel guesses them on opening. That guess is where most import errors start."),

        T("2.1.3", "Importing from .xlsx and .xls Workbooks",
          simple="Import from another workbook through Get Data, so the link is recorded and can be refreshed, instead of "
                 "copying and pasting.",
          steps=["Data › Get Data › From File › From Excel Workbook", "Pick SuperMart_Online_RAW_Export.xlsx › Import",
                 "The Navigator lists every sheet and table: tick Transactions", "Click Transform Data to clean, or Load"],
          steps_label="HOW TO DO IT IN EXCEL",
          grid=dict(label="What the Navigator shows", widths=[1.5, 1, 1], size=12, rows=[
              ["Sheet", "Rows", "Columns"], ["Transactions", "11,000", "15"], ["Deliveries", "10,430", "13"],
              ["Customers", "2,500", "15"], ["Products", "800", "15"], ["Sellers", "200", "14"]]),
          order=["steps", "grid"], split=0.55,
          watch="Prefer named tables to whole sheets: a table has fixed edges; a sheet may include notes and totals."),

        T("2.1.4", "Importing from CSV Files",
          defn="A CSV (comma-separated values) file is plain text in which each line is one record and commas separate the fields.",
          simple="Use Data › From Text/CSV, not a double-click. The import window lets you check the delimiter, the "
                 "encoding and the types before anything reaches the sheet.",
          steps=["Data › From Text/CSV › SuperMart_Sales_2025_DIRTY.csv",
                 "Check File Origin = 65001: Unicode (UTF-8) so ₦ displays correctly",
                 "Delimiter = Comma; Data Type Detection = Based on first 200 rows",
                 "Click Transform Data to review the columns in Power Query"],
          points=[f"{DF['rows']:,} lines after the header, {DF['cols']} columns", "Sales arrives as text where a value "
                  f"was typed as '₦1,399,930' ({DF['sales_text']} rows)", "Double-clicking would apply your PC's regional settings"],
          points_label="WHAT TO EXPECT IN THIS FILE", order=["steps", "points"], split=0.55,
          watch="Never save over the original CSV after opening it in Excel: leading zeros and date text are changed for good."),

        T("2.1.5", "Importing from TXT Files",
          simple="A .txt file is the same idea as a CSV with a different separator. Tell Excel which character splits the fields.",
          grid=dict(label="First lines of Mercyland_Episodes_2023_pipe.txt", widths=[1], size=11, rows=[
              ["Raw text (first 6 of 22 fields)"],
              ["Episode_ID|Patient_No|Patient_Name|Age|Gender|Blood_Group|…"],
              ["ADM2023-40124|MSH-P001804|Ahmed Y. Oseni|23|Male|O+|…"],
              ["ADM2023-40125|MSH-P004812|Musa C. Lawal|80|Male|AB-|…"]]),
          steps=["Data › From Text/CSV › pick the .txt file", "Delimiter › Custom › type |",
                 f"Check the preview: 22 columns, {len(D.adm23()):,} records", "Transform Data or Load"],
          order=["grid", "steps"], split=0.55,
          watch="Fixed-width text (no delimiter, columns at fixed positions) needs Delimiter › Fixed Width."),

        T("2.1.6", "Understanding Delimiters and Text Qualifiers",
          defn="A delimiter is the character that separates fields in a text file. A text qualifier (usually a double "
               "quote) wraps a field that itself contains the delimiter.",
          simple="If a product name contains a comma, quotes around it stop the comma being read as a new column.",
          grid=dict(label="A real line from the monthly sales files (Product_Name field)", widths=[2.4, 2.2], size=12, rows=[
              ["In the file", "Read by Excel as"],
              ["…,\"Linden 10\"\" Round Wall Clock, Black\",48896,…", "Linden 10\" Round Wall Clock, Black"],
              ["Comma inside quotes", "Part of the text, not a new column"],
              ["Two double quotes \"\"", "One quote character inside the text"]]),
          points=["Comma (,): CSV", "Semicolon (;): European CSV", "Pipe (|): system exports", "Tab: copied from web pages"],
          points_label="COMMON DELIMITERS", order=["grid", "points"], split=0.62),

        T("2.1.7", "Importing Data into Excel Safely",
          simple="A safe import changes nothing in the source, keeps every record and gives each column the right type.",
          like="Like receiving goods at a warehouse: count the cartons against the waybill before you sign.",
          points=["Keep the raw file read-only; work on a copy or a query", "Import with Get Data, never by double-click",
                  "Set column types yourself: IDs and phone numbers as Text", "Compare the row count with the source",
                  "Check a control total (e.g. total Sales) after loading"],
          points_label="SAFE-IMPORT CHECKLIST",
          grid=dict(label="What goes wrong without it", widths=[1.6, 1.6], size=12, rows=[
              ["Value in the file", "After a careless open"], ["07053939110 (phone)", "7053939110"],
              ["1234567890123456789 (19-digit ID)", "1.23457E+18: digits after the 15th lost"], ["04/05/2025", "5 April or 4 May"],
              ["0012 (code)", "12"]]),
          order=["points", "grid"], split=0.5),

        T("2.1.8", "Locale and Encoding Traps (UTF-8, Windows-1252)",
          defn="Locale is the set of regional rules for numbers and dates (decimal mark, separators, day/month order). "
               "Encoding is the code table that turns bytes into characters.",
          simple="The same text means different things under different rules. Tell Excel the file's locale and encoding.",
          grid=dict(label="SuperMart_Store_Targets_EU_semicolon.csv (Windows-1252, European format)", widths=[1.8, 1.7, 1.6],
                    size=12, rows=[
                        ["In the file", "Wrong reading", "Right reading"],
                        ["2400,0", "24000 or text", "2,400.0 sq m"],
                        ["142 000 000,00", "Text", "₦142,000,000.00"],
                        ["₦ saved as UTF-8, opened as 1252", "â‚¦", "₦"]]),
          steps=["File Origin: pick the file's encoding", "Delimiter: Semicolon",
                 "In Power Query: Change Type › Using Locale › Decimal Number, French (France)"],
          steps_label="THE FIX", order=["grid", "steps"], split=0.6),
    ],
    takeaways=["External data comes from files, databases, web pages and services; bring it in with Get Data.",
               "The format decides what survives: CSV and TXT carry values only.",
               "Check delimiter, text qualifier, encoding and locale before loading.",
               "Set types yourself: IDs, codes and phone numbers are text.",
               "Confirm the row count and a control total after every import."],
)

M22 = dict(
    id="2.2", title="Introduction to Power Query",
    subtitle="The data preparation engine inside Excel: why it exists, how the editor is laid out and how it records every step.",
    objectives=["Explain what Power Query is and why it beats manual cleaning for repeated work.",
                "Describe the Extract-Transform-Load cycle.",
                "Open the editor and name each part of it.",
                "Read Applied Steps and use the data profiling tools.",
                "Manage queries and choose how to load the result."],
    topics=[
        T("2.2.1", "What Is Power Query?",
          defn="Power Query (Get & Transform in Excel) is Microsoft's data connection and preparation engine. It imports "
               "data from many sources and records each transformation as a step that can be replayed.",
          simple="You clean the data once; Power Query remembers how. Next month you press Refresh and the same steps run "
                 "on the new file.",
          like="Like a recorded recipe: the cook writes each step down, so anyone can make the same dish tomorrow.",
          points=["Data tab › Get & Transform Data group", "Also in Power BI, with the same engine and M language",
                  "Never changes the source file", "Excel 2016 and later; Microsoft 365 gets new features first"],
          points_label="WHERE TO FIND IT"),

        T("2.2.2", "Why Use Power Query: The Case for Automation",
          simple="Anything you will do more than once should be a query. The cost of manual cleaning repeats every month; "
                 "a query is built once.",
          ex=dict(t="tMonthFiles", cols=["File", "Rows", "Sales"], n=4,
                  note=f"12 monthly files, {MF.Rows.sum():,} rows, {N(MF.Sales.sum())}: one Combine Files query stacks them all."),
          points=["Repeatable: Refresh reruns every step", "Auditable: every step is listed and named",
                  "Scales past 1,048,576 rows via the Data Model", "Fewer copy-paste errors"],
          points_label="WHAT YOU GAIN", order=["ex", "points"], split=0.58),

        T("2.2.3", "Power Query vs Traditional Excel Cleaning",
          simple="Both can produce a clean table. The difference is what happens next month.",
          grid=dict(label="Same job, two ways", widths=[1.5, 2.1, 2.1], size=12.5, rows=[
              ["", "Traditional Excel", "Power Query"],
              ["Where the work happens", "In the cells, overwriting values", "In recorded steps; source untouched"],
              ["Next month", "Repeat every action by hand", "Click Refresh"],
              ["Audit trail", "None unless you write notes", "Applied Steps list"],
              ["Data size", "Up to 1,048,576 rows per sheet", "Millions via the Data Model"],
              ["Best for", "One-off fixes, small lists", "Repeated imports and merges"]])),

        T("2.2.4", "The ETL Cycle in Power Query",
          defn="ETL (Extract, Transform, Load) is the process of reading data from sources, reshaping and cleaning it, "
               "and writing the result to a destination for analysis.",
          diagram=("flow", dict(steps=["Extract", "Transform", "Load", "Refresh"],
                                sub=["Connect to the CSV, folder, workbook, web page or PDF",
                                     "Remove, fix, split, combine, reshape: each a step",
                                     "Table on a sheet, PivotTable, connection or Data Model",
                                     "Re-run every step on new data"])),
          simple="Extract and Load are usually one click each; nearly all the work is in Transform.",
          order=["diagram"]),

        T("2.2.5", "Opening the Power Query Editor",
          simple="There are four doors into the same editor.",
          grid=dict(label="Ways in", widths=[2.2, 2.4], size=12.5, rows=[
              ["Route", "When to use it"],
              ["Data › Get Data › From File… › Transform Data", "Starting a new import"],
              ["Data › From Table/Range", "The data is already an Excel table in this workbook"],
              ["Data › Queries & Connections › double-click a query", "Editing an existing query"],
              ["Data › Get Data › Launch Power Query Editor", "Opening the editor directly"]]),
          watch="From Table/Range turns your range into a table first. Check the headers it proposes."),

        T("2.2.6", "Power Query Editor Interface Anatomy",
          simple="Learn the five areas once; every Power Query lesson refers to them.",
          diagram=("pqeditor", dict(highlight=("steps",), labels={
              "Ribbon": "Home, Transform, Add Column, View",
              "Queries pane": "every query in the workbook",
              "Formula bar": "the M code of the selected step (View › Formula Bar)",
              "Data preview": "column headers with type icons and quality bar",
              "Query Settings": "name and Applied Steps"})),
          order=["diagram"]),

        T("2.2.7", "Queries, Applied Steps and Query Settings",
          defn="A query is a named set of instructions that produces one table. Each instruction is an applied step, "
               "listed in Query Settings in the order it runs.",
          simple="Click any step to see the data as it was at that point. The last step is what gets loaded.",
          grid=dict(label="Applied Steps of a cleaning query on the dirty sales file", widths=[1.6, 2.6], size=12, rows=[
              ["Step", "What it did"], ["Source", "Read SuperMart_Sales_2025_DIRTY.csv"],
              ["Promoted Headers", "First row became the column names"], ["Changed Type", "Types guessed for 21 columns"],
              ["Removed Columns", "Dropped the empty Column1"], ["Trimmed Text", "Removed spaces around Zone"],
              ["Removed Duplicates", "Dropped repeated records"]]),
          watch="Deleting or changing an early step can break the steps after it. Check the last step after any edit."),

        T("2.2.8", "Data Preview, Column Quality, Distribution and Profile",
          simple="The View tab can show, under every header, how many values are valid, errors or empty, and how many "
                 "are distinct.",
          grid=dict(label="View tab › Data Preview options", widths=[1.6, 2.8], size=12, rows=[
              ["Option", "Shows"], ["Column quality", "% Valid, % Error, % Empty"],
              ["Column distribution", "Distinct and unique counts, with a mini histogram"],
              ["Column profile", "Statistics and value distribution for the selected column"]]),
          points=[f"Quantity in the dirty file: {DF['qty_blank'] + DF['blank_rows']} empty cells "
                  f"({DF['qty_blank']} missing values + {DF['blank_rows']} spacer rows)",
                  f"Zone: {DF['zone_space']} values with extra spaces show as separate distinct values"],
          points_label="SEEN IN THE DIRTY SALES FILE", order=["grid", "points"],
          watch="Profiling uses the top 1,000 rows by default. Click the status bar to profile the entire data set."),

        T("2.2.9", "Managing Queries: Groups, Duplicate, Reference",
          simple="As a workbook grows, organise queries in groups and reuse them instead of rebuilding.",
          grid=dict(label="Right-click a query", widths=[1.2, 2.6, 2.0], size=12, rows=[
              ["Command", "What it creates", "Use it when"],
              ["Duplicate", "An independent copy of all the steps", "You want a variation that will not follow changes"],
              ["Reference", "A new query that starts from the other query's result", "Several outputs share one cleaning query"],
              ["Move To Group", "A folder in the Queries pane", "Separating Staging, Clean and Reports"]]),
          watch="Prefer Reference to Duplicate: fix the cleaning once and every referencing query gets the fix."),

        T("2.2.10", "Loading Data into Excel and the Data Model",
          simple="Close & Load To… decides where the result goes.",
          grid=dict(label="Home › Close & Load To…", widths=[1.8, 2.8], size=12, rows=[
              ["Option", "Use it for"], ["Table", "A normal Excel table on a sheet (up to 1,048,576 rows)"],
              ["PivotTable Report", "Summarise without putting the rows on a sheet"],
              ["Only Create Connection", "Staging queries other queries reference"],
              ["Add this data to the Data Model", "Large data, or tables you will relate (Week 3)"]]),
          watch="Staging queries should be Connection only. Loading every query as a table doubles the file size."),
    ],
    takeaways=["Power Query records each cleaning step and replays it on Refresh.",
               "ETL: extract, transform, load; nearly all the work is in Transform.",
               "Five editor areas: ribbon, Queries pane, formula bar, data preview, Query Settings.",
               "Use column quality and profiling on the entire data set, not the first 1,000 rows.",
               "Reference instead of duplicate; load staging queries as connections only."],
)

M23 = dict(
    id="2.3", title="Connecting and Importing with Power Query",
    subtitle="The connectors you will actually use: workbooks, CSV, text, several sheets, whole folders, parameters and refresh.",
    objectives=["Connect Power Query to workbooks, CSV and text files.",
                "Import several worksheets from one workbook.",
                "Combine every file in a folder into one table.",
                "Change source paths and use parameters.",
                "Refresh queries and control refresh options."],
    topics=[
        T("2.3.1", "Connecting to Excel Workbooks",
          simple="Power Query reads a workbook as a list of sheets and tables. You choose which one to open.",
          fx=[["Excel.Workbook(File.Contents(\"…\\SuperMart_Online_RAW_Export.xlsx\"), null, true)", "Lists the 5 sheets"],
              ["Source{[Item=\"Transactions\",Kind=\"Sheet\"]}[Data]", "Opens the Transactions sheet (11,000 rows)"],
              ["Table.PromoteHeaders(…)", "Makes row 1 the headers"]],
          fx_label="THE M CODE BEHIND THE NAVIGATOR", fx_head=["Step formula", "What it does"], fx_widths=[3.6, 2.0],
          watch="A sheet is read up to its last used cell. Notes typed below the data become extra rows."),

        T("2.3.2", "Connecting to CSV Files",
          simple="The CSV connector exposes the options you checked in the preview as arguments you can edit later.",
          fx=[["Csv.Document(File.Contents(path), [Delimiter=\",\", Encoding=65001, QuoteStyle=QuoteStyle.Csv])",
               "Reads the file as UTF-8, honouring quotes"],
              ["Table.PromoteHeaders(Source, [PromoteAllScalars=true])", "Header row"],
              ["Table.TransformColumnTypes(…)", "Changed Type step"]],
          fx_label="SOURCE STEP FOR SuperMart_Sales_2025_DIRTY.csv", fx_head=["Step formula", "What it does"],
          fx_widths=[3.6, 2.0],
          points=[f"Result: {DF['rows']:,} rows × {DF['cols']} columns", "QuoteStyle.Csv keeps commas inside quotes",
                  "Encoding 65001 = UTF-8; 1252 = Western European"],
          points_label="CHECK", order=["fx", "points"], stack=True),

        T("2.3.3", "Connecting to Text Files",
          simple="Text files use the same connector; only the delimiter changes.",
          fx=[["Csv.Document(File.Contents(path), [Delimiter=\"|\", Columns=22, Encoding=65001])",
               f"Mercyland pipe file: {len(D.adm23()):,} records"],
              ["Splitter.SplitTextByPositions({0, 12, 24})", "Fixed-width files: split at character positions"]],
          fx_label="PIPE AND FIXED-WIDTH TEXT", fx_head=["Step formula", "What it does"], fx_widths=[3.6, 2.0],
          watch="Columns=22 is fixed in the step. If the source adds a column, edit the step or remove the Columns option."),

        T("2.3.4", "Importing Multiple Worksheets from One Workbook",
          simple="Tick Select multiple items in the Navigator to get one query per sheet, or expand the sheet list to stack "
                 "sheets that share the same columns.",
          steps=["From Excel Workbook › tick Select multiple items", "Tick each sheet you need › Transform Data",
                 "To stack similar sheets: connect to the workbook, keep the Data column, expand it"],
          grid=dict(label="SuperMart_Online_RAW_Export.xlsx", widths=[1.4, 1, 2.2], size=12, rows=[
              ["Sheet", "Rows", "Joins on"], ["Transactions", "11,000", "seller_id, product_id, customer_id"],
              ["Sellers", "200", "seller_id"], ["Products", "800", "product_id"], ["Customers", "2,500", "customer_id"],
              ["Deliveries", "10,430", "order_id"]]),
          order=["steps", "grid"], split=0.48,
          watch="These five sheets have different columns: they become separate queries to relate, not one stacked table."),

        T("2.3.5", "Combining Files from a Folder",
          defn="The Folder connector lists every file in a folder; Combine Files applies the same transformation to each "
               "file and appends the results into one table.",
          simple="Point at the folder once. Next month, drop the new file in and press Refresh.",
          ex=dict(t="tMonthFiles", cols=["File", "Rows", "Sales"], n=4, size=11,
                  note=f"4 of 12 files. Combined: {MF.Rows.sum():,} rows, {N(MF.Sales.sum())}."),
          steps=["Data › Get Data › From File › From Folder", "Browse to SuperMart_Monthly_Sales_2025 › OK",
                 "Combine › Combine & Transform Data", "Pick the sample file › OK",
                 "Keep Source.Name: it records which file each row came from"],
          order=["ex", "steps"], split=0.5,
          watch="Every file must have the same columns, and only data files may sit in the folder."),

        T("2.3.6", "File Paths and Source Settings",
          simple="A query remembers the full path of its source. Move the file, or send the workbook to a colleague, and "
                 "the path breaks.",
          steps=["Home › Data source settings", "Select the source › Change Source… › browse to the new location",
                 "Or click the gear beside the Source step"],
          grid=dict(label="Typical error after moving a file", widths=[1], size=12, rows=[
              ["DataSource.Error: Could not find file"],
              ["'C:\\Users\\Ade\\Downloads\\SuperMart_Sales_2025_DIRTY.csv'"]]),
          order=["steps", "grid"], split=0.5,
          watch="Keep sources in one shared project folder (e.g. OneDrive or SharePoint), not in Downloads."),

        T("2.3.7", "Using Parameters for Paths and Credentials",
          defn="A parameter is a named value that queries can use in place of a typed constant, such as a folder path.",
          simple="Store the folder path once as a parameter. Point every query at the parameter; to move the project, "
                 "change one value.",
          steps=["Home › Manage Parameters › New", "Name: DataFolder; Type: Text; Current Value: the folder path",
                 "In each Source step replace the typed path with DataFolder & \"SuperMart_Sales_2025_DIRTY.csv\""],
          points=["Usernames and passwords live in Data source settings, never in M code",
                  "Privacy levels stop data from one source leaking into another"],
          points_label="CREDENTIALS", order=["steps", "points"], split=0.58),

        T("2.3.8", "Refreshing Queries After Source Changes",
          simple="Refresh reruns every step on the current source. Choose when it happens.",
          grid=dict(label="Refresh options", widths=[2.2, 2.4], size=12, rows=[
              ["Command or setting", "Effect"], ["Data › Refresh All (Ctrl+Alt+F5)", "Every query and PivotTable"],
              ["Right-click a table › Refresh (Alt+F5)", "That query only"],
              ["Query Properties › Refresh data when opening the file", "Always up to date on open"],
              ["Query Properties › Refresh every n minutes", "Live dashboards"]]),
          points=["Add the 2026-01 file to the folder", "Refresh All", "Row count and Source.Name list grow by themselves"],
          points_label="TRY IT WITH THE MONTHLY FOLDER", order=["grid", "points"], split=0.62),
    ],
    takeaways=["Workbook, CSV and text connectors all show their options as editable step arguments.",
               "Select multiple items for several sheets; stack only sheets with the same columns.",
               "From Folder › Combine appends every file; keep Source.Name for traceability.",
               "Change Source fixes broken paths; parameters make projects portable.",
               "Refresh All reruns every query; set refresh-on-open where it helps."],
)

M24 = dict(
    id="2.4", title="Web and PDF Data Sources (Practice Files)",
    subtitle="Practise on two published sources: inspect and import the store directory web page, extract the PDF sales "
             "bulletin, tidy it, reconcile it to the ledger and refresh it safely.",
    objectives=["Read a web page before importing its tables.", "Extract a table from a PDF and spot-check it.",
                "Unpivot a cross-tab report into a tidy table.", "Reconcile imported figures to the system of record.",
                "Build a refresh that anyone can rerun and prove correct."],
    topics=[
        T("2.4.1", "Beginner: Inspect the Store Directory Web Page",
          simple="Data › From Web reads the tables inside a web page (online, or a saved .html file). Read the page in a "
                 "browser first: which tables exist, what one row means, the units and the footnotes.",
          like="Like reading a menu before ordering: you check what each dish is and its size before you tell the waiter.",
          ex=dict(t="tWebStores", cols=["Store ID", "Store", "Opened", "Floor area (sq m)"], n=4, size=11.5,
                  note="Table 1 of SuperMart_Store_Directory.html (12 stores)."),
          steps=["Open the file in a browser; note 2 tables and a footnote", "Data › From Web › paste the file:/// path",
                 "Navigator › Table 1 › Transform Data", "Opened is text such as 02 Mar 2015: Change Type › Using Locale "
                 "› Date, English (United Kingdom)"],
          order=["ex", "steps"], split=0.52,
          watch="Ignoring the footnote. Sales per square metre must be labelled 'excluding car parks and warehouse'."),

        T("2.4.2", "Beginner: Import and Extract the PDF Bulletin",
          simple="Get Data › From File › From PDF lists each table it detects (Table001…) and each whole page (Page001…). "
                 "PDFs are made for reading, so compare a few extracted numbers with the printed page.",
          ex=dict(t="tBulletin", cols=["Store", "Jan", "Feb", "Oct", "Dec", "FY_2025"], n=4, size=11,
                  note="SuperMart_Monthly_Sales_Bulletin_2025.pdf, Table001: 12 stores × 12 months, ₦ millions."),
          grid=dict(label="Spot-check three cells", widths=[2.0, 0.9, 0.9], size=12, rows=[
              ["Cell", "Printed", "Extracted"],
              ["Lagos - Ikeja, Dec", f"{BUL.loc['Lagos - Ikeja', 'Dec']:.1f}", f"{BUL.loc['Lagos - Ikeja', 'Dec']:.1f}"],
              ["Port Harcourt - GRA, Oct", f"{BUL.loc['Port Harcourt - GRA', 'Oct']:.1f}",
               f"{BUL.loc['Port Harcourt - GRA', 'Oct']:.1f}"],
              ["Jos - Rayfield, FY 2025", f"{BUL.loc['Jos - Rayfield', 'FY_2025']:.1f}",
               f"{BUL.loc['Jos - Rayfield', 'FY_2025']:.1f}"]]),
          order=["ex", "grid"], split=0.6,
          watch="Loading the first item found without looking: Page001 is the whole page text, not the table."),

        T("2.4.3", "Intermediate: Tidy and Normalise the Bulletin",
          defn="Unpivoting turns a set of columns (here one per month) into two columns, Attribute and Value, giving one "
               "row per original cell.",
          simple="The bulletin is a cross-tab: easy to read, hard to analyse. Unpivot it into Store | Month | Sales_m.",
          steps=["Use First Row as Headers; filter out the Total row", "Remove the FY_2025 column",
                 "Select Store › Transform › Unpivot Other Columns", "Rename Attribute › Month and Value › Sales_m; set Decimal"],
          grid=dict(label="Shape at each stage", widths=[2.3, 1.2], size=12, rows=[
              ["Stage", "Rows × columns"], ["As extracted, totals removed", "12 × 14"], ["After removing FY_2025", "12 × 13"],
              ["After Unpivot Other Columns", "144 × 3"]]),
          order=["steps", "grid"], split=0.56,
          watch="Use Unpivot Other Columns on Store, not Unpivot on the months, so a new month column is picked up automatically."),

        T("2.4.4", "Intermediate: Reconcile the PDF to the Ledger",
          defn="Reconciliation is checking that figures from one source agree with the system of record, and explaining "
               "every difference.",
          simple="Merge the PDF table with the ledger on Store, round the ledger to ₦0.1m and flag any store that differs.",
          grid=dict(label="Result (₦ millions)", widths=[2.6, 0.9, 0.9], size=12, rows=[
              ["Check", "PDF", "Ledger"],
              ["Lagos - Ikeja, FY 2025", f"{BUL.loc['Lagos - Ikeja', 'FY_2025']:.1f}",
               f"{S[S.Store_Name == 'Lagos - Ikeja'].Sales.sum() / 1e6:.1f}"],
              ["Stores that differ (of 12)", "0", "0"],
              ["Benin: 12 printed months added up", f"{BEN_SUM:.1f}", f"{_ben['FY_2025']:.1f}"],
              ["All 144 printed monthly cells added up", f"{BUL[D.MONTHS].values.sum():.1f}", f"{S.Sales.sum() / 1e6:.1f}"]]),
          points=["Every store's FY figure agrees with the ledger", "Adding rounded cells drifts: each is rounded to 0.1",
                  "Take totals from the ledger in naira, never from rounded report cells"],
          points_label="CONCLUSION", order=["grid", "points"], split=0.6),

        T("2.4.5", "Advanced Practical: Audit-Ready Refresh",
          simple="An audit-ready refresh can be rerun by anyone, and proved correct, without asking you.",
          like="Like a pilot's checklist before take-off: the same checks, ticked every time, by whoever is flying.",
          points=["Add a Source column: file name and table ID (Table001)", "Control queries: row count and total check",
                  "Load the controls to a 99_Checks sheet", "Log each refresh: date, checks, initials"],
          points_label="BUILD IT",
          grid=dict(label="99_Checks after each refresh", widths=[2.2, 1.0], size=12, rows=[
              ["Check", "Expected"], ["Stores in Table001", "12"], ["Rows after unpivot", "144"],
              ["Stores differing from ledger", "0"]]),
          order=["points", "grid"], split=0.56,
          watch="A refresh that runs without errors can still be wrong. The checks are what prove it is right."),
    ],
    takeaways=["Read a web page or PDF before importing it; footnotes change meanings.",
               "Pick the table (Table001), not the page, and spot-check three cells.",
               "Unpivot Other Columns turns report cross-tabs into tidy rows.",
               "Reconcile to the ledger; never total rounded report cells.",
               "Audit-ready means a Source column, control checks and a refresh log."],
)

MODULES = [M21, M22, M23, M24]
