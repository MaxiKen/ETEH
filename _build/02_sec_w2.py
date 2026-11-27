"""Week 2 Section Examples workbook (modules 2.1 - 2.11): import, Power Query, cleaning, 65 functions, M, PivotTables."""
from __future__ import annotations

import datetime as dt
import json
import os
import re
import sys

import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(__file__))
from common import SEC, W2D  # noqa: E402
from wk_common import DATASETS_TEXT, L, lc_adm, sm_sales, smo  # noqa: E402
from xlhelp import LAST_COL, Book  # noqa: E402

OUT = SEC / "ETEH_Week2_Section_Examples.xlsx"
TODAY = dt.date(2026, 11, 27)

sm = sm_sales()
sm["Year"] = sm.Order_Date.dt.year
sm["Quarter"] = "Q" + sm.Order_Date.dt.quarter.astype(str)
sm["Month"] = sm.Order_Date.dt.strftime("%m-%b")
s25 = sm[sm.Year == 2025].reset_index(drop=True)
stores = L("sm_stores")
lc = lc_adm()
o = smo()
dirty = L("dirty_sm")
ctrl = L("dirty_sm_ctrl").set_index("Control total (after cleaning)").Value
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

b = Book(str(OUT), "ETEH Week 2 - Section Examples", "Week 2 · Power Query, Cleaning, Functions and PivotTables")
F = b.F


def spill(sh, r, c, formula, values, kinds=None):
    """Dynamic-array formula at (r,c) with cached spill values (2-D list)."""
    nr, nc = len(values), len(values[0])
    kinds = kinds or [None] * nc
    first = values[0][0]
    sh.ws.write_dynamic_array_formula(r, c, r + nr - 1, c + nc - 1, formula, b.nf(kinds[0] or "gen", style="result"),
                                      first.item() if hasattr(first, "item") else first)
    for i in range(nr):
        for j in range(nc):
            if i == 0 and j == 0:
                continue
            sh._w(r + i, c + j, values[i][j], b.nf(kinds[j] or "gen", style="result"))


# ---------------------------------------------------------------------- data sheets
stores_x = stores.copy()
act = s25.groupby("Store_ID").agg(Sales_2025=("Sales", "sum"), Profit_2025=("Profit", "sum")).reset_index()
stores_x = stores_x.merge(act, on="Store_ID")
stores_x = stores_x.rename(columns={"Annual_Target_2025": "Target_2025"})
stores_x["Opened"] = pd.to_datetime(stores_x.Opened)

sales_cols = ["Row_ID", "Order_ID", "Order_Date", "Ship_Date", "Ship_Mode", "Customer_ID", "Customer_Name", "Segment", "Store_ID",
              "Store_Name", "City", "State", "Zone", "Product_ID", "Category", "Sub_Category", "Product_Name", "Sales", "Quantity",
              "Discount", "Profit"]


def iso_or_text(v):
    if isinstance(v, str) and re.match(r"^\d{4}-\d{2}-\d{2}$", v):
        return pd.Timestamp(v)
    return v


dd = dirty.copy()
dd["Order_Date"] = dd.Order_Date.map(iso_or_text)
dd["Ship_Date"] = dd.Ship_Date.map(iso_or_text)
dd = dd.astype(object).where(dd.notna(), None)

# ======================================================================================
# 2.1 External formats
# ======================================================================================
s = b.sheet("2.1 External Formats", "2.1  Understanding External Data Sources",
            "The Week 2 practice files: formats, delimiters, qualifiers, locale and encoding", "2.1.1 - 2.1.8")
s.band("2.1.1-2", "Every practice file in Week2_Data_Files - what it is and how to open it")
inv = []
for fn, fmt, delim, enc, how in [
    ("SuperMart_Sales_2025_DIRTY.csv", "CSV", "Comma", "UTF-8 with BOM", "Data > From Text/CSV"),
    ("SuperMart_Monthly_Sales_2025\\ (12 files)", "Folder of CSV", "Comma", "UTF-8", "Data > Get Data > From Folder > Combine"),
    ("Lifecare_Admissions_2023_pipe.txt", "TXT", "Pipe |", "UTF-8", "From Text/CSV, delimiter = Custom '|'"),
    ("SuperMart_Store_Targets_EU_semicolon.csv", "CSV (EU locale)", "Semicolon ;", "Windows-1252", "From Text/CSV, File Origin 1252, locale German/French"),
    ("Lifecare_Admissions_2023_DIRTY.csv", "CSV", "Comma", "UTF-8", "From Text/CSV, then clean (2.5/2.6)"),
    ("SuperMart_Online_RAW_Export.xlsx", "Workbook", "-", "-", "Data > From Workbook (5 sheets)"),
    ("SuperMart_Online_Sellers.json", "JSON", "-", "UTF-8", "Get Data > From File > From JSON, expand records"),
    ("SuperMart_Store_Directory.html", "Web page", "-", "UTF-8", "Data > From Web (paste file:/// path)"),
    ("SuperMart_Monthly_Sales_Bulletin_2025.pdf", "PDF", "-", "-", "Get Data > From File > From PDF (Microsoft 365)")]:
    p = W2D / fn.split("\\")[0]
    if p.is_dir():
        size = sum(f.stat().st_size for f in p.iterdir()) / 1024
        rows = sum(sum(1 for _ in open(f, encoding="utf-8")) - 1 for f in p.iterdir())
    else:
        size = p.stat().st_size / 1024
        rows = None
        if p.suffix in (".csv", ".txt"):
            rows = sum(1 for _ in open(p, encoding="cp1252" if "EU" in fn else "utf-8-sig", errors="replace")) - 1
    inv.append([fn, fmt, delim, enc, rows, size, how])
s.grid(inv, header=["File", "Format", "Delimiter", "Encoding", "Data rows", "Size KB", "How to import"], fmts=[None, None, None, None, "int", "int", None])
s.para("Data rows for CSV/TXT include any blank or duplicate rows in the file - which is exactly why you check counts after import.", "note")
s.gap()

s.band("2.1.6", "Delimiters and text qualifiers - why product names have quotes around them")
s.scenario("SuperMart product names contain commas and inch-marks, e.g. Linden 10\" Round Wall Clock, Black. In a comma-delimited file the "
           "comma inside the name would split the field in two, so the exporter wraps the field in double quotes and doubles any quote inside it.")
raw_line = open(W2D / "SuperMart_Monthly_Sales_2025" / "2025-01_SuperMart_Sales.csv", encoding="utf-8").readlines()[1].strip()
s.code("Raw line in 2025-01_SuperMart_Sales.csv:\n" + raw_line)
s.grid([["Delimiter ,", "Separates the 21 fields"], ['Text qualifier "', 'Wraps a field that itself contains a comma'],
        ['Escaped quote ""', 'A literal " inside a qualified field (10" = 10 inches)'],
        ["Result after import", "Product_Name = Linden 10\" Round Wall Clock, Black  (one field, not two)"]], header=["Element", "Role"], merge_last=True)
s.gap()

s.band("2.1.8", "Locale and encoding traps - the EU targets file")
s.scenario("SuperMart's European franchise consultant sent targets from a French-locale laptop: semicolons between fields, comma as decimal "
           "mark, a space (or non-breaking space) as thousands separator, saved in Windows-1252. Opened with UK settings, every target becomes text.")
eu = open(W2D / "SuperMart_Store_Targets_EU_semicolon.csv", encoding="cp1252").read().splitlines()
s.code("\n".join(eu[:5]))
eut = pd.DataFrame({"Store_ID": [l_.split(";")[0] for l_ in eu[1:7]], "Target_as_text": [l_.split(";")[-1] for l_ in eu[1:7]]})
vals = [float(x.replace(" ", "").replace(",", ".")) for x in eut.Target_as_text]
s.table(eut, "tEuTargets", calc={"Target_NUMBERVALUE": ('=NUMBERVALUE([@Target_as_text],","," ")', "naira", vals),
                                  "Is_Number_Before": ("=ISNUMBER([@Target_as_text])", None, [False] * len(eut))},
        widths={"Target_as_text": 18, "Target_NUMBERVALUE": 20})
s.para("NUMBERVALUE(text, decimal_separator, group_separator) converts any locale. In Power Query use Change Type > Using Locale > French (France).", "note")
mj = "₦".encode("utf-8").decode("cp1252", errors="replace")
s.grid([["₦1,250,000 saved as UTF-8, opened as Windows-1252", f"{mj}1,250,000", "Re-import with File Origin 65001: Unicode (UTF-8)"],
        ["Nnamdi Azikiwe with é from a French keyboard (Azikiwé) saved as 1252, opened as UTF-8", "Azikiw�", "Re-import with File Origin 1252"]],
       header=["What happened", "What you see", "Fix"])

# ======================================================================================
# 2.2 Introduction to Power Query
# ======================================================================================
s = b.sheet("2.2 Power Query Intro", "2.2  Introduction to Power Query",
            "The ETL cycle on SuperMart_Sales_2025_DIRTY.csv: applied steps, row counts, column quality", "2.2.1 - 2.2.10")
s.band("2.2.4", "ETL in one picture - the FY2025 POS clean-up")
s.grid([["EXTRACT", "Data > From Text/CSV > SuperMart_Sales_2025_DIRTY.csv", "3,387 rows x 21 columns arrive as text and numbers"],
        ["TRANSFORM", "14 applied steps (below) - each one recorded, none typed by hand", "Defects removed, types set, columns split"],
        ["LOAD", "Close & Load To... > Table on a new sheet, or Only Create Connection + Data Model", "3,312 clean rows - matches the POS control total"]],
       header=["Stage", "What you click", "Result"])
s.gap()

s.band("2.2.7", "Applied Steps - what each step does and how many rows survive")
steps = []
d0 = dirty.copy()
steps.append(("Source", "Read the CSV file (UTF-8, comma)", len(d0)))
d1 = d0.dropna(how="all")
steps.append(("Promoted Headers + Removed Blank Rows", "Use row 1 as headers; drop the 15 spacer rows", len(d1)))
d2 = d1.drop(columns=["Column1"]).drop_duplicates()
steps.append(("Removed Columns", "Delete the empty Column1", len(d1)))
steps.append(("Removed Duplicates", "Table.Distinct on all columns (60 exact copies)", len(d2)))
steps.append(("Trimmed Text", "Text.Trim on Zone; Text.Clean + replace CHAR(160) on Customer_Name", len(d2)))
steps.append(("Replaced Category", "Mapping: Furnitre/FURNITURE → Furniture, Tech/Technolgy → Technology ...", len(d2)))
steps.append(("Merged Store Directory", "Bring the official Store_Name from the directory by Store_ID (fixes casing safely)", len(d2)))
steps.append(("Split Store_Location", "Split by ', ' into City and State", len(d2)))
steps.append(("Cleaned Sales", "Remove ₦ and commas, then type Whole Number", len(d2)))
steps.append(("Cleaned Discount", "'20%' → 0.2, type Decimal", len(d2)))
steps.append(("Parsed Order_Date", "Change Type with Locale en-GB for dd/mm/yyyy, dd-Mon-yyyy and yyyy.mm.dd", len(d2)))
steps.append(("Flagged Exceptions", "Add Exception column: blank Quantity, negative Quantity, Ship before Order, keyed x100", len(d2)))
steps.append(("Changed Type", "Final types: dates, whole numbers, decimal, text", len(d2)))
steps.append(("Loaded", "Close & Load", len(d2)))
stp = pd.DataFrame(steps, columns=["Applied_Step", "What it does", "Rows_after_step"])
stp.insert(0, "Step_No", range(1, len(stp) + 1))
s.table(stp, "tSteps", fmts={"Rows_after_step": "int", "Step_No": "gen"}, widths={"Applied_Step": 28, "What it does": 60})
s.calc([("Rows removed in total", "=INDEX(tSteps[Rows_after_step],1)-INDEX(tSteps[Rows_after_step],ROWS(tSteps))", len(d0) - len(d2), "int", "15 blank + 60 duplicates"),
        ("Matches POS control row count?", f'=IF(INDEX(tSteps[Rows_after_step],ROWS(tSteps))={int(ctrl["Clean row count"])},"PASS","FAIL")', "PASS", None,
         f"POS reports {int(ctrl['Clean row count']):,} lines for FY2025")])

s.band("2.2.8", "Column quality and distribution - what the Power Query profile shows for the raw file")
s.scenario("View > Column quality / Column distribution / Column profile. The table reproduces those bars for key columns of the raw file "
           "(profile based on entire data set - change this in the status bar, the default is the top 1,000 rows!).")
prof = []
for c in ["Order_Date", "Customer_Name", "Store_Name", "Zone", "Category", "Sales", "Quantity", "Discount", "Column1"]:
    col = d1[c] if c != "Column1" else d0[c]
    n = len(col)
    empty = col.isna().sum()
    if c == "Sales":
        err = col.map(lambda v: isinstance(v, str)).sum()
    elif c == "Discount":
        err = col.map(lambda v: isinstance(v, str)).sum()
    elif c == "Order_Date":
        err = col.map(lambda v: isinstance(v, str) and not re.match(r"^\d{4}-\d{2}-\d{2}$", v)).sum()
    else:
        err = 0
    prof.append([c, n, n - empty - err, err, empty, col.nunique(dropna=True)])
pf = pd.DataFrame(prof, columns=["Column", "Rows", "Valid", "Error_if_typed", "Empty", "Distinct"])
tp = s.table(pf, "tProfile", fmts={k: "int" for k in ["Rows", "Valid", "Error_if_typed", "Empty", "Distinct"]},
             calc={"Valid_pct": ("=[@Valid]/[@Rows]", "pct", (pf.Valid / pf.Rows).tolist())})
s.ws.conditional_format(tp.rng("Valid_pct", absolute=False), {"type": "data_bar", "bar_color": "#7FB069", "bar_solid": True, "min_type": "num", "min_value": 0, "max_type": "num", "max_value": 1})
s.insight("Category has 12 distinct values for what should be 3 - the profile tells you a mapping step is needed before you write a single formula.")

s.band("2.2.9-10", "Managing and loading queries")
s.grid([["Reference", "qSales_Clean → new query qSales_ByStore that starts from it", "When the new query depends on the cleaned output (one source of truth)"],
        ["Duplicate", "Copies every step into an independent query", "When you want to experiment without touching the original"],
        ["Groups", "Folders: 01 Sources, 02 Staging, 03 Final", "Keeps 20+ queries navigable"],
        ["Load to Table", "Writes rows to a worksheet (max 1,048,576)", "Small/medium results people will read"],
        ["Only Create Connection + Data Model", "Rows live in the model, not on a sheet", "Large tables, relationships, DAX (Week 3)"]],
       header=["Option", "What it does", "Use it when"])

# ======================================================================================
# 2.3 Connecting, folder, parameters
# ======================================================================================
s = b.sheet("2.3 Connect & Combine", "2.3  Connecting and Importing with Power Query",
            "Workbooks, CSV/TXT, combine a folder of 12 monthly files, parameters for paths, refresh", "2.3.1 - 2.3.8")
s.band("2.3.5", "Combine Files from a Folder - twelve monthly CSVs become one table")
mf = []
for f in sorted((W2D / "SuperMart_Monthly_Sales_2025").iterdir()):
    g = pd.read_csv(f)
    mf.append([f.name, len(g), int(g.Sales.sum()), int(g.Profit.sum())])
mfd = pd.DataFrame(mf, columns=["File_Name", "Rows", "Sales", "Profit"])
tm = s.table(mfd, "tMonthFiles", fmts={"Rows": "int", "Sales": "naira", "Profit": "naira"},
             total={"File_Name": "Total", "Rows": "sum", "Sales": "sum", "Profit": "sum"}, widths={"File_Name": 30})
s.calc([("Combined rows = POS control?", f'=IF(SUM(tMonthFiles[Rows])={int(ctrl["Clean row count"])},"PASS","FAIL")', "PASS", None, "3,312 lines"),
        ("Combined sales = POS control?", f'=IF(SUM(tMonthFiles[Sales])={int(ctrl["Total Sales (₦)"])},"PASS","FAIL")', "PASS", None, "₦733,215,255")])
s.code('''let
    // pFolderPath is a parameter (Manage Parameters) - change it once, every query follows
    Source   = Folder.Files(pFolderPath & "SuperMart_Monthly_Sales_2025\\"),
    CsvOnly  = Table.SelectRows(Source, each [Extension] = ".csv" and not Text.StartsWith([Name], "~")),
    Parsed   = Table.AddColumn(CsvOnly, "Data", each Table.PromoteHeaders(
                   Csv.Document([Content], [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
                   [PromoteAllScalars = true])),
    KeepName = Table.SelectColumns(Parsed, {"Name", "Data"}),
    Expanded = Table.ExpandTableColumn(KeepName, "Data", Table.ColumnNames(KeepName{0}[Data])),
    Typed    = Table.TransformColumnTypes(Expanded, {{"Order_Date", type date}, {"Ship_Date", type date},
                   {"Sales", Int64.Type}, {"Quantity", Int64.Type}, {"Discount", type number}, {"Profit", Int64.Type}}),
    Renamed  = Table.RenameColumns(Typed, {{"Name", "Source_File"}})
in
    Renamed''')
s.chart("column", [{"name": "Sales", "categories": tm.ref("File_Name"), "values": tm.ref("Sales")}], "FY2025 sales by monthly file",
        y_fmt='₦#,##0,,"m"', legend=False, rows=16)
s.insight(f"Q4 files are the largest: Sep, Nov and Dec each carry about ₦{mfd.Sales.iloc[[8, 10, 11]].mean() / 1e6:,.0f}m. "
          "When January 2026 arrives, drop the file into the folder and press Refresh All - no new formulas.")

s.band("2.3.6-7", "File paths and parameters - one cell controls every query")
s.scenario("Type your own folder path in the yellow cell (it is a named cell, FolderPath). The M code reads it, so the workbook works on any laptop.")
FP_LABEL = "Folder path (end with a backslash)"
fp = s.inputs([(FP_LABEL, "C:\\ETEH\\01_Section_Examples\\Week2_Data_Files\\", "text", "Named FolderPath")])
b.wb.define_name("FolderPath", "='2.3 Connect & Combine'!" + fp[FP_LABEL])
s.code('''// Query: pFolderPath  (reads the named cell, so no hard-coded C:\\Users\\... paths)
let
    Source = Excel.CurrentWorkbook(){[Name = "FolderPath"]}[Content]{0}[Column1]
in
    Source''')
s.band("2.3.3", "Text file with a pipe delimiter - Lifecare admissions 2023")
pip = open(W2D / "Lifecare_Admissions_2023_pipe.txt", encoding="utf-8").read().splitlines()
s.code("\n".join(x[:150] for x in pip[:3]))
lc23 = lc[lc.Admission_Date.dt.year == 2023]
s.calc([("Rows in the file (excluding header)", len(pip) - 1, len(pip) - 1, "int", "Count after import must equal this"),
        ("Lifecare 2023 admissions in the core dataset", len(lc23), len(lc23), "int", "Independent control"),
        ("Check", f'=IF({len(pip) - 1}={len(lc23)},"PASS","FAIL")', "PASS" if len(pip) - 1 == len(lc23) else "FAIL", None, "")])
s.code('''let
    Source   = Csv.Document(File.Contents(pFolderPath & "Lifecare_Admissions_2023_pipe.txt"),
                            [Delimiter = "|", Encoding = 65001, QuoteStyle = QuoteStyle.None]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed    = Table.TransformColumnTypes(Promoted, {{"Admission_Date", type date}, {"Discharge_Date", type date},
                   {"Age", Int64.Type}, {"Length_of_Stay", Int64.Type}, {"Billing_Amount", Currency.Type}})
in
    Typed''')
s.band("2.3.4 / 2.3.8", "Several sheets from one workbook, and refreshing")
s.grid([["SuperMart_Online_RAW_Export.xlsx", "Excel.Workbook(File.Contents(...), true) lists 5 sheets: customers, sellers, products, deliveries, transactions", "Pick each sheet as its own query; do NOT append - they are different tables (joined in Week 3)"],
        ["Refresh All (Ctrl+Alt+F5)", "Re-runs every query from source", "After dropping a new monthly file into the folder"],
        ["Refresh on open", "Query Properties > Refresh data when opening the file", "Dashboards others open"],
        ["Refresh every N minutes", "Query Properties", "Live operational sheets only - not big files"]],
       header=["Item", "What happens", "Practical use"])

# ======================================================================================
# 2.4 Web & PDF  (rewritten - no external sources: everything is a local practice file)
# ======================================================================================
s = b.sheet("2.4 Web & PDF Sources", "2.4  Web and PDF Data Sources (local practice files)",
            "From Web on a saved HTML page, From PDF on the FY2025 bulletin, unpivot, join and reconcile", "2.4.1 - 2.4.5")
s.band("2.4.1", "Beginner: inspect the web source - SuperMart_Store_Directory.html")
s.scenario("Head Office publishes its store directory as a web page. Open the HTML file in a browser first: it has TWO tables "
           "(the directory and opening hours) and a footnote. Data > From Web, paste the file path (file:///C:/ETEH/.../SuperMart_Store_Directory.html), "
           "and the Navigator lists Table 1 and Table 2 - pick 'stores'.")
html = pd.read_html(str(W2D / "SuperMart_Store_Directory.html"))
hd = html[0]
s.table(hd.rename(columns=lambda c: c.replace(" ", "_").replace("(", "").replace(")", "")), "tWebStores",
        fmts={"Floor_area_sq_m": "gen"}, widths={"Store": 26, "Manager": 18})
s.warn("Opened arrives as text ('02 Mar 2015'): set type Date using Locale English (United Kingdom). Read the page footnote too - floor area excludes car parks, so do not use it for sales-per-square-metre without saying so.")

s.band("2.4.2", "Beginner: import and extract the PDF table - SuperMart_Monthly_Sales_Bulletin_2025.pdf")
s.scenario("Finance's bulletin is a PDF cross-tab: stores down the side, months across, ₦ millions to one decimal. "
           "Get Data > From File > From PDF shows Table001 (store x month) and Table002 (zone summary). This is what Table001 contains:")
ct = s25.assign(M=s25.Order_Date.dt.month).pivot_table(index="Store_Name", columns="M", values="Sales", aggfunc="sum", fill_value=0) / 1e6
ct = ct.reindex(stores.Store_Name).round(1)
ct.columns = MONTHS
pdf_tab = ct.reset_index()
pdf_tab["FY_2025"] = (s25.groupby("Store_Name").Sales.sum().reindex(stores.Store_Name) / 1e6).round(1).values
tpdf = s.table(pdf_tab, "tPdfBulletin", fmts={m: "dec1" for m in MONTHS + ["FY_2025"]}, widths={"Store_Name": 26})

s.band("2.4.3", "Intermediate: tidy and normalise - unpivot the 12 month columns")
s.code('''let
    Source    = Pdf.Tables(File.Contents(pFolderPath & "SuperMart_Monthly_Sales_Bulletin_2025.pdf"), [Implementation = "1.3"]),
    Table001  = Source{[Id = "Table001"]}[Data],
    Promoted  = Table.PromoteHeaders(Table001, [PromoteAllScalars = true]),
    NoTotal   = Table.SelectRows(Promoted, each [Store] <> "Total"),          // never import a total row as data
    NoFY      = Table.RemoveColumns(NoTotal, {"FY 2025"}),
    Unpivoted = Table.UnpivotOtherColumns(NoFY, {"Store"}, "Month", "Sales_m"),
    Typed     = Table.TransformColumnTypes(Unpivoted, {{"Sales_m", type number}}, "en-GB"),
    MonthNo   = Table.AddColumn(Typed, "Month_No", each List.PositionOf({"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"}, [Month]) + 1, Int64.Type)
in
    MonthNo''')
lng = pdf_tab.drop(columns="FY_2025").melt(id_vars="Store_Name", var_name="Month", value_name="Sales_m")
s.para(f"Result: {len(lng)} tidy rows (12 stores x 12 months), columns Store | Month | Sales_m | Month_No.", "note")

s.band("2.4.4", "Intermediate: join with quality controls - PDF vs the ledger")
s.scenario("Is the PDF trustworthy? Compare each store's FY total in the PDF with the ledger (SuperMart_Stores_Sales_2022_2025.csv). "
           "Rounding to ₦0.1m means small differences are expected; anything above ₦0.1m is an exception.")
led = s25.groupby("Store_Name").Sales.sum().reindex(stores.Store_Name)
rec = pd.DataFrame({"Store_Name": stores.Store_Name, "PDF_FY_m": pdf_tab.FY_2025.values, "Ledger_Sales": led.values})
tr_ = s.table(rec, "tRecon", fmts={"PDF_FY_m": "dec1", "Ledger_Sales": "naira"},
              calc={"Ledger_m": ("=ROUND([@Ledger_Sales]/1000000,1)", "dec1", (rec.Ledger_Sales / 1e6).round(1).tolist()),
                    "Difference_m": ("=[@PDF_FY_m]-[@Ledger_m]", "dec1", (rec.PDF_FY_m - (rec.Ledger_Sales / 1e6).round(1)).round(1).tolist()),
                    "Status": ('=IF(ABS([@Difference_m])<=0.1,"OK","EXCEPTION")', None,
                               ["OK" if abs(x) <= 0.1 + 1e-9 else "EXCEPTION" for x in (rec.PDF_FY_m - (rec.Ledger_Sales / 1e6).round(1))])},
              widths={"Store_Name": 26})
s.calc([("Exceptions", '=COUNTIF(tRecon[Status],"EXCEPTION")', 0, "int", "0 = the PDF agrees with the ledger to rounding"),
        ("Sum of 144 rounded monthly cells vs ledger (₦m)", f"=SUM(tPdfBulletin[[Jan]:[Dec]])-SUM(tRecon[Ledger_Sales])/1000000",
         float(ct.values.sum() - led.sum() / 1e6), "dec2", "Rounding drift: why you never rebuild totals from rounded figures")])

s.band("2.4.5", "Advanced practical: an audit-ready refresh")
s.bullets(["Keep the source file name, page and table id (Table001) in a 'Source' column - provenance travels with the data.",
           "Add a Row_Count query and a Total_Check query; load both to a 99_Checks sheet.",
           "If the bulletin layout changes (e.g. a new store column), the Unpivot step still works - Unpivot OTHER columns is layout-proof.",
           "Never overwrite the ledger with PDF numbers: the PDF is a presentation copy; the ledger is the system of record."])

# ======================================================================================
# 2.5 Data cleaning with Excel  (raw sheet + 40-row worked sample)
# ======================================================================================
s = b.sheet("2.5 Excel Cleaning", "2.5  Data Cleaning with Excel",
            "The twelve data-quality problems in the FY2025 POS extract - detect with formulas, fix, verify", "2.5.1 - 2.5.12")
s.band("2.5.1-2", "Detect: one formula per problem, run over the whole raw extract (sheet D_Dirty_2025)")
s.scenario("tDirty (sheet D_Dirty_2025) holds the raw file exactly as Excel opens it: 3,387 rows. Every count below is a live formula over that table.")
nb = dirty.dropna(how="all")
det = [
    ("Rows in the raw table", "=ROWS(tDirty)", len(dirty), "int", "Includes blank spacer rows"),
    ("Blank spacer rows", "=COUNTBLANK(tDirty[Order_ID])", int(dirty.Order_ID.isna().sum()), "int", "Problem 10: blank rows"),
    ("Empty column (Column1) - non-blank cells", "=COUNTA(tDirty[Column1])", 0, "int", "Problem 10: empty column → delete"),
    ("Duplicate lines (Row_ID seen before)", "=ROWS(tDirty)-COUNTBLANK(tDirty[Row_ID])-SUMPRODUCT((tDirty[Row_ID]<>\"\")/COUNTIF(tDirty[Row_ID],tDirty[Row_ID]&\"\"))",
     len(nb) - nb.Row_ID.nunique(), "int", "Problem 3: duplicates (60)"),
    ("Sales stored as text", "=SUMPRODUCT(--ISTEXT(tDirty[Sales]))", int(dirty.Sales.map(lambda v: isinstance(v, str)).sum()), "int", "Problem 5: numbers as text"),
    ("Discount stored as text ('20%')", "=SUMPRODUCT(--ISTEXT(tDirty[Discount]))", int(dirty.Discount.map(lambda v: isinstance(v, str)).sum()), "int", "Problem 5"),
    ("Order_Date stored as text (not a real date)", "=SUMPRODUCT(--ISTEXT(tDirty[Order_Date]))",
     int(dd.Order_Date.map(lambda v: isinstance(v, str)).sum()), "int", "Problem 8: incorrect dates"),
    ("Zone with extra spaces", "=SUMPRODUCT(--(LEN(tDirty[Zone])<>LEN(TRIM(tDirty[Zone]))))",
     int(nb.Zone.map(lambda v: len(v) != len(" ".join(v.split()))).sum()), "int", "Problem 6: spaces"),
    ("Customer_Name with non-printing characters", "=SUMPRODUCT(--(LEN(tDirty[Customer_Name])<>LEN(CLEAN(SUBSTITUTE(tDirty[Customer_Name],CHAR(160),\"\")))))",
     int(nb.Customer_Name.map(lambda v: ("\xa0" in v) or ("\n" in v)).sum()), "int", "Problem 6: CHAR(160)/line feeds (40)"),
    ("Category not exactly in the official list", '=SUMPRODUCT((tDirty[Category]<>"")*(EXACT(tDirty[Category],"Furniture")+EXACT(tDirty[Category],"Office Supplies")+EXACT(tDirty[Category],"Technology")=0))',
     int((~nb.Category.isin(["Furniture", "Office Supplies", "Technology"])).sum()), "int", "Problem 7: uncontrolled vocabulary. EXACT is case-sensitive; MATCH/COUNTIF are NOT, so they would miss 'FURNITURE'."),
    ("…of which only a case problem", '=SUMPRODUCT((tDirty[Category]<>"")*ISNUMBER(MATCH(tDirty[Category],{"Furniture","Office Supplies","Technology"},0)))-SUMPRODUCT(EXACT(tDirty[Category],"Furniture")+EXACT(tDirty[Category],"Office Supplies")+EXACT(tDirty[Category],"Technology"))',
     int(nb.Category.str.lower().isin(["furniture", "office supplies", "technology"]).sum() - nb.Category.isin(["Furniture", "Office Supplies", "Technology"]).sum()), "int",
     "FURNITURE, furniture, office supplies, technology - PROPER fixes these; misspellings need the mapping table."),
    ("Missing Quantity (on real rows)", '=SUMPRODUCT((tDirty[Order_ID]<>"")*(tDirty[Quantity]=""))', int(nb.Quantity.isna().sum()), "int", "Problem 9: missing values"),
    ("Negative Quantity", '=COUNTIF(tDirty[Quantity],"<0")', int((pd.to_numeric(nb.Quantity, errors="coerce") < 0).sum()), "int", "Problem 11: sign errors"),
    ("Customer_Name = 'N/A'", '=COUNTIF(tDirty[Customer_Name],"N/A")', int((nb.Customer_Name == "N/A").sum()), "int", "Problem 9: placeholder text"),
]
s.calc(det)

s.band("2.5.3-11", "Fix: a 40-row worked sample with one helper column per repair")
samp_idx = []
for cond in [dirty.Sales.map(lambda v: isinstance(v, str)), dirty.Discount.map(lambda v: isinstance(v, str)),
             dirty.Order_Date.map(lambda v: isinstance(v, str) and not re.match(r"^\d{4}-", v or "")),
             dirty.Category.isin(["Furnitre", "Tech", "Office Supply", "FURNITURE", "Technolgy"]),
             dirty.Zone.map(lambda v: isinstance(v, str) and v != v.strip()), dirty.Customer_Name.map(lambda v: isinstance(v, str) and "\xa0" in v),
             dirty.Quantity.isna() & dirty.Order_ID.notna()]:
    samp_idx += dirty[cond].index[:6].tolist()
samp_idx = list(dict.fromkeys(samp_idx))[:40]
sp = dd.loc[samp_idx, ["Row_ID", "Order_Date", "Customer_Name", "Zone", "Category", "Sales", "Quantity", "Discount"]].reset_index(drop=True)
catmap = {"furniture": "Furniture", "furnitre": "Furniture", "office supplies": "Office Supplies", "office supply": "Office Supplies",
          "office-supplies": "Office Supplies", "technology": "Technology", "technolgy": "Technology", "tech": "Technology"}
mon = {m.lower(): i + 1 for i, m in enumerate(MONTHS)}


def fix_date(v):
    if isinstance(v, pd.Timestamp):
        return v
    v = str(v)
    if re.match(r"^\d{2}/\d{2}/\d{4}$", v):
        return pd.Timestamp(int(v[6:]), int(v[3:5]), int(v[:2]))
    if re.match(r"^\d{2}-[A-Za-z]{3}-\d{4}$", v):
        return pd.Timestamp(int(v[7:]), mon[v[3:6].lower()], int(v[:2]))
    if re.match(r"^\d{4}\.\d{2}\.\d{2}$", v):
        return pd.Timestamp(int(v[:4]), int(v[5:7]), int(v[8:]))
    return pd.NaT


def xl_serial(t):
    return (t - pd.Timestamp("1899-12-30")).days if pd.notna(t) else ""


date_f = ('=IF(ISNUMBER([@Order_Date]),[@Order_Date],IF(MID([@Order_Date],3,1)="/",DATE(RIGHT([@Order_Date],4),MID([@Order_Date],4,2),LEFT([@Order_Date],2)),'
          'IF(MID([@Order_Date],5,1)=".",DATE(LEFT([@Order_Date],4),MID([@Order_Date],6,2),RIGHT([@Order_Date],2)),DATEVALUE([@Order_Date]))))')
mp = pd.DataFrame({"Variant": list(catmap.keys()), "Standard": list(catmap.values())})
tcs = s.table(sp, "tFix40", fmts={"Order_Date": "date", "Row_ID": "gen"},
              calc={"Date_Fixed": (date_f, "date", [xl_serial(fix_date(v)) for v in sp.Order_Date]),
                    "Name_Fixed": ('=TRIM(CLEAN(SUBSTITUTE([@Customer_Name],CHAR(160)," ")))', None,
                                   [" ".join(str(v).replace("\xa0", " ").replace("\n", "").split()) for v in sp.Customer_Name]),
                    "Zone_Fixed": ("=TRIM([@Zone])", None, [" ".join(str(v).split()) for v in sp.Zone]),
                    "Category_Fixed": ("=XLOOKUP(LOWER(TRIM([@Category])),tCatMap[Variant],tCatMap[Standard],\"CHECK\")", None,
                                       [catmap.get(str(v).strip().lower(), "CHECK") for v in sp.Category]),
                    "Sales_Fixed": ('=IF(ISNUMBER([@Sales]),[@Sales],VALUE(SUBSTITUTE(SUBSTITUTE([@Sales],"₦",""),",","")))', "naira",
                                    [v if not isinstance(v, str) else float(v.replace("₦", "").replace(",", "")) for v in sp.Sales]),
                    "Discount_Fixed": ('=IF(ISNUMBER([@Discount]),[@Discount],VALUE(SUBSTITUTE([@Discount],"%",""))/100)', "pct0",
                                       [v if not isinstance(v, str) else float(v.replace("%", "")) / 100 for v in sp.Discount]),
                    "Qty_Flag": ('=IF([@Quantity]="","MISSING - recover from POS",IF([@Quantity]<0,"NEGATIVE - confirm","OK"))', None,
                                 ["MISSING - recover from POS" if v is None or (isinstance(v, float) and np.isnan(v)) else ("NEGATIVE - confirm" if v < 0 else "OK") for v in sp.Quantity])},
              widths={"Customer_Name": 18, "Name_Fixed": 18, "Category_Fixed": 15, "Date_Fixed": 11, "Qty_Flag": 24, "Zone": 14, "Zone_Fixed": 13})
s.h2("Mapping table used by Category_Fixed (problem 7 - standardising categorical values)")
s.table(mp, "tCatMap", widths={"Variant": 16, "Standard": 16})
s.warn("Never 'fix' a missing Quantity by typing 0 or the average. Flag it, count it, and recover it from the till roll - the dirty file has 30.")

s.band("2.5.12", "Verify: the cleaned dataset must reproduce the POS control totals")
clean = L("dirty_sm_truth")
s.grid([[k, int(v)] for k, v in ctrl.items()], header=["POS control total (FY2025)", "Value"], fmts=[None, "int"])
s.calc([("Clean rows (sheet D_Sales_2025)", "=ROWS(tSales)", len(s25), "int", ""),
        ("Clean total sales", "=SUM(tSales[Sales])", int(s25.Sales.sum()), "naira", ""),
        ("Clean total profit", "=SUM(tSales[Profit])", int(s25.Profit.sum()), "naira", ""),
        ("Clean total quantity", "=SUM(tSales[Quantity])", int(s25.Quantity.sum()), "int", ""),
        ("Distinct orders", "=COUNTA(UNIQUE(tSales[Order_ID]))", int(s25.Order_ID.nunique()), "int", ""),
        ("ALL CHECKS", f'=IF(AND(ROWS(tSales)={int(ctrl["Clean row count"])},SUM(tSales[Sales])={int(ctrl["Total Sales (₦)"])},SUM(tSales[Profit])={int(ctrl["Total Profit (₦)"])},SUM(tSales[Quantity])={int(ctrl["Total Quantity"])}),"PASS","FAIL")',
         "PASS", None, "One cell that tells a manager the data can be trusted")])

# ======================================================================================
# 2.6 Power Query transforms
# ======================================================================================
s = b.sheet("2.6 PQ Transforms", "2.6  Cleaning and Transforming in Power Query",
            "Twenty transformations - each with the M step and a before/after on our data", "2.6.1 - 2.6.20")
TR = [
    ("2.6.1", "Remove columns", 'Table.RemoveColumns(Source, {"Column1"})', "Raw file: 21 columns incl. empty Column1", "20 columns"),
    ("2.6.2", "Remove blank rows", 'Table.SelectRows(#"Prev", each not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), {"", null})))', "3,387 rows", "3,372 rows"),
    ("2.6.3", "Remove duplicates", 'Table.Distinct(#"Prev")', "3,372 rows", "3,312 rows"),
    ("2.6.4", "Filter rows", 'Table.SelectRows(#"Prev", each [Profit] < 0)', "3,312 lines", f"{int((s25.Profit < 0).sum()):,} loss-making lines"),
    ("2.6.5", "Replace values", 'Table.ReplaceValue(#"Prev", "Furnitre", "Furniture", Replacer.ReplaceText, {"Category"})', "Furnitre (8 rows)", "Furniture"),
    ("2.6.6", "Change type with locale", 'Table.TransformColumnTypes(#"Prev", {{"Order_Date", type date}}, "en-GB")', "'03/04/2025' (text)", "2025-04-03 (date)"),
    ("2.6.7", "Split column by delimiter", 'Table.SplitColumn(#"Prev", "Store_Location", Splitter.SplitTextByDelimiter(", "), {"City", "State"})', "Benin City, Edo", "City = Benin City | State = Edo"),
    ("2.6.8", "Merge columns", 'Table.CombineColumns(#"Prev", {"Store_ID", "Store_Name"}, Combiner.CombineTextByDelimiter(" · "), "Store_Label")', "ST05 | Benin City - Ring Road", "ST05 · Benin City - Ring Road"),
    ("2.6.9", "Trim and clean text", 'Table.TransformColumns(#"Prev", {{"Customer_Name", each Text.Trim(Text.Clean(Text.Replace(_, "#(00A0)", " "))), type text}})', "'Omolara Yusuf' + line feed", "Omolara Yusuf"),
    ("2.6.10", "Extract text", 'Table.AddColumn(#"Prev", "Cat_Code", each Text.Start([Product_ID], 3), type text)', "FUR-FU-10001940", "FUR"),
    ("2.6.11", "Conditional column", 'Table.AddColumn(#"Prev", "Disc_Band", each if [Discount] = 0 then "None" else if [Discount] <= 0.2 then "1-20%" else "Above 20%")', "Discount 0.45", "Above 20%"),
    ("2.6.12", "Custom column", 'Table.AddColumn(#"Prev", "Margin", each [Profit] / [Sales], Percentage.Type)', "Sales 25,472 | Profit 7,642", "Margin 30.0%"),
    ("2.6.13", "Index column", 'Table.AddIndexColumn(#"Prev", "Line_No", 1, 1, Int64.Type)', "No row number", "1, 2, 3 ..."),
    ("2.6.14", "Group By", 'Table.Group(#"Prev", {"Zone"}, {{"Sales", each List.Sum([Sales]), type number}, {"Orders", each List.Count(List.Distinct([Order_ID])), Int64.Type}})', "3,312 lines", "4 zone rows (see table below)"),
    ("2.6.15", "Pivot / unpivot", 'Table.UnpivotOtherColumns(#"Prev", {"Store"}, "Month", "Sales_m")', "12 x 13 cross-tab (PDF)", "144 tidy rows"),
    ("2.6.16", "Handle nulls", 'Table.FillDown(#"Prev", {"Store_Name"})  /  Table.ReplaceValue(..., null, "Not Recorded", ...)', "Blank under a group header", "Value copied down / labelled"),
    ("2.6.17", "Sort and filter", 'Table.Sort(#"Prev", {{"Sales", Order.Descending}})', "Random order", "Largest sale first"),
    ("2.6.18", "Reorder and rename", 'Table.RenameColumns(#"Prev", {{"Sales", "Sales_NGN"}})', "Sales", "Sales_NGN"),
    ("2.6.19", "Applied steps history", "Right-click a step > Rename / Delete Until End / Move Before", "Step1, Step2 ...", "Named steps: Removed Blank Rows ..."),
    ("2.6.20", "Refresh after source change", "Data > Refresh All", "Jan-Nov files", "Dec file added → totals update"),
]
s.table(pd.DataFrame(TR, columns=["Topic", "Transform", "M step (formula bar)", "Before (our data)", "After"]), "tTransforms",
        widths={"Topic": 8, "Transform": 20, "M step (formula bar)": 70, "Before (our data)": 26, "After": 28})
s.band("2.6.14", "Group By result - Zone summary for FY2025 (what the query returns)")
gz = s25.groupby("Zone").agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"), Orders=("Order_ID", "nunique"), Lines=("Row_ID", "count")).reset_index()
tgz = s.table(gz, "tGroupZone", fmts={"Sales": "naira", "Profit": "naira", "Orders": "int", "Lines": "int"},
              calc={"Margin": ("=[@Profit]/[@Sales]", "pct", (gz.Profit / gz.Sales).tolist())},
              total={"Zone": "Total", "Sales": "sum", "Profit": "sum", "Orders": "sum", "Lines": "sum"})
s.calc([("Check: Group By total = tSales total", '=IF(SUM(tGroupZone[Sales])=SUM(tSales[Sales]),"PASS","FAIL")', "PASS", None, "Every group-by needs a reconciliation")])
s.band("2.6.11", "Conditional column result - discount band (FY2025)")
s25["Disc_Band"] = np.where(s25.Discount == 0, "None", np.where(s25.Discount <= 0.2, "1-20%", "Above 20%"))
dbz = s25.groupby("Disc_Band").agg(Lines=("Row_ID", "count"), Sales=("Sales", "sum"), Profit=("Profit", "sum")).reindex(["None", "1-20%", "Above 20%"]).reset_index()
s.table(dbz, "tDiscBand25", fmts={"Lines": "int", "Sales": "naira", "Profit": "naira"},
        calc={"Margin": ("=[@Profit]/[@Sales]", "pct", (dbz.Profit / dbz.Sales).tolist())})
s.insight(f"Lines discounted above 20% lost ₦{-dbz.Profit.iloc[2] / 1e6:,.1f}m in FY2025 - the same conclusion as Week 1, now produced by a refreshable query.")

# ======================================================================================
# 2.7 Excel vs Power Query
# ======================================================================================
s = b.sheet("2.7 Excel vs PQ", "2.7  Excel Cleaning versus Power Query Cleaning",
            "The same FY2025 clean-up done both ways - effort, repeatability and audit trail", "2.7.1 - 2.7.6")
s.band("2.7.6", "Practical comparison on the same dataset")
cmp_ = pd.DataFrame([
    ["Remove 15 blank rows", "Go To Special > Blanks > Delete rows", 1, "Removed Blank Rows step", 1],
    ["Remove 60 duplicates", "Data > Remove Duplicates", 1, "Removed Duplicates step", 1],
    ["Fix 234 text Sales values", "Helper column + Paste Values", 3, "Replace ₦ and ',' then Change Type", 2],
    ["Fix 331 text dates (3 formats)", "Nested IF/DATE helper column", 4, "Change Type with Locale", 1],
    ["Standardise 132 categories", "Mapping table + XLOOKUP", 3, "Merge with mapping query", 2],
    ["Split Store_Location", "Text to Columns", 1, "Split Column by Delimiter", 1],
    ["Trim/clean names and zones", "TRIM/CLEAN helper columns", 2, "Transform > Format > Trim / Clean", 1],
    ["Next month's file arrives", "Repeat everything by hand", 15, "Refresh", 0]],
    columns=["Task", "Excel method", "Excel_steps", "Power Query method", "PQ_steps"])
tcmp = s.table(cmp_, "tCompare", fmts={"Excel_steps": "int", "PQ_steps": "int"}, total={"Task": "Total", "Excel_steps": "sum", "PQ_steps": "sum"},
               widths={"Task": 28, "Excel method": 30, "Power Query method": 32})
s.calc([("Manual steps per year (12 monthly files)", "=12*SUM(tCompare[Excel_steps])", 12 * cmp_.Excel_steps.sum(), "int", "Each one a chance for error"),
        ("Power Query steps per year", "=SUM(tCompare[PQ_steps])", cmp_.PQ_steps.sum(), "int", "Built once, refreshed 12 times")])
s.grid([["Reproducibility", "Depends on the person and their notes", "Identical result every refresh"],
        ["Auditability", "Hard to prove what was changed", "Applied Steps show every change, in order"],
        ["Speed on 3,000 rows", "Fine", "Fine"],
        ["Speed on 1,000,000+ rows", "Sheet limit / slow", "Streams data; load to Data Model"],
        ["Best for", "One-off fixes, quick look, small lists", "Anything you will do twice"]], header=["Criterion", "Manual Excel", "Power Query"])

# ======================================================================================
# 2.8 The 65 functions (Part A + Part B)
# ======================================================================================
sA = b.sheet("2.8A Functions 1-33", "2.8  The 65 Analytical Excel Functions - Part A (1-33)",
             "Aggregation, statistics, rounding, logic and text output - all on tSales (FY2025, sheet D_Sales_2025)", "2.8.1 - 2.8.33")
x = s25
st = stores_x
tech = x[x.Category == "Technology"]
sw = x[x.Zone == "South-West"]
m_chain = x.Profit.sum() / x.Sales.sum()
ik = st.set_index("Store_ID").loc["ST01"]
ow = st.set_index("Store_ID").loc["ST12"]
sA.band("2.8.1-12", "Aggregation and counting")
sA.calc([
    ("2.8.1 SUM - FY2025 chain sales", "=SUM(tSales[Sales])", x.Sales.sum(), "naira", "Adds every number in the column."),
    ("2.8.2 SUMIF - Technology sales", '=SUMIF(tSales[Category],"Technology",tSales[Sales])', tech.Sales.sum(), "naira", "One condition: range, criteria, sum_range."),
    ("2.8.3 SUMIFS - Technology sales in South-West", '=SUMIFS(tSales[Sales],tSales[Category],"Technology",tSales[Zone],"South-West")',
     tech[tech.Zone == "South-West"].Sales.sum(), "naira", "Many conditions - sum_range comes FIRST."),
    ("2.8.4 AVERAGE - average sales line", "=AVERAGE(tSales[Sales])", x.Sales.mean(), "naira", "Mean of all lines."),
    ("2.8.5 AVERAGEIFS - avg Corporate line at Lagos - Ikeja", '=AVERAGEIFS(tSales[Sales],tSales[Segment],"Corporate",tSales[Store_ID],"ST01")',
     x[(x.Segment == "Corporate") & (x.Store_ID == "ST01")].Sales.mean(), "naira", ""),
    ("2.8.6 COUNT - numeric Quantity cells", "=COUNT(tSales[Quantity])", int(x.Quantity.count()), "int", "Counts numbers only."),
    ("2.8.7 COUNTA - order lines (non-blank IDs)", "=COUNTA(tSales[Order_ID])", len(x), "int", "Counts anything non-blank."),
    ("2.8.8 COUNTBLANK - blank Quantity in the RAW file", "=COUNTBLANK(tDirty[Quantity])", int(dirty.Quantity.isna().sum()), "int", "31 missing + 15 spacer rows = 46."),
    ("2.8.9 COUNTIF - loss-making lines", '=COUNTIF(tSales[Profit],"<0")', int((x.Profit < 0).sum()), "int", ""),
    ("2.8.10 COUNTIFS - Furniture losses at ≥30% discount", '=COUNTIFS(tSales[Category],"Furniture",tSales[Profit],"<0",tSales[Discount],">=0.3")',
     int(((x.Category == "Furniture") & (x.Profit < 0) & (x.Discount >= 0.3)).sum()), "int", ""),
    ("2.8.11 SUBTOTAL - visible-rows sales (filter the data sheet!)", "=SUBTOTAL(109,tSales[Sales])", x.Sales.sum(), "naira", "109 = SUM ignoring hidden/filtered rows."),
    ("2.8.12 AGGREGATE - largest South-East sale", '=AGGREGATE(14,6,tSales[Sales]/(tSales[Zone]="South-East"),1)',
     x[x.Zone == "South-East"].Sales.max(), "naira", "14 = LARGE, 6 = ignore errors; dividing by FALSE makes #DIV/0! which is skipped.")])
sA.band("2.8.13-22", "Statistics and rounding")
sA.calc([
    ("2.8.13 MEDIAN - middle sales line", "=MEDIAN(tSales[Sales])", float(x.Sales.median()), "naira", "Far below the mean - sales are right-skewed."),
    ("2.8.14 MODE.SNGL - most common quantity", "=MODE.SNGL(tSales[Quantity])", int(x.Quantity.mode().iloc[0]), "int", ""),
    ("2.8.15 MAX - biggest single line", "=MAX(tSales[Sales])", x.Sales.max(), "naira", ""),
    ("2.8.16 MIN - worst loss on one line", "=MIN(tSales[Profit])", x.Profit.min(), "naira", ""),
    ("2.8.17 LARGE - 3rd largest line", "=LARGE(tSales[Sales],3)", float(np.sort(x.Sales.values)[-3]), "naira", ""),
    ("2.8.18 SMALL - 3rd worst profit", "=SMALL(tSales[Profit],3)", float(np.sort(x.Profit.values)[2]), "naira", ""),
    ("2.8.19 ROUND - avg line to nearest ₦1,000", "=ROUND(AVERAGE(tSales[Sales]),-3)", round(x.Sales.mean(), -3), "naira", "Negative digits round to the left of the point."),
    ("2.8.20 ROUNDUP - cartons for Technology units (12 per carton)", '=ROUNDUP(SUMIFS(tSales[Quantity],tSales[Category],"Technology")/12,0)',
     float(np.ceil(tech.Quantity.sum() / 12)), "int", "Always up - you cannot ship part of a carton."),
    ("2.8.21 ROUNDDOWN - chain margin to 3 dp", "=ROUNDDOWN(SUM(tSales[Profit])/SUM(tSales[Sales]),3)", np.floor(m_chain * 1000) / 1000, "dec3", ""),
    ("2.8.22 ABS - worst loss as a positive amount", "=ABS(MIN(tSales[Profit]))", abs(x.Profit.min()), "naira", "")])
sA.band("2.8.23-30", "Logic")
att_ik = ik.Sales_2025 / ik.Target_2025
att_ow = ow.Sales_2025 / ow.Target_2025
m_ow = ow.Profit_2025 / ow.Sales_2025
sA.calc([
    ("2.8.23 IF - did the chain meet its total target?", '=IF(SUM(tSales[Sales])>=SUM(tStores[Target_2025]),"Met","Missed")',
     "Met" if x.Sales.sum() >= st.Target_2025.sum() else "Missed", None, f"Targets total ₦{st.Target_2025.sum() / 1e6:,.0f}m"),
    ("2.8.24 IFS - chain margin band", '=IFS(SUM(tSales[Profit])/SUM(tSales[Sales])>=0.15,"Strong",SUM(tSales[Profit])/SUM(tSales[Sales])>=0.08,"Healthy",TRUE,"Weak")',
     "Strong" if m_chain >= .15 else ("Healthy" if m_chain >= .08 else "Weak"), None, "First TRUE test wins."),
    ("2.8.25 AND - Ikeja met target AND margin above 10%?",
     '=AND(SUMIFS(tSales[Sales],tSales[Store_ID],"ST01")>=XLOOKUP("ST01",tStores[Store_ID],tStores[Target_2025]),SUMIFS(tSales[Profit],tSales[Store_ID],"ST01")/SUMIFS(tSales[Sales],tSales[Store_ID],"ST01")>0.1)',
     bool(att_ik >= 1 and ik.Profit_2025 / ik.Sales_2025 > 0.1), None, ""),
    ("2.8.26 OR - Owerri needs review (attainment<85% OR margin<5%)?",
     '=OR(SUMIFS(tSales[Sales],tSales[Store_ID],"ST12")/XLOOKUP("ST12",tStores[Store_ID],tStores[Target_2025])<0.85,SUMIFS(tSales[Profit],tSales[Store_ID],"ST12")/SUMIFS(tSales[Sales],tSales[Store_ID],"ST12")<0.05)',
     bool(att_ow < .85 or m_ow < .05), None, f"Owerri attainment {att_ow:.0%}"),
    ("2.8.27 NOT - are there any loss-making lines?", '=NOT(COUNTIF(tSales[Profit],"<0")=0)', bool((x.Profit < 0).any()), None, "Reverses TRUE/FALSE."),
    ("2.8.28 IFERROR - margin for store ST13 (does not exist)", '=IFERROR(SUMIFS(tSales[Profit],tSales[Store_ID],"ST13")/SUMIFS(tSales[Sales],tSales[Store_ID],"ST13"),"No sales")',
     "No sales", None, "0/0 would be #DIV/0!"),
    ("2.8.29 IFNA - manager of store ST13", '=IFNA(VLOOKUP("ST13",tStores[[Store_ID]:[Store_Manager]],8,FALSE),"Unknown store")', "Unknown store", None, "Catches #N/A only - other errors still show."),
    ("2.8.30 SWITCH - category from a product code", '=SWITCH(LEFT(INDEX(tSales[Product_ID],1),3),"FUR","Furniture","OFF","Office Supplies","TEC","Technology","Unknown")',
     {"FUR": "Furniture", "OFF": "Office Supplies", "TEC": "Technology"}[x.Product_ID.iloc[0][:3]], None, f"First line's Product_ID is {x.Product_ID.iloc[0]}")])
sA.band("2.8.31-33", "Text output")
se_st = st[st.Zone == "South-East"].Store_Name.tolist()
sA.calc([
    ("2.8.31 TEXT - a sentence for a report", '="FY2025 sales: "&TEXT(SUM(tSales[Sales]),"₦#,##0")', f"FY2025 sales: ₦{x.Sales.sum():,.0f}", None, "Formats a number inside text."),
    ("2.8.32 CONCAT - store label", '=CONCAT(INDEX(tStores[Store_Name],1)," (",INDEX(tStores[Store_ID],1),")")', f"{st.Store_Name.iloc[0]} ({st.Store_ID.iloc[0]})", None, "Joins pieces; & does the same."),
    ("2.8.33 TEXTJOIN - South-East stores in one cell", '=TEXTJOIN(", ",TRUE,FILTER(tStores[Store_Name],tStores[Zone]="South-East"))', ", ".join(se_st), None, "Delimiter, ignore blanks, values.")])

sB = b.sheet("2.8B Functions 34-65", "2.8  The 65 Analytical Excel Functions - Part B (34-65)",
             "Text parsing, dates, lookups and dynamic arrays on real IDs, names and dates", "2.8.34 - 2.8.65")
sB.band("2.8.34-39", "Parsing codes: SuperMart order and product IDs")
cd_ = x.drop_duplicates("Product_ID").head(6)[["Order_ID", "Product_ID", "Product_Name"]].reset_index(drop=True)
cd_.loc[1, "Product_Name"] = x[x.Product_Name.str.contains("Chair")].Product_Name.iloc[0]
cd_.loc[3, "Product_Name"] = x[x.Product_Name.str.contains("chair", case=False)].Product_Name.iloc[1]


def _search(sub, s_):
    p = s_.lower().find(sub.lower())
    return p + 1 if p >= 0 else "not found"


sB.table(cd_, "tCodes",
         calc={"LEFT_CatCode": ("=LEFT([@Product_ID],3)", None, [v[:3] for v in cd_.Product_ID]),
               "RIGHT_OrderNo": ("=RIGHT([@Order_ID],6)", None, [v[-6:] for v in cd_.Order_ID]),
               "MID_Year": ("=MID([@Order_ID],4,4)", None, [v[3:7] for v in cd_.Order_ID]),
               "LEN_ProdID": ("=LEN([@Product_ID])", "int", [len(v) for v in cd_.Product_ID]),
               "FIND_2nd_dash": ('=FIND("-",[@Product_ID],FIND("-",[@Product_ID])+1)', "int", [v.find("-", v.find("-") + 1) + 1 for v in cd_.Product_ID]),
               "SEARCH_chair": ('=IFERROR(SEARCH("chair",[@Product_Name]),"not found")', None, [_search("chair", v) for v in cd_.Product_Name])},
         widths={"Product_Name": 34, "Order_ID": 15, "Product_ID": 16, "SEARCH_chair": 12, "FIND_2nd_dash": 12})
sB.para("LEFT/RIGHT/MID extract by position; FIND is case-sensitive, SEARCH is not (it finds 'Chair' and 'chair').", "note")
sB.band("2.8.40-46", "Cleaning names and contacts: SuperMart Online customers")
cu = o["customers"].head(6)[["customer_id", "first_name", "last_name", "email", "phone_number"]].copy()
raw_names = ["  " + cu.first_name.iloc[0].upper() + "   " + cu.last_name.iloc[0].upper(),
             cu.first_name.iloc[1].lower() + "\u00a0" + cu.last_name.iloc[1].lower(),
             cu.first_name.iloc[2] + " " + cu.last_name.iloc[2] + "\n",
             cu.first_name.iloc[3].upper() + " " + cu.last_name.iloc[3].lower() + " ",
             cu.first_name.iloc[4] + "  " + cu.last_name.iloc[4],
             cu.first_name.iloc[5].lower() + " " + cu.last_name.iloc[5].upper()]
cn = pd.DataFrame({"customer_id": cu.customer_id, "Name_as_keyed": raw_names, "email": [e.upper() if i % 2 == 0 else e for i, e in enumerate(cu.email)],
                   "phone_number": cu.phone_number.astype(str)})


def _tclean(v):
    return " ".join(v.replace("\u00a0", " ").replace("\n", "").split())


sB.table(cn, "tCustNames",
         calc={"SUBSTITUTE_nbsp": ('=SUBSTITUTE([@Name_as_keyed],CHAR(160)," ")', None, [v.replace("\u00a0", " ") for v in cn.Name_as_keyed]),
               "CLEAN": ("=CLEAN([@SUBSTITUTE_nbsp])", None, [v.replace("\u00a0", " ").replace("\n", "") for v in cn.Name_as_keyed]),
               "TRIM": ("=TRIM([@CLEAN])", None, [_tclean(v) for v in cn.Name_as_keyed]),
               "PROPER": ("=PROPER([@TRIM])", None, [_tclean(v).title() for v in cn.Name_as_keyed]),
               "UPPER_surname": ('=UPPER(TEXTAFTER([@TRIM]," ",-1))', None, [_tclean(v).split(" ")[-1].upper() for v in cn.Name_as_keyed]),
               "LOWER_email": ("=LOWER([@email])", None, [e.lower() for e in cn.email]),
               "REPLACE_masked_phone": ('=REPLACE([@phone_number],5,4,"****")', None, [p[:4] + "****" + p[8:] for p in cn.phone_number])},
         widths={"Name_as_keyed": 22, "email": 24, "LOWER_email": 24, "PROPER": 18, "TRIM": 18, "CLEAN": 18, "SUBSTITUTE_nbsp": 18, "REPLACE_masked_phone": 16})
sB.para("Order matters: SUBSTITUTE the non-breaking space (CHAR 160) first, CLEAN removes line feeds, TRIM collapses spaces, PROPER fixes case. "
        "REPLACE masks a phone number before sharing - a data-protection habit (4.13).", "note")
sB.band("2.8.47-49", "Text to numbers and splitting: the dirty POS values")
tv = dirty[dirty.Sales.map(lambda v: isinstance(v, str))].head(5)[["Row_ID", "Sales", "Store_Location"]].reset_index(drop=True)
tv.columns = ["Row_ID", "Sales_as_text", "Store_Location"]
sB.table(tv, "tTextVals", fmts={"Row_ID": "gen"},
         calc={"VALUE_Sales": ('=VALUE(SUBSTITUTE(SUBSTITUTE([@Sales_as_text],"₦",""),",",""))', "naira",
                               [float(v.replace("₦", "").replace(",", "")) for v in tv.Sales_as_text]),
               "TEXTBEFORE_City": ('=TEXTBEFORE([@Store_Location],", ")', None, [v.split(", ")[0] for v in tv.Store_Location]),
               "TEXTAFTER_State": ('=TEXTAFTER([@Store_Location],", ")', None, [v.split(", ")[1] for v in tv.Store_Location])},
         widths={"Sales_as_text": 16, "Store_Location": 24, "TEXTBEFORE_City": 16})
sB.band("2.8.50-56", "Dates: Lifecare stays and SuperMart shipping")
a6 = lc[lc.Admission_Date.dt.year == 2024].head(6)[["Admission_ID", "Department", "Admission_Date", "Discharge_Date"]].reset_index(drop=True)
sB.table(a6, "tStays", fmts={"Admission_Date": "date", "Discharge_Date": "date"},
         calc={"DATEDIF_days": ('=DATEDIF([@Admission_Date],[@Discharge_Date],"d")', "int", (a6.Discharge_Date - a6.Admission_Date).dt.days.tolist()),
               "EOMONTH_bill_month_end": ("=EOMONTH([@Discharge_Date],0)", "date",
                                          [xl_serial(d + pd.offsets.MonthEnd(0)) for d in a6.Discharge_Date]),
               "WORKDAY_follow_up": ("=WORKDAY([@Discharge_Date],7)", "date",
                                     [xl_serial(pd.Timestamp(np.busday_offset(d.date(), 7, roll="backward"))) for d in a6.Discharge_Date])},
         widths={"Department": 22, "Admission_ID": 15, "EOMONTH_bill_month_end": 14, "WORKDAY_follow_up": 13})
sh6 = x.head(6)[["Order_ID", "Order_Date", "Ship_Date", "Ship_Mode"]].reset_index(drop=True)
sB.table(sh6, "tShip", fmts={"Order_Date": "date", "Ship_Date": "date"},
         calc={"NETWORKDAYS_to_ship": ("=NETWORKDAYS([@Order_Date],[@Ship_Date])", "int",
                                       [int(np.busday_count(a.date(), (bb + pd.Timedelta(days=1)).date())) for a, bb in zip(sh6.Order_Date, sh6.Ship_Date)])},
         widths={"Order_ID": 15, "Ship_Mode": 15})
sB.calc([("2.8.50 TODAY - today's date", "=TODAY()", xl_serial(pd.Timestamp(TODAY)), "date", "Volatile: changes every day the file opens."),
         ("2.8.51 NOW - date and time", "=NOW()", xl_serial(pd.Timestamp(TODAY)), "yyyy-mm-dd hh:mm", "Use for 'last refreshed' stamps."),
         ("2.8.52 DATE - FY2025 year end", "=DATE(2025,12,31)", xl_serial(pd.Timestamp(2025, 12, 31)), "date", "Builds a real date from parts."),
         ("Days since FY2025 year end", "=TODAY()-DATE(2025,12,31)", (pd.Timestamp(TODAY) - pd.Timestamp(2025, 12, 31)).days, "int", "Dates are numbers - subtract them.")])
sB.para("Table columns above: 2.8.53 DATEDIF (length of stay), 2.8.54 EOMONTH (HMO billing month-end), 2.8.56 WORKDAY (7 working days to the follow-up call), "
        "2.8.55 NETWORKDAYS (working days order-to-ship, both ends counted).", "note")

sB.band("2.8.57-60", "Lookups: type a store ID")
lk = sB.inputs([("Store_ID to look up", "ST09", "text", "Try ST02, ST12 - or ST13 to see the not-found handling")])
key = lk["Store_ID to look up"]
row9 = st.set_index("Store_ID").loc["ST09"]
sB.calc([("2.8.57 VLOOKUP - store manager", f"=VLOOKUP({key},tStores[[Store_ID]:[Store_Manager]],8,FALSE)", row9.Store_Manager, None, "Column 8 counted from Store_ID; FALSE = exact match."),
         ("2.8.58 XLOOKUP - FY2025 target", f'=XLOOKUP({key},tStores[Store_ID],tStores[Target_2025],"not found")', row9.Target_2025, "naira", "Looks left or right; built-in not-found text."),
         ("2.8.60 MATCH - position in the list", f"=MATCH({key},tStores[Store_ID],0)", int(st.index[st.Store_ID == "ST09"][0]) + 1, "int", ""),
         ("2.8.59 INDEX - city at that position", f"=INDEX(tStores[City],MATCH({key},tStores[Store_ID],0))", row9.City, None, "INDEX+MATCH: the classic two-step lookup."),
         ("Attainment for this store", f"=XLOOKUP({key},tStores[Store_ID],tStores[Sales_2025])/XLOOKUP({key},tStores[Store_ID],tStores[Target_2025])",
          row9.Sales_2025 / row9.Target_2025, "pct", "")])
sB.band("2.8.61-64", "Dynamic arrays - results spill into neighbouring cells")
z_in = sB.inputs([("Zone for FILTER", "South-South", "text", "Type another zone and the list below re-spills")])["Zone for FILTER"]
sB.h2("2.8.61 FILTER - stores in the chosen zone, with FY2025 sales")
f_rows = st[st.Zone == "South-South"][["Store_Name", "Sales_2025"]].values.tolist()
sB.ws.write(sB.r, 1, "Store_Name", F["hdr"]); sB.ws.write(sB.r, 2, "Sales_2025", F["hdr"]); sB.r += 1
spill(sB, sB.r, 1, f"=FILTER(tStores[[Store_Name]:[Store_Name]],tStores[Zone]={z_in})", [[r_[0]] for r_ in f_rows])
spill(sB, sB.r, 2, f"=FILTER(tStores[Sales_2025],tStores[Zone]={z_in})", [[r_[1]] for r_ in f_rows], ["naira"])
sB.ws.set_column(1, 1, 24)
sB.r += 6
c0 = sB.r
sB.h2("2.8.62 UNIQUE - zones   |   2.8.63 SORT - stores by sales, high to low   |   2.8.64 SORTBY - store names by attainment")
sB.ws.write(sB.r, 1, "UNIQUE(Zone)", F["hdr"]); sB.ws.write(sB.r, 3, "SORT: Store", F["hdr"]); sB.ws.write(sB.r, 4, "SORT: Sales", F["hdr"])
sB.ws.write(sB.r, 6, "SORTBY: Store", F["hdr"]); sB.r += 1
uz = list(dict.fromkeys(st.Zone))
spill(sB, sB.r, 1, "=UNIQUE(tStores[Zone])", [[z] for z in uz])
srt = st.sort_values("Sales_2025", ascending=False)[["Store_Name", "Sales_2025"]].values.tolist()
spill(sB, sB.r, 3, "=SORT(CHOOSE({1,2},tStores[Store_Name],tStores[Sales_2025]),2,-1)", srt, [None, "naira"])
sby = st.assign(a=st.Sales_2025 / st.Target_2025).sort_values("a", ascending=False).Store_Name.tolist()
spill(sB, sB.r, 6, "=SORTBY(tStores[Store_Name],tStores[Sales_2025]/tStores[Target_2025],-1)", [[v] for v in sby])
sB.ws.set_column(3, 3, 24); sB.ws.set_column(6, 6, 24)
sB.r += 13
sB.band("2.8.65", "LET - name the pieces of a long formula")
sB.calc([("Chain margin with LET", "=LET(s,SUM(tSales[Sales]),p,SUM(tSales[Profit]),p/s)", m_chain, "pct", "s and p are computed once and reused."),
         ("Gap to total target with LET", "=LET(actual,SUM(tSales[Sales]),target,SUM(tStores[Target_2025]),gap,actual-target,IF(gap>=0,\"Ahead by \",\"Behind by \")&TEXT(ABS(gap),\"₦#,##0\"))",
          ("Ahead by " if x.Sales.sum() >= st.Target_2025.sum() else "Behind by ") + f"₦{abs(x.Sales.sum() - st.Target_2025.sum()):,.0f}", None, "Readable, and faster on big tables.")])

# ======================================================================================
# 2.9 M language
# ======================================================================================
s = b.sheet("2.9 M Language", "2.9  Introduction to the Power Query M Language",
            "let/in, steps, types, and complete queries for the three datasets", "2.9.1 - 2.9.11")
s.band("2.9.3-5", "Syntax and the step chain - every step names the previous one")
s.code('''let
    Source        = Csv.Document(File.Contents(pFolderPath & "Lifecare_Admissions_2023_DIRTY.csv"), [Delimiter = ",", Encoding = 65001]),
    Promoted      = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    #"Proper Names" = Table.TransformColumns(Promoted, {{"Patient_Name", Text.Proper, type text}}),
    Result        = #"Proper Names"
in
    Result
// Names with spaces need #"...". M is CASE-SENSITIVE: Text.Proper works, text.proper fails.''')
s.band("2.9.4", "M data types you will meet in our files")
s.grid([["type text", "Patient_Name, Store_ID", '"ST01"'], ["Int64.Type", "Quantity, Length_of_Stay", "17"], ["type number", "Discount, BMI", "0.2"],
        ["Currency.Type", "Billing_Amount (4 decimal places fixed)", "1846591.57"], ["type date", "Admission_Date", "#date(2023, 1, 1)"],
        ["type logical", "Is_Weekend flag", "true"], ["null", "Missing Discharge_Date", "null (not 0, not \"\")"],
        ["list / record / table", "Folder.Files output, a row, a whole query", '{1,2,3} / [A=1] / #table(...)']], header=["M type", "Example column", "Literal"])
s.band("2.9.11", "Practical M across the three datasets (paste into Advanced Editor)")
s.h2("A. Lifecare - clean the 2023 dirty admissions extract")
s.code('''let
    Source    = Csv.Document(File.Contents(pFolderPath & "Lifecare_Admissions_2023_DIRTY.csv"), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Promoted  = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Dedup     = Table.Distinct(Promoted),
    Names     = Table.TransformColumns(Dedup, {{"Patient_Name", each Text.Proper(Text.Trim(_)), type text}}),
    WardMap   = Table.TransformColumns(Names, {{"Ward", each
                    let w = Text.Lower(Text.Trim(Text.Replace(Text.Replace(_, "  ", " "), "Saphire", "Sapphire"))) in
                    if Text.StartsWith(w, "emerald") then "Emerald Ward" else if Text.StartsWith(w, "sapphire") then "Sapphire Ward"
                    else if Text.StartsWith(w, "ruby") then "Ruby Ward" else if Text.StartsWith(w, "topaz") then "Topaz Ward" else "CHECK", type text}}),
    Gender    = Table.TransformColumns(WardMap, {{"Gender", each if List.Contains({"m", "male"}, Text.Lower(_)) then "Male" else "Female", type text}}),
    AdmType   = Table.TransformColumns(Gender, {{"Admission_Type", each
                    let t = Text.Lower(Text.Trim(_)) in if Text.StartsWith(t, "emerg") then "Emergency" else Text.Proper(t), type text}}),
    Bill      = Table.TransformColumnTypes(AdmType, {{"Billing_Amount", type number}}, "en-US"),   // '1,234,567.89' parses with en-US
    Dates     = Table.TransformColumnTypes(Bill, {{"Admission_Date", type date}, {"Discharge_Date", type date}, {"Age", Int64.Type}}),
    LOS       = Table.AddColumn(Dates, "Length_of_Stay", each if [Discharge_Date] = null then null else Duration.Days([Discharge_Date] - [Admission_Date]), Int64.Type),
    Exception = Table.AddColumn(LOS, "Exception", each
                    if [Billing_Amount] < 0 then "Negative bill" else if [Discharge_Date] = null then "No discharge date"
                    else if [Length_of_Stay] < 0 then "Discharge before admission" else if [Age] = 0 or [Age] > 120 then "Impossible age" else null, type text),
    Clean     = Table.SelectRows(Exception, each [Exception] = null)
in
    Clean''')
s.h2("B. Crestview - attendance band and at-risk flag")
s.code('''let
    Source  = Csv.Document(File.Contents(pFolderPath & "..\\00_Core_Datasets\\Crestview_Student_Performance.csv"), [Delimiter = ",", Encoding = 65001]),
    Promote = Table.PromoteHeaders(Source),
    Typed   = Table.TransformColumnTypes(Promote, {{"Attendance_Pct", type number}, {"Exam_Score", Int64.Type}, {"Hours_Studied_Per_Week", Int64.Type}}),
    Band    = Table.AddColumn(Typed, "Attendance_Band2", each
                if [Attendance_Pct] >= 90 then "90%+" else if [Attendance_Pct] >= 75 then "75-89%" else "Below 75%", type text),
    Risk    = Table.AddColumn(Band, "Needs_Mentor", each [Exam_Score] < 60 and [Attendance_Pct] < 75, type logical)
in
    Risk''')
s.h2("C. SuperMart - FY2025 sales by store and month (Group By)")
s.code('''let
    Source  = qSales_Clean,                                     // reference the cleaned query (2.2.9)
    Month   = Table.AddColumn(Source, "Month", each Date.ToText([Order_Date], "yyyy-MM"), type text),
    Grouped = Table.Group(Month, {"Store_Name", "Month"}, {{"Sales", each List.Sum([Sales]), type number},
                                                         {"Profit", each List.Sum([Profit]), type number},
                                                         {"Orders", each List.Count(List.Distinct([Order_ID])), Int64.Type}}),
    Margin  = Table.AddColumn(Grouped, "Margin", each [Profit] / [Sales], Percentage.Type)
in
    Margin''')
s.band("2.9.8-10", "Function families you will use most")
s.grid([["Text.*", "Text.Trim, Text.Proper, Text.Start, Text.BeforeDelimiter, Text.Replace", 'Text.BeforeDelimiter("Benin City, Edo", ",") = "Benin City"'],
        ["Number.*", "Number.Round, Number.FromText, Number.Abs", 'Number.Round(0.12749, 3) = 0.127'],
        ["Date.*", "Date.Year, Date.Month, Date.EndOfMonth, Date.DayOfWeekName", "Date.EndOfMonth(#date(2024,2,10)) = #date(2024,2,29)"],
        ["List.*", "List.Sum, List.Distinct, List.Contains, List.PositionOf", 'List.Contains({"m","male"}, "male") = true'],
        ["Table.*", "Table.SelectRows, Table.Group, Table.NestedJoin, Table.UnpivotOtherColumns", "Whole-table operations - most applied steps"]],
       header=["Family", "Most used", "Example (our data)"])

# ======================================================================================
# 2.10 PivotTables (formula twins + live pivots)
# ======================================================================================
s = b.sheet("2.10 PivotTables", "2.10  Introduction to PivotTables",
            "Requirements, the four areas, aggregation, grouping, Show Values As - with formula twins you can check", "2.10.1 - 2.10.19")
s.band("2.10.3", "PivotTable data requirements - is tSM ready?")
s.grid([["One header row, unique names", "✔ 24 headers, no blanks", "Pivot field list uses them"],
        ["No blank rows or columns", "✔ cleaned in 2.5/2.6", "A blank row truncates the source"],
        ["One record per row", "✔ one sales line", "Totals come out right"],
        ["Consistent types per column", "✔ Sales numeric, dates real dates", "Otherwise 'Count of Sales' appears instead of Sum"],
        ["Source is an Excel Table", "✔ tSM (sheet D_SuperMart_All)", "New rows are picked up on Refresh"]], header=["Requirement", "tSM status", "Why it matters"])
s.para("LIVE PIVOTTABLES: sheet '2.10 Live Pivots' holds real PivotTables built on tSM - Excel builds them when the file opens (Refresh All if blank). "
       "The formula twins below give identical numbers, so you can check a pivot - and see why the pivot is faster to build.", "insight")
s.band("2.10.4-9", "Rows, Columns, Values: Sales by Zone (rows) and Year (columns)")
zy = sm.pivot_table(index="Zone", columns="Year", values="Sales", aggfunc="sum")
years = [2022, 2023, 2024, 2025]
s.ws.write(s.r, 1, "Sum of Sales", F["hdr"])
for j, y in enumerate(years):
    s.ws.write_number(s.r, 2 + j, y, F["hdr"])
s.ws.write(s.r, 6, "Grand Total", F["hdr"])
hdr_r = s.r
s.r += 1
z_first = s.r
for z in zy.index:
    s.ws.write(s.r, 1, z, F["lab"])
    for j, y in enumerate(years):
        s.ws.write_formula(s.r, 2 + j, f'=SUMIFS(tSM[Sales],tSM[Zone],$B{s.r + 1},tSM[Year],{s.addr(hdr_r, 2 + j)})', b.nf("naira"), float(zy.loc[z, y]))
    s.ws.write_formula(s.r, 6, f"=SUM(C{s.r + 1}:F{s.r + 1})", b.nf("naira", bold=True), float(zy.loc[z].sum()))
    s.r += 1
s.ws.write(s.r, 1, "Grand Total", F["lab"])
for j in range(5):
    col = chr(67 + j)
    s.ws.write_formula(s.r, 2 + j, f"=SUM({col}{z_first + 1}:{col}{s.r})", b.nf("naira", bold=True),
                       float(zy[years[j]].sum()) if j < 4 else float(zy.values.sum()))
zy_last = s.r
s.r += 2
s.band("2.10.13", "The default summarisation trap: Sum vs Count of Sales")
s.scenario("If even one Sales cell is text (like the 234 '₦1,234,560' values in the raw file), Excel's pivot defaults to COUNT of Sales. "
           "The numbers look plausible - they are just line counts. Always read the field caption.")
cnt = s25.groupby("Zone").Sales.agg(["sum", "count"]).reset_index()
cnt.columns = ["Zone", "Sum_of_Sales", "Count_of_Sales"]
s.table(cnt, "tSumVsCount", fmts={"Sum_of_Sales": "naira", "Count_of_Sales": "int"})
s.band("2.10.14", "Show Values As: % of row total and difference from previous year")
s.ws.write(s.r, 1, "% of Zone total", F["hdr"])
for j, y in enumerate(years):
    s.ws.write_number(s.r, 2 + j, y, F["hdr"])
s.r += 1
for i, z in enumerate(zy.index):
    s.ws.write(s.r, 1, z, F["lab"])
    for j, y in enumerate(years):
        src = f"{chr(67 + j)}{z_first + 1 + i}"
        s.ws.write_formula(s.r, 2 + j, f"={src}/$G${z_first + 1 + i}", b.nf("pct"), float(zy.loc[z, y] / zy.loc[z].sum()))
    s.r += 1
s.r += 1
s.ws.write(s.r, 1, "Difference from previous year", F["hdr"])
for j, y in enumerate(years[1:]):
    s.ws.write_number(s.r, 3 + j, y, F["hdr"])
s.r += 1
for i, z in enumerate(zy.index):
    s.ws.write(s.r, 1, z, F["lab"])
    for j, y in enumerate(years[1:], start=1):
        cur, prev = f"{chr(67 + j)}{z_first + 1 + i}", f"{chr(66 + j)}{z_first + 1 + i}"
        s.ws.write_formula(s.r, 2 + j, f"={cur}-{prev}", b.nf("naira"), float(zy.loc[z, y] - zy.loc[z, years[j - 1]]))
    s.r += 1
s.r += 1
s.band("2.10.11-12", "Grouping dates and numbers: quarterly sales 2025 and order-size bands")
q = s25.groupby("Quarter").Sales.sum()
s.calc([(f"2025 {qq} sales", f'=SUMIFS(tSM[Sales],tSM[Year],2025,tSM[Quarter],"{qq}")', float(q[qq]), "naira", "Pivot: right-click a date > Group > Quarters + Years")
        for qq in ["Q1", "Q2", "Q3", "Q4"]])
bands = [(0, 50000, "Under ₦50k"), (50000, 200000, "₦50k-200k"), (200000, 1000000, "₦200k-1m"), (1000000, 1e12, "₦1m and above")]
s.calc([(f"Lines {lbl}", f'=COUNTIFS(tSales[Sales],">={lo}",tSales[Sales],"<{hi:.0f}")', int(((s25.Sales >= lo) & (s25.Sales < hi)).sum()), "int",
         "Pivot: put Sales in Rows, right-click > Group > By 50000 ...") for lo, hi, lbl in bands])
s.band("2.10.18", "Calculated field: Margin = Profit / Sales")
s.scenario("In a pivot, PivotTable Analyze > Fields, Items & Sets > Calculated Field: Margin = Profit / Sales. Excel divides the SUMS (correct), "
           "not the average of each line's margin (wrong). The twin below proves the difference.")
s.calc([("Correct: SUM(Profit)/SUM(Sales) for FY2025", "=SUM(tSales[Profit])/SUM(tSales[Sales])", m_chain, "pct", "What a calculated field returns"),
        ("Wrong: AVERAGE of line margins", "=SUMPRODUCT(tSales[Profit]/tSales[Sales])/ROWS(tSales)", float((s25.Profit / s25.Sales).mean()), "pct",
         "Small loss-making lines drag this down - never average ratios")])
s.band("2.10.19", "Slicers and timelines")
s.bullets(["Click inside a live pivot > PivotTable Analyze > Insert Slicer > Zone, Category.",
           "Insert Timeline > Order_Date: drag across Q4 2025 to see the peak quarter.",
           "Report Connections: connect one Zone slicer to both live pivots so they filter together (3.8.10)."])

# live pivot sheet
lp = b.sheet("2.10 Live Pivots", "2.10  Live PivotTables on tSM (built by Excel when the file opens)",
             "If a pivot looks empty, press Data > Refresh All. Each pivot matches a formula twin on sheet 2.10 PivotTables.", "2.10.4 - 2.10.19")
lp.band("Pivot 1", "Sum of Sales by Zone (rows) and Year (columns)  -  twin: 2.10.4-9")
# the data sheet TRef is created later; store pivot specs and add after data sheets exist
PIVOTS = [("Pivot 1", dict(rows=["Zone"], columns=["Year"], values=[{"field": "Sales", "function": "sum", "num_format": "₦#,##0"}]), 9),
          ("Pivot 2", dict(rows=["Store_Name"], values=[{"field": "Sales", "function": "sum", "num_format": "₦#,##0"},
                                                       {"field": "Profit", "function": "sum", "num_format": "₦#,##0"}],
                           filters=["Year"]), 18),
          ("Pivot 3", dict(rows=["Category", "Sub_Category"], values=[{"field": "Profit", "function": "sum", "num_format": "₦#,##0"},
                                                                      {"field": "Order_ID", "function": "count", "num_format": "#,##0", "name": "Lines"}]), 26),
          ("Pivot 4", dict(rows=["Year", "Quarter"], values=[{"field": "Sales", "function": "sum", "num_format": "₦#,##0"}]), 26)]

# ======================================================================================
# 2.11 PivotCharts
# ======================================================================================
s2 = b.sheet("2.11 PivotCharts", "2.11  Introduction to PivotCharts",
             "Chart types for pivot output, built on the formula twins (insert a real PivotChart from any live pivot)", "2.11.1 - 2.11.11")
s2.band("2.11.4", "Creating a PivotChart")
s2.bullets(["Click inside Pivot 1 on '2.10 Live Pivots' > PivotTable Analyze > PivotChart > Clustered Column.",
            "The chart and pivot are linked both ways (2.11.7): filter Year on the chart's field button and the pivot filters too.",
            "Below: the same charts drawn from the formula twins so they display in any viewer."])
s2.band("2.11.9", "Clustered column: Sales by zone and year")
ch = b.wb.add_chart({"type": "column"})
pal = ["#0B2545", "#13807A", "#E0A526", "#B23A48"]
for j, y in enumerate(years):
    ch.add_series({"name": ["2.10 PivotTables", hdr_r, 2 + j], "categories": ["2.10 PivotTables", z_first, 1, z_first + 3, 1],
                   "values": ["2.10 PivotTables", z_first, 2 + j, z_first + 3, 2 + j], "fill": {"color": pal[j]}})
ch.set_title({"name": "SuperMart sales by zone, 2022-2025", "name_font": {"size": 11, "color": "#0B2545"}})
ch.set_y_axis({"num_format": '₦#,##0,,"m"', "major_gridlines": {"visible": True, "line": {"color": "#E5E5E5"}}})
ch.set_legend({"position": "bottom"}); ch.set_size({"width": 720, "height": 320})
s2.ws.insert_chart(s2.r, 1, ch); s2.r += 17
s2.band("2.11.11", "Practical PivotChart examples across the three datasets")
dep = lc.groupby("Department").agg(Admissions=("Admission_ID", "count"), ALOS=("Length_of_Stay", "mean")).reset_index()
td = s2.table(dep, "tDeptPC", fmts={"Admissions": "int", "ALOS": "dec1"}, widths={"Department": 26})
s2.chart("bar", [{"name": "Admissions", "categories": td.ref("Department"), "values": td.ref("Admissions"), "labels": True, "label_fmt": "#,##0"}],
         "Lifecare admissions by department, 2019-2024", legend=False, rows=16)
cv = L("cv_perf")
gr = cv.Grade.value_counts().reindex(["A", "B", "C", "D", "F"]).reset_index()
gr.columns = ["Grade", "Students"]
tg = s2.table(gr, "tGradePC", fmts={"Students": "int"})
s2.chart("pie", [{"name": "Students", "categories": tg.ref("Grade"), "values": tg.ref("Students"), "labels": True,
                  "points": [{"fill": {"color": c}} for c in ["#13807A", "#7FB069", "#E0A526", "#D98E04", "#B23A48"]]}],
         "Crestview grade distribution (5 slices - pie is acceptable)", rows=16, size=(480, 300))
s2.warn("Pie charts only work with a handful of parts that sum to a whole. For 12 stores use a sorted bar chart instead (3.7.14).")

# ---------------------------------------------------------------------- data sheets (end of workbook)
sm_cols = sales_cols + ["Year", "Quarter", "Month"]
tSM = b.data_sheet("D_SuperMart_All", sm[sm_cols], "tSM", fmts={"Order_Date": "date", "Ship_Date": "date", "Sales": "naira", "Profit": "naira",
                                                                "Discount": "pct0", "Quantity": "int", "Row_ID": "gen", "Year": "gen"},
                   title="SuperMart Stores - all sales lines 2022-2025 (clean)", note="Source: 00_Core_Datasets/SuperMart_Stores_Sales_2022_2025.csv + Year/Quarter/Month helper columns")
tS = b.data_sheet("D_Sales_2025", s25[sales_cols], "tSales", fmts={"Order_Date": "date", "Ship_Date": "date", "Sales": "naira", "Profit": "naira",
                                                                   "Discount": "pct0", "Quantity": "int", "Row_ID": "gen"},
                  title="SuperMart Stores - FY2025 sales lines (clean)", note="3,312 lines - the cleaned version of SuperMart_Sales_2025_DIRTY.csv")
b.data_sheet("D_Stores", stores_x, "tStores", fmts={"Opened": "date", "Target_2025": "naira", "Sales_2025": "naira", "Profit_2025": "naira", "Floor_Area_sqm": "int"},
             title="SuperMart store directory with FY2025 targets and actuals", note="Source: SuperMart_Stores_Directory.csv")
b.data_sheet("D_Dirty_2025", dd, "tDirty", fmts={"Order_Date": "date", "Ship_Date": "date", "Row_ID": "gen"},
             title="RAW - SuperMart_Sales_2025_DIRTY.csv as Excel opens it (do not clean here)", note="3,387 rows incl. 15 blank rows and 60 duplicates")
for tag, spec, h_ in PIVOTS:
    if tag != "Pivot 1":
        lp.band(tag, {"Pivot 2": "Sales and Profit by store - Report Filter: Year (choose 2025)",
                      "Pivot 3": "Profit and line count by Category > Sub-Category (all years)",
                      "Pivot 4": "Sales by Year > Quarter (seasonality - Q4 peaks every year)"}[tag])
    lp.pivot(tSM, height=h_, **spec)

b.cover("Week 2 Section Examples - Power Query, Cleaning, Functions and PivotTables",
        ["One worksheet per module (2.1 - 2.11); the 65 functions are split into Part A and Part B like the slides.",
         "Data sheets at the end (D_...) are official Excel Tables: tSM (all SuperMart lines), tSales (FY2025 clean), tStores, tDirty (raw FY2025 extract).",
         "Practice files for import (CSV, TXT, EU CSV, folder, JSON, HTML, PDF, XLSX) are in 01_Section_Examples/Week2_Data_Files. "
         "Set your own folder path once on sheet 2.3 (named cell FolderPath).",
         "Live PivotTables are built by Excel when the workbook opens; every pivot has a formula twin so results can be checked."],
        [(k, DATASETS_TEXT[k]) for k in ["SuperMart Stores", "SuperMart Online", "Lifecare", "Crestview"]])
b.close()
print("Week 2 ->", OUT)
