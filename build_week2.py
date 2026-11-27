#!/usr/bin/env python3
"""Week 2 Section Examples — Power Query, Cleaning, 65 Functions, M, Pivots (Modules 2.1–2.11)."""
import pandas as pd, numpy as np
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from xl_helpers import *

ROOT=Path("/home/user/ETEH")
CLEAN=ROOT/"ETEH_Bootcamp/Datasets/CLEAN"
DIRTY=ROOT/"ETEH_Bootcamp/Datasets/DIRTY_FOR_WEEK2"
OUT=ROOT/"01_Section_Examples/Week2_PowerQuery_Pivots_Section_Examples.xlsx"

tr=pd.read_csv(CLEAN/"ecommerce_transactions_CLEAN.csv")
pr=pd.read_csv(CLEAN/"ecommerce_products_CLEAN.csv")
ss=pd.read_csv(CLEAN/"superstore_CLEAN.csv")
sal=pd.read_csv(CLEAN/"salary_CLEAN.csv")
sal_d=pd.read_csv(DIRTY/"salary_DIRTY.csv")
spf=pd.read_csv(CLEAN/"student_performance_CLEAN.csv")

wb=Workbook()
readme_sheet(wb,
 "Week 2 — Power Query, Cleaning, Functions, M & Pivots: Section Examples (Modules 2.1–2.11)",
 "ETEH Ltd · Eaglet2EagleHub Limited · Dirty files are intentional — clean them to match the CLEAN truth. All formulas live.",
 ["Week 2 turns messy extracts into repeatable pipelines. Modules 2.1–2.3 teach importing attached CSV/XLSX/TXT files safely. Modules 2.4–2.6 clean the DIRTY files (salary, superstore, e-commerce, education, healthcare) to reconcile to CLEAN. Modules 2.7–2.8 master 65 analytical functions on real data. Module 2.9 reads/writes M. Modules 2.10–2.11 build PivotTables/Charts.",
  "Power Query steps cannot be executed by openpyxl — this workbook gives you the exact clicks, the M code to paste, and the expected output to verify against. Excel formulas (green) do execute. PivotTables must be built by hand (Insert > PivotTable) — expected pivot outputs are provided as values to check against.",
  "Numbering note: Vol 2 was renumbered 2.1–2.11 in the CLEANED deck (old duplicate 2.4/2.7 fixed; External Web/PDF + labs modules removed). Old→new map is on the ANSWERS sheet."],
 [["salary_DIRTY.csv (1,580 rows) → salary_CLEAN.csv (1,791)","12_DATA_SalaryDirty + 13_DATA_SalaryClean","2.4, 2.5, 2.6","500","One person-job record"],
  ["superstore_DIRTY.csv (1,201) → superstore_CLEAN.csv","14_DATA_SuperDirty","2.4, 2.5","500","One order line"],
  ["ecommerce_transactions_DIRTY.csv (2,100) → CLEAN","15_DATA_EcomDirty","2.4, 2.5, 2.10","500","One transaction line"],
  ["student_performance_DIRTY.csv → CLEAN","16_DATA_EduDirty","2.4, 2.5","500","One student"],
  ["healthcare_DIRTY.csv → ANONYMISED","17_DATA_HealthDirty","2.4, 2.5","500","One admission"]],
 ["1) Work in order 2.1→2.11.  2) For cleaning: copy DIRTY to your folder, clean in Excel/Power Query, then reconcile row counts + SUMs to CLEAN (use DIRTY_ISSUES_LOG.csv as checklist).  3) For functions: type the formula in the yellow column, compare to green Expected.  4) For pivots: build by hand, compare to Expected pivot values."])

# ---------- 2.1 ----------
ws=wb.create_sheet("01_2.1_Sources")
setup_sheet(ws,"Module 2.1 — Understanding Data Sources & File Formats",".xlsx/.csv/.txt imports, delimiters, encoding. All files are attached datasets — no web/PDF sources.")
r=4
r=section(ws,r,"Objectives",ncols=6)
r=note(ws,r,"Choose the right import route for each attached file and avoid locale/encoding traps (UTF-8 vs Windows-1252, M/D/YYYY vs ISO).",ncols=6,height=24)
r=section(ws,r,"File-format chooser — attached datasets",ncols=6)
r=write_table(ws,r,["Attached file","Format","Import route (Data > Get Data)","Trap to check"],[
 ["ecommerce_transactions_CLEAN.csv","CSV (UTF-8)","From Text/CSV → ISO dates, ₦ as number","delivery_id padding; 570 Cancelled have no delivery"],
 ["superstore_CLEAN.csv","CSV (UTF-8)","From Text/CSV → dates as Date","Original was Latin-1 + M/D/YYYY; CLEAN is ISO"],
 ["salary_CLEAN.csv","CSV (UTF-8, BOM fixed)","From Text/CSV → Salary as Number","Original had BOM + ₦ text in DIRTY"],
 ["ecommerce_EXTRACT.xlsx","XLSX (multi-sheet)","From Workbook → pick sheet","One sheet = one tidy table; check header row"],
 ["hotel_reservations_CLEAN.csv","CSV","From Text/CSV → build Arrival_Date_ISO","Y/M/D split columns; 37 invalid Feb-29 flagged"],
],widths=[36,18,42,42])
r=section(ws,r,"Delimiter drill — TXT legacy export (tab-delimited)",ncols=6)
r=note(ws,r,"A legacy TXT export uses TAB (\\t) with \" text qualifier. In Power Query: From Text/CSV → Delimiter: Tab → Encoding: UTF-8 → Locale: English (Nigeria). If commas appear inside product names ('Chair, Rounded Back'), CSV with \" qualifier handles them; TXT with comma delimiter would split wrongly. Always preview: check that Row count matches the file's line count minus header.",ncols=6,height=52)
r=section(ws,r,"Encoding drill — Superstore original (Latin-1)",ncols=6)
r=note(ws,r,"Original Sample - Superstore.csv is Latin-1 (byte 0xF6 = ö in 'Björk' etc.). Opening as UTF-8 shows �. Fix: Power Query Encoding → Windows-1252 for the RAW file, then save CLEAN as UTF-8. CLEAN file is already UTF-8 — verify by opening in Notepad: no � characters.",ncols=6,height=48)

# ---------- 2.2 ----------
ws=wb.create_sheet("02_2.2_PowerQuery")
setup_sheet(ws,"Module 2.2 — Introduction to Power Query","ETL, editor anatomy, applied steps, preview tools. Worked on salary_DIRTY.")
r=4
r=section(ws,r,"Objectives",ncols=6)
r=note(ws,r,"Explain ETL, navigate the editor, and use Column Quality/Distribution/Profile to find problems before cleaning.",ncols=6,height=24)
r=section(ws,r,"ETL on salary_DIRTY (1,580 rows) — exact clicks",ncols=6)
r=write_table(ws,r,["ETL stage","Clicks (Power Query Editor)","What you should see"],[
 ["Extract","Data > From Text/CSV → salary_DIRTY.csv → Encoding UTF-8, Delimiter Comma","1,580 rows; Salary shows ₦ text; Education shows BACHELORS/masters"],
 ["Transform","Applied Steps: Changed Type → Replace Values → Trim → Remove Duplicates","Each step appears in Applied Steps (right pane); row count drops as you clean"],
 ["Load","Home > Close & Load → Only Create Connection (for pivots) or Table (for analysis)","Query appears in Queries & Connections; Refresh All re-runs steps"],
],widths=[18,62,58])
r=section(ws,r,"Preview tools — what to check first",ncols=6)
r=write_table(ws,r,["Tool (Home tab)","What it tells you","Salary_DIRTY finding"],[
 ["Column Quality","% valid / error / null per column","Salary: ~8% error (text); Age: ~4% null"],
 ["Column Distribution","Distinct values + counts","Education: BACHELORS vs Bachelor's Degree split"],
 ["Column Profile","Types, lengths, min/max","Salary min -5000 (impossible); Age max 150"],
],widths=[28,52,58])
r=note(ws,r,"Watch out: Power Query types are strict. If Salary is Text, SUM fails. Change Type to Number AFTER stripping ₦/commas (Replace Values), not before. Keep a RAW query (Load = No) and a CLEAN query that references it — never overwrite RAW.",ncols=6,height=42)

# ---------- 2.3 ----------
ws=wb.create_sheet("03_2.3_Connecting")
setup_sheet(ws,"Module 2.3 — Connecting & Importing with Power Query","Folder combine, parameters, refresh. E-commerce monthly files example.")
r=4
r=section(ws,r,"Objectives",ncols=6)
r=note(ws,r,"Combine monthly CSVs from a folder, parameterise paths, and refresh after new files arrive.",ncols=6,height=24)
r=section(ws,r,"Folder combine — E-commerce monthly exports",ncols=6)
r=note(ws,r,"Scenario: Finance drops ecommerce_2025-01.csv, ecommerce_2025-02.csv… into Datasets/CLEAN/. Clicks: Data > Get Data > From Folder → select folder → Combine > Combine Files → check 'Use first file for preview'. Power Query infers columns from the first file; if a later file adds a column, the combine fails — fix by promoting headers consistently. Expected: combined rows = sum of monthly rows; order_month derived from filename or order_date.",ncols=6,height=58)
r=section(ws,r,"Parameters & refresh",ncols=6)
r=write_table(ws,r,["Parameter","How to set","Why"],[
 ["FolderPath (Text)","New Parameter → FolderPath → value = C:\\ETEH\\Datasets\\CLEAN","Change once, refresh everywhere; no hard-coded paths in M"],
 ["Refresh date","Add step: Table.AddColumn(..., DateTime.LocalNow())","Audit trail: proves when data was pulled"],
 ["Refresh All","Data > Refresh All (Ctrl+Alt+F5)","Re-runs all queries; check row counts after refresh"],
],widths=[28,62,48])
r=note(ws,r,"Watch out: absolute paths break on another machine. Use a relative folder + Parameter, and document the expected folder layout in 00_README. Test refresh by adding a test CSV and confirming row count increases.",ncols=6,height=36)

# ---------- 2.4 ----------
ws=wb.create_sheet("04_2.4_CleanExcel")
setup_sheet(ws,"Module 2.4 — Data Cleaning with Excel","The 12 quality problems on DIRTY files. Manual fixes + verification to CLEAN truth.")
r=4
r=section(ws,r,"Objectives",ncols=6)
r=note(ws,r,"Find and fix all 12 problems in Excel (Find/Replace, TRIM/CLEAN, Text to Columns, validation) and prove the result reconciles to CLEAN.",ncols=6,height=24)
r=section(ws,r,"The 12 problems — where they hide in DIRTY files",ncols=6)
r=write_table(ws,r,["#","Problem","DIRTY example","Excel fix"],[
 ["1","Duplicates","80 exact salary copies","Data > Remove Duplicates (check row count)"],
 ["2","Inconsistent categories","BACHELORS / masters / phd","Lookup table + XLOOKUP/VLOOKUP; PROPER"],
 ["3","Numbers as text","₦120,000 / $1,234.56","SUBSTITUTE ₦,$,comma → VALUE; Number format"],
 ["4","Extra spaces","'  Software Engineer '","TRIM; Find '  ' → ' '"],
 ["5","Case mess","MALE / female / High","UPPER/LOWER/PROPER standardise"],
 ["6","Mixed dates","2025-01-05 vs 01/05/2025 vs 5-Jan-25","DATEVALUE with locale; reformat ISO"],
 ["7","Missing (blank)","Salary/Age blanks","Decide blank vs 0 vs Unknown; document"],
 ["8","Impossible values","Salary -5000, Age 150","IF flag; quarantine; do not delete silently"],
 ["9","Spacer rows","Blank row in superstore_DIRTY","Go To Special > Blanks → delete"],
 ["10","Text in numeric ('23 hrs')","Hours_Studied '23 hrs'","SUBSTITUTE ' hrs' → VALUE"],
 ["11","Key padding broken","DEL1 vs DEL000001","'DEL'&TEXT(VALUE(...),'000000')"],
 ["12","Status case + spaces","'DELIVERED' / ' Paid '","TRIM + UPPER list; validation"],
],widths=[8,28,42,60])
r=section(ws,r,"Verification — reconcile to CLEAN (type your counts)",ncols=6)
r=write_table(ws,r,["Check","DIRTY (before)","Your CLEAN (after)","CLEAN truth (Python)"],[
 ["Salary rows","1,580","", "1,791 (full CLEAN; DIRTY is 1,500-sample + 80 dups)"],
 ["Salary SUM","(text — SUM fails)","", f"₦{sal['Salary'].sum():,.0f} (CLEAN)"],
 ["Superstore Sales SUM","(text — SUM fails)","", f"${ss['Sales'].sum():,.0f}"],
 ["E-commerce Total SUM","(text — SUM fails)","", f"₦{tr['total_price'].sum():,.0f}"],
],widths=[28,28,28,54])
r=note(ws,r,"Watch out: Remove Duplicates is irreversible — copy first. Text-to-Columns splits on delimiter; preview and keep original. After cleaning, re-run: COUNT vs COUNTA, SUM, and a spot-check of 5 rows against CLEAN.",ncols=6,height=42)

# ---------- 2.5 ----------
ws=wb.create_sheet("05_2.5_CleanPQ")
setup_sheet(ws,"Module 2.5 — Cleaning & Transforming in Power Query","Remove/filter/replace/split/merge/group/pivot/unpivot. M code included.")
r=4
r=section(ws,r,"Objectives",ncols=6)
r=note(ws,r,"Reproduce every Module 2.4 fix as Applied Steps so Refresh All re-cleans automatically.",ncols=6,height=24)
r=section(ws,r,"Step-by-step — salary_DIRTY → CLEAN (Power Query)",ncols=6)
r=write_table(ws,r,["Step","Power Query action","M (Advanced Editor) snippet"],[
 ["1 Remove blanks","Select rows → Remove Rows → Remove Blank Rows","Table.RemoveRowsWithErrors / FilterRows"],
 ["2 Dedupe","Remove Duplicates (on all columns)","Table.Distinct(Source)"],
 ["3 Replace","Replace Values: BACHELORS→Bachelor's Degree (3 replacements)","Table.ReplaceValue(Source,\"BACHELORS\",\"Bachelor's Degree\",Replacer.ReplaceText,{\"Education Level\"})"],
 ["4 Split","Split Column by Delimiter (if needed)","Table.SplitColumn(Source,\"FullName\",\" \")"],
 ["5 Custom col","Add Column → Custom Column: = [Salary] * 1.1 (bonus)","Table.AddColumn(Source,\"Salary_Bonus\",each [Salary]*1.1)"],
 ["6 Group","Group By Job_Family → Sum Salary","Table.Group(Source,{\"Job_Family\"},{\"SumSalary\",each List.Sum([Salary]),type number})"],
 ["7 Unpivot","Unpivot Other Columns (wide→tidy)","Table.UnpivotOtherColumns(Source,{},\"Attribute\",\"Value\")"],
],widths=[18,52,68])
r=note(ws,r,"Applied Steps discipline: rename each step ('01_RemoveBlanks', '02_Dedupe'…), disable steps to test (right-click > Disable), and keep RAW + CLEAN queries separate. If Refresh changes the total unexpectedly, disable steps one by one to find the culprit — that is the audit trail.",ncols=6,height=48)

# ---------- 2.6 ----------
ws=wb.create_sheet("06_2.6_Compare")
setup_sheet(ws,"Module 2.6 — Excel Cleaning vs Power Query Cleaning","Reproducibility, auditability, tool choice. Same dataset, two routes.")
r=4
r=section(ws,r,"Objectives",ncols=6)
r=note(ws,r,"Choose the right tool and prove your number with a second independent check.",ncols=6,height=24)
r=write_table(ws,r,["Criterion","Manual Excel","Power Query","Verdict for DIRTY files"],[
 ["Reproducibility","Steps in head; Refresh does nothing","Steps recorded; Refresh All re-cleans","Power Query for monthly/large files"],
 ["Auditability","No log; hard to prove","Applied Steps = log; M code exportable","Power Query for assessed work"],
 ["Speed","Fast for <500 rows, one-off","Slower to set up, faster to re-run","Excel for ad-hoc; PQ for repeat"],
 ["Error handling","Silent (SUM ignores text)","Errors flagged; exception table possible","PQ + exception table for finance"],
],widths=[24,42,42,30])
r=note(ws,r,"Practical comparison (do both): Clean salary_DIRTY Education labels manually (Find/Replace) and via Power Query (Replace Values). Compare: manual takes ~2 min, PQ takes ~5 min first time but 0 min on refresh. Record both times — that is your evidence for the choice.",ncols=6,height=42)

# ---------- 2.7 (functions part 1) ----------
ws=wb.create_sheet("07_2.7_Functions1")
setup_sheet(ws,"Module 2.7 — The 65 Functions (Part 1: Aggregation & Logic)","SUM…TEXTJOIN (33 functions). Practice data: 30 E-commerce lines. Green = model formula; Yellow = your attempt.")
r=4
r=section(ws,r,"Practice data (30 lines, same as Week 1 Module 1.7 — values below)",ncols=6)
p30 = tr.head(30)[["quantity","unit_price","total_price","commission_amount","order_status","payment_method"]]
r=write_table(ws,r,["Qty","Unit (₦)","Total (₦)","Commission (₦)","Status","Payment"],p30.values.tolist(),widths=[10,18,18,18,14,20],money_cols=(2,3,4))
def _find_hdr(sheet,label,col=1):
    for _row in sheet.iter_rows(min_row=1,max_row=sheet.max_row,min_col=col,max_col=col):
        if _row[0].value==label: return _row[0].row
    raise ValueError(label)
hdr27 = _find_hdr(ws,"Qty")
r=section(ws,r,"Function lab — type your formula in column D (yellow), compare to Expected (green model in E)",ncols=6)
funcs1 = [
 ("SUM","=SUM(C{0}:C{1})","Total revenue"),
 ("SUMIF","=SUMIF(E{0}:E{1},\"Delivered\",C{0}:C{1})","Delivered revenue"),
 ("SUMIFS","=SUMIFS(C{0}:C{1},E{0}:E{1},\"Delivered\",F{0}:F{1},\"Paystack\")","Delivered via Paystack"),
 ("AVERAGE","=AVERAGE(C{0}:C{1})","Average line"),
 ("AVERAGEIFS","=AVERAGEIFS(C{0}:C{1},E{0}:E{1},\"Delivered\")","Average Delivered line"),
 ("COUNT","=COUNT(C{0}:C{1})","Count numbers in Total"),
 ("COUNTA","=COUNTA(E{0}:E{1})","Count non-blank Status"),
 ("COUNTBLANK","=COUNTBLANK(E{0}:E{1})","Blanks in Status (expect 0)"),
 ("COUNTIF","=COUNTIF(E{0}:E{1},\"Delivered\")","Delivered lines"),
 ("COUNTIFS","=COUNTIFS(E{0}:E{1},\"Delivered\",F{0}:F{1},\"Flutterwave\")","Delivered + Flutterwave"),
 ("SUBTOTAL","=SUBTOTAL(9,C{0}:C{1})","Visible-rows sum (9=SUM)"),
 ("AGGREGATE","=AGGREGATE(9,6,C{0}:C{1})","Sum ignoring errors"),
 ("MEDIAN","=MEDIAN(C{0}:C{1})","Median line"),
 ("MODE.SNGL","=MODE.SNGL(A{0}:A{1})","Most common Qty"),
 ("MAX","=MAX(C{0}:C{1})","Largest line"),
 ("MIN","=MIN(C{0}:C{1})","Smallest line"),
 ("LARGE","=LARGE(C{0}:C{1},2)","2nd largest line"),
 ("SMALL","=SMALL(C{0}:C{1},2)","2nd smallest line"),
 ("ROUND","=ROUND(AVERAGE(C{0}:C{1}),2)","Average to 2dp"),
 ("ROUNDUP","=ROUNDUP(AVERAGE(C{0}:C{1}),0)","Average rounded up"),
 ("ROUNDDOWN","=ROUNDDOWN(AVERAGE(C{0}:C{1}),0)","Average rounded down"),
 ("ABS","=ABS(MIN(C{0}:C{1}))","Absolute of min"),
 ("IF","=IF(AVERAGE(C{0}:C{1})>100000,\"High\",\"Low\")","High/Low average"),
 ("IFS","=IFS(AVERAGE(C{0}:C{1})>150000,\"Very high\",AVERAGE(C{0}:C{1})>100000,\"High\",TRUE,\"Low\")","Band"),
 ("AND","=AND(AVERAGE(C{0}:C{1})>50000,COUNTIF(E{0}:E{1},\"Delivered\")>10)","Both true?"),
 ("OR","=OR(AVERAGE(C{0}:C{1})>150000,COUNTIF(E{0}:E{1},\"Delivered\")>25)","Either true?"),
 ("NOT","=NOT(AVERAGE(C{0}:C{1})>150000)","Not very high?"),
 ("IFERROR","=IFERROR(AVERAGE(C{0}:C{1})/0,\"No data\")","Error demo (divide by 0)"),
 ("IFNA","=IFNA(VLOOKUP(\"X\",A{0}:F{1},6,FALSE),\"Not found\")","NA demo"),
 ("SWITCH","=SWITCH(COUNTIF(E{0}:E{1},\"Delivered\"),\"Many\",\"Few\")","Switch demo (simplified)"),
 ("TEXT","=TEXT(AVERAGE(C{0}:C{1}),\"#,##0\")","Average as text"),
 ("CONCAT","=CONCAT(\"Lines: \",COUNT(E{0}:E{1}))","Label + count"),
 ("TEXTJOIN","=TEXTJOIN(\", \",TRUE,\"Delivered\",\"Completed\")","Join list"),
]
r=header_row(ws,r,["#","Function","What it does","Your formula (yellow)","Model formula (green)"],widths=[8,16,36,44,44])
for i,(fn,model,desc) in enumerate(funcs1,1):
    ws.cell(row=r,column=1,value=i).font=Font(name="Calibri",size=9)
    ws.cell(row=r,column=2,value=fn).font=Font(name="Calibri",size=9,bold=True)
    ws.cell(row=r,column=3,value=desc).font=Font(name="Calibri",size=9)
    ws.cell(row=r,column=4,value="").fill=PatternFill("solid",fgColor=GOLD)
    ws.cell(row=r,column=5,value=model.format(hdr27+1, hdr27+30)).fill=PatternFill("solid",fgColor=GREEN)
    ws.cell(row=r,column=5).font=Font(name="Calibri",size=9)
    for c in range(1,6):
        ws.cell(row=r,column=c).border=BORDER
    ws.row_dimensions[r].height=16
    r+=1
r=note(ws,r,"Copy the model formula down column E to verify it calculates. Then retype it in D without looking. SUBTOTAL(9) respects filters; SUM does not — test by filtering Status=Delivered and comparing. AGGREGATE(9,6) ignores errors — useful when DIRTY numbers contain text.",ncols=6,height=42)

# ---------- 2.8 (functions part 2) ----------
ws=wb.create_sheet("08_2.8_Functions2")
setup_sheet(ws,"Module 2.8 — Text, Date, Lookup & Dynamic Arrays (32 functions)","LEFT…LET. Practice data: E-commerce IDs + Superstore dates.")
r=4
r=section(ws,r,"Practice data (12 rows)",ncols=6)
prac = tr.head(12)[["transaction_id","order_id","customer_id","order_date","total_price"]]
# add a product name for lookup demo
prac = prac.merge(pr[["product_id","product_name"]], left_index=True, right_index=True, how="left") if False else prac
# simpler: take product names separately
pnames = pr.head(12)["product_name"].tolist()
prac["product_name"]=pnames
r=write_table(ws,r,["Transaction","Order ID","Customer","Order date","Total (₦)","Product"],prac.values.tolist(),widths=[16,22,14,14,18,30],money_cols=(5,))
hdr28 = _find_hdr(ws,"Transaction")
r=section(ws,r,"Function lab — model formulas in green (copy to verify)",ncols=6)
funcs2 = [
 ("LEFT","=LEFT(B{0},3)","ORD"),
 ("RIGHT","=RIGHT(B{0},4)","Last 4 of Order ID"),
 ("MID","=MID(B{0},4,8)","Date part of Order ID (YYYYMMDD)"),
 ("LEN","=LEN(B{0})","Length of Order ID"),
 ("FIND","=FIND(\"2024\",B{0})","Position of 2024 (error if absent)"),
 ("SEARCH","=SEARCH(\"ord\",B{0})","Case-insensitive find"),
 ("SUBSTITUTE","=SUBSTITUTE(C{0},\"CUST\",\"Customer \")","Rename prefix"),
 ("REPLACE","=REPLACE(C{0},1,4,\"Client \")","Replace first 4 chars"),
 ("TRIM","=TRIM(\"  \"&C{0}&\"  \")","Remove spaces demo"),
 ("CLEAN","=CLEAN(CHAR(9)&C{0})","Remove non-printing demo"),
 ("PROPER","=PROPER(\"e-commerce nigeria\")","Title case"),
 ("UPPER","=UPPER(C{0})","Uppercase ID"),
 ("LOWER","=LOWER(C{0})","Lowercase ID"),
 ("VALUE","=VALUE(\"1,200\")","Text→number (1200)"),
 ("TEXTBEFORE","=TEXTBEFORE(B{0},\"-\")","Before first dash"),
 ("TEXTAFTER","=TEXTAFTER(B{0},\"-\")","After first dash"),
 ("TODAY","=TODAY()","Today (volatile)"),
 ("NOW","=NOW()","Now (volatile)"),
 ("DATE","=DATE(2025,1,5)","Date from Y/M/D"),
 ("DATEDIF","=DATEDIF(D{0},TODAY(),\"D\")","Days since order (approx)"),
 ("EOMONTH","=EOMONTH(D{0},0)","Month-end of order"),
 ("NETWORKDAYS","=NETWORKDAYS(D{0},EOMONTH(D{0},0))","Workdays in order month"),
 ("WORKDAY","=WORKDAY(D{0},5)","5 workdays after order"),
 ("VLOOKUP","=VLOOKUP(C{0},'10_DATA_Ecommerce'!C:C,1,FALSE)","Lookup demo (see note)"),
 ("XLOOKUP","=XLOOKUP(C{0},'10_DATA_Ecommerce'!C:C,'10_DATA_Ecommerce'!C:C,\"Not found\")","XLOOKUP demo"),
 ("INDEX","=INDEX(E{0}:E{1},1)","First Total"),
 ("MATCH","=MATCH(MAX(E{0}:E{1}),E{0}:E{1},0)","Row of largest Total"),
 ("FILTER","=FILTER(A{0}:E{1},E{0}:E{1}>100000,\"None\")","Dynamic spill: large lines"),
 ("UNIQUE","=UNIQUE(E{0}:E{1})","Unique Totals (spill)"),
 ("SORT","=SORT(E{0}:E{1},1,-1)","Sorted Totals desc (spill)"),
 ("SORTBY","=SORTBY(A{0}:A{1},E{0}:E{1},-1)","Transactions by Total desc"),
 ("LET","=LET(x,AVERAGE(E{0}:E{1}),IF(x>100000,\"High\",\"Low\"))","Named calc"),
]
r=header_row(ws,r,["#","Function","What it does","Model formula (green)"],widths=[8,16,44,70])
for i,(fn,model,desc) in enumerate(funcs2,1):
    ws.cell(row=r,column=1,value=i).font=Font(name="Calibri",size=9)
    ws.cell(row=r,column=2,value=fn).font=Font(name="Calibri",size=9,bold=True)
    ws.cell(row=r,column=3,value=desc).font=Font(name="Calibri",size=9)
    ws.cell(row=r,column=4,value=model.format(hdr28+1, hdr28+12)).fill=PatternFill("solid",fgColor=GREEN)
    ws.cell(row=r,column=4).font=Font(name="Calibri",size=9)
    for c in range(1,5):
        ws.cell(row=r,column=c).border=BORDER
    ws.row_dimensions[r].height=16
    r+=1
r=note(ws,r,"VLOOKUP/XLOOKUP note: they reference 10_DATA_Ecommerce (500 rows). Exact-match FALSE is essential — approximate match returns wrong rows. FILTER/UNIQUE/SORT spill — keep cells below empty. TODAY/NOW recalculate on open; do not use them as stable keys. TEXTBEFORE/TEXTAFTER need Microsoft 365; older Excel: use LEFT/FIND.",ncols=6,height=48)

# ---------- 2.9 M ----------
ws=wb.create_sheet("09_2.9_M")
setup_sheet(ws,"Module 2.9 — Introduction to the M Language","Syntax, let-expressions, core functions. Paste-ready M for salary cleaning.")
r=4
r=section(ws,r,"Objectives",ncols=6)
r=note(ws,r,"Read any M query, edit a step safely, and know when to use the GUI vs Advanced Editor.",ncols=6,height=24)
r=section(ws,r,"M anatomy — salary cleaning query (paste into Advanced Editor to inspect)",ncols=6)
m_code = """let
    Source = Csv.Document(File.Contents("C:\\ETEH\\Datasets\\DIRTY_FOR_WEEK2\\salary_DIRTY.csv"), [Delimiter=",", Encoding=65001]),
    PromotedHeaders = Table.PromoteHeaders(Source, [MissingField.Error]),
    ChangedType = Table.TransformColumnTypes(PromotedHeaders, {{"Salary", type number}}),
    Replaced = Table.ReplaceValue(ChangedType, "BACHELORS", "Bachelor's Degree", Replacer.ReplaceText, {"Education Level"}),
    Trimmed = Table.TransformColumns(Replaced, {{"Job Title", Text.Trim, type text}}),
    Deduped = Table.Distinct(Trimmed),
    Filtered = Table.SelectRows(Deduped, each [Salary] <> null)
in
    Filtered"""
ws.merge_cells(start_row=r,start_column=1,end_row=r+10,end_column=6)
c=ws.cell(row=r,column=1,value=m_code); c.font=Font(name="Consolas",size=9); c.alignment=Alignment(wrap_text=True,vertical="top")
ws.row_dimensions[r].height=150
r+=11
r=section(ws,r,"M function families (with Excel equivalents)",ncols=6)
r=write_table(ws,r,["Family","M example","Excel equivalent"],[
 ["Text","Text.Trim([Job Title])","TRIM"],
 ["Number","Number.Round([Salary],2)","ROUND"],
 ["Date","Date.From([order_date])","DATEVALUE"],
 ["Table","Table.Group(Source,{\"Job_Family\"},…)","PivotTable / GROUP BY"],
 ["List","List.Sum([Salary])","SUM"],
],widths=[18,62,58])
r=note(ws,r,"M is case-sensitive (Table.Distinct ≠ table.distinct). Each step is a let-binding; the final expression is the result. The Formula Bar shows the current step's M — edit there for small fixes, Advanced Editor for rewrites. Always keep a copy before editing.",ncols=6,height=42)

# ---------- 2.10 Pivots ----------
ws=wb.create_sheet("10_2.10_Pivots")
setup_sheet(ws,"Module 2.10 — Introduction to PivotTables","Four areas, grouping, Show Values As, slicers. Build by hand on E-commerce extract.")
r=4
r=section(ws,r,"Objectives",ncols=6)
r=note(ws,r,"Build a pivot that answers a question in under 2 minutes and spot the Sum-vs-Count trap.",ncols=6,height=24)
r=section(ws,r,"Build 1 — Revenue by Category (E-commerce)",ncols=6)
r=note(ws,r,"Clicks: Select 10_DATA_Ecommerce (or 15_DATA_EcomDirty CLEANED) → Insert > PivotTable → New Worksheet. Rows: Category (from Products join — see note), Values: Sum of total_price. Expected (full 11,000 lines): Fashion ₦243.1m, Health ₦224.6m, Automotive ₦219.1m, Electronics ₦212.5m, Home ₦177.3m, Food ₦170.5m, Beauty ₦155.0m. Note: Category lives in Products — for a pure Excel pivot, first VLOOKUP Category onto Transactions (Week 2 Module 2.8) or use the Data Model (Week 3).",ncols=6,height=62)
r=section(ws,r,"Build 2 — Monthly revenue trend (for forecasting)",ncols=6)
r=note(ws,r,"Rows: order_month (group order_date by Months), Values: Sum of total_price. Expected shape: 2020–2025 monthly totals (see Datasets/XLSX_EXTRACTS/ecommerce_EXTRACT.xlsx Monthly_Summary). Group dates: right-click a date in Rows > Group > Months/Quarters/Years. Show Values As: % of Total for share; Difference from Previous for growth.",ncols=6,height=52)
r=section(ws,r,"The Sum-vs-Count trap",ncols=6)
r=note(ws,r,"If Values shows 'Count of order_id' instead of 'Sum of total_price', Excel defaulted because the field is text. Fix: drag total_price to Values → Sum. Always check the Values header before quoting. Refresh: Data > Refresh All after source changes; if numbers don't move, check the pivot cache source range.",ncols=6,height=48)
r=section(ws,r,"Slicers & Timelines",ncols=6)
r=note(ws,r,"Insert > Slicer: order_status, payment_method. Insert > Timeline: order_date. Connect one slicer to multiple pivots (Slicer > Report Connections). Test: filter Status=Delivered — every pivot must update.",ncols=6,height=36)

# ---------- 2.11 PivotCharts ----------
ws=wb.create_sheet("11_2.11_PivotCharts")
setup_sheet(ws,"Module 2.11 — Introduction to PivotCharts","Visual twin of the pivot. Three-dataset walkthroughs.")
r=4
r=section(ws,r,"Objectives",ncols=6)
r=note(ws,r,"Create a PivotChart that updates with its pivot and choose the right chart type.",ncols=6,height=24)
r=write_table(ws,r,["Dataset","Pivot (Rows × Values)","PivotChart","When to use"],[
 ["E-commerce","order_month × Sum total_price","Line (with markers)","Trend + seasonality (Week 4 forecasting)"],
 ["Superstore","Category × Sum Profit","Column","Compare categories; add data labels"],
 ["Education","Parental_Involvement × Average Exam_Score","Bar","Compare groups; sort descending"],
],widths=[24,42,28,44])
r=note(ws,r,"Clicks: Select pivot → Insert > PivotChart → choose type. The chart and pivot stay linked: filtering the pivot filters the chart. Formatting: Chart Design > Add Chart Element (titles state the finding, e.g., 'Fashion leads revenue at ₦243m'). Never paste a static chart image — keep it live so Refresh updates both.",ncols=6,height=52)

# ---------- DATA ----------
data_sheet(wb,"12_DATA_SalaryDirty",sal_d.head(500).reset_index(drop=True),"DATA — Salary DIRTY (500-row extract)","Full DIRTY: Datasets/DIRTY_FOR_WEEK2/salary_DIRTY.csv (1,580 rows). Clean this to match 13_DATA_SalaryClean.",table_name="SalDirty500")
data_sheet(wb,"13_DATA_SalaryClean",sal.head(500).reset_index(drop=True),"DATA — Salary CLEAN (500-row extract)","Full CLEAN: Datasets/CLEAN/salary_CLEAN.csv (1,791 rows). Truth for reconciliation.",table_name="SalClean500")
data_sheet(wb,"14_DATA_SuperDirty",pd.read_csv(DIRTY/"superstore_DIRTY.csv").head(500).reset_index(drop=True),"DATA — Superstore DIRTY (500-row extract)","Full DIRTY: Datasets/DIRTY_FOR_WEEK2/superstore_DIRTY.csv. Mixed dates, $ text, spacer row.",table_name="SuperDirty500")
data_sheet(wb,"15_DATA_EcomDirty",pd.read_csv(DIRTY/"ecommerce_transactions_DIRTY.csv").head(500).reset_index(drop=True),"DATA — E-commerce DIRTY (500-row extract)","Full DIRTY: Datasets/DIRTY_FOR_WEEK2/ecommerce_transactions_DIRTY.csv. Broken keys, ₦ text, blanks.",table_name="EcomDirty500")
data_sheet(wb,"16_DATA_EduDirty",pd.read_csv(DIRTY/"student_performance_DIRTY.csv").head(500).reset_index(drop=True),"DATA — Education DIRTY (500-row extract)","Full DIRTY: Datasets/DIRTY_FOR_WEEK2/student_performance_DIRTY.csv. Ordinal mess, '23 hrs' text.",table_name="EduDirty500")
data_sheet(wb,"17_DATA_HealthDirty",pd.read_csv(DIRTY/"healthcare_DIRTY.csv").head(500).reset_index(drop=True),"DATA — Healthcare DIRTY (500-row extract)","Full DIRTY: Datasets/DIRTY_FOR_WEEK2/healthcare_DIRTY.csv. Case mess, negative billing.",table_name="HealthDirty500")

# ---------- ANSWERS ----------
ws=wb.create_sheet("18_ANSWERS")
setup_sheet(ws,"Week 2 — Answers & Marking Guide (Instructor)","Old→new module map, function answers, pivot expected values.",tab_color="C00000")
r=4
r=section(ws,r,"Old → new module map (Vol 2 renumbering)",ncols=6)
r=note(ws,r,"Old 2.1→2.1, 2.2→2.2, 2.3→2.3, old 2.4 External Web/PDF → REMOVED, old 2.4 Cleaning→2.4, old 2.5→2.5, old 2.6→2.6, old 2.7a→2.7, old 2.7b→2.8, old 2.8 M→2.9, old 2.9 Pivots→2.10, old 2.10 PivotCharts→2.11, old 2.12/2.13 → REMOVED.",ncols=6,height=42)
r=section(ws,r,"Function answers (Modules 2.7–2.8)",ncols=6)
r=note(ws,r,"Model formulas are in green columns E (2.7) and D (2.8). Mark for: correct function, correct range, correct criteria syntax (\"Delivered\" with quotes), and correct handling of text numbers. Accept XLOOKUP for VLOOKUP and FILTER for manual filtering.",ncols=6,height=42)
r=section(ws,r,"Pivot expected values (Module 2.10)",ncols=6)
r=note(ws,r,"E-commerce revenue by Category (11,000 lines): Fashion ₦243,127,900; Health ₦224,570,000; Automotive ₦219,082,900; Electronics ₦212,480,300; Home ₦177,273,900; Food ₦170,535,600; Beauty ₦155,003,700 (rounded). Superstore Sales by Category: Technology $836,154; Furniture $741,999; Office Supplies $719,047. Education Average Exam_Score by Parental_Involvement: High 68.09, Medium 67.10, Low 66.36.",ncols=6,height=52)

wb.active=wb["00_README"]
wb.save(OUT)
print(f"saved {OUT} sheets={wb.sheetnames} size={OUT.stat().st_size/1024:.0f} KB")
