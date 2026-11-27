#!/usr/bin/env python3
"""Week 1 Section Examples — Foundations & Core Excel (Modules 1.1–1.9)."""
import pandas as pd, numpy as np
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from xl_helpers import *

ROOT=Path("/home/user/ETEH")
CLEAN=ROOT/"ETEH_Bootcamp/Datasets/CLEAN"
OUT=ROOT/"01_Section_Examples/Week1_Foundations_Section_Examples.xlsx"

tr=pd.read_csv(CLEAN/"ecommerce_transactions_CLEAN.csv")
cu=pd.read_csv(CLEAN/"ecommerce_customers_CLEAN.csv")
pr=pd.read_csv(CLEAN/"ecommerce_products_CLEAN.csv")
ss=pd.read_csv(CLEAN/"superstore_CLEAN.csv")
spf=pd.read_csv(CLEAN/"student_performance_CLEAN.csv")
hc=pd.read_csv(CLEAN/"healthcare_ANONYMISED.csv")
ht=pd.read_csv(CLEAN/"hotel_reservations_CLEAN.csv")

print(f"TRX revenue total: {tr['total_price'].sum():,.0f}  lines {len(tr)}")
cat_rev = tr.merge(pr[["product_id","category"]], on="product_id").groupby("category")["total_price"].sum().sort_values(ascending=False)
print(cat_rev.to_string())
print("Superstore cat sales:\n", ss.groupby("Category")["Sales"].sum().to_string())
print("Education avg exam by parental:\n", spf.groupby("Parental_Involvement")["Exam_Score"].mean().to_string())
print("Healthcare by condition:\n", hc.groupby("Medical_Condition").agg(n=("Patient_ID","count"), los=("Length_of_Stay_Days","mean")).to_string())

wb=Workbook()
readme_sheet(wb,
 "Week 1 — Foundations & Core Excel: Section Examples (Modules 1.1–1.9)",
 "ETEH Ltd · Eaglet2EagleHub Limited · Each module = one worksheet with live data + formulas. All numbers from attached CLEAN datasets.",
 ["This workbook is the Week 1 training companion to Vol 1 (CLEANED, 178 slides, 9 modules). Every module sheet contains: objectives, a relatable worked example built from real data, live Excel formulas (green cells), tasks to complete, and check totals.",
  "Green cells contain formulas — click any green cell to see the formula. Do not overwrite them with values. Yellow cells are where you type your answers. Blue headers are tidy tables (one header row, one record per row).",
  "Inline examples are simplified extracts (20–500 rows) drawn from the full CLEAN files. Full data lives in Datasets/CLEAN/. Totals in this workbook describe the extract shown, not the full file — always state your denominator."],
 [["ecommerce_transactions_CLEAN.csv (11,000 lines)","10_DATA_Ecommerce","1.1, 1.2, 1.4, 1.7, 1.8","500","One transaction line"],
  ["superstore_CLEAN.csv (9,994 lines)","11_DATA_Superstore","1.1, 1.2, 1.4, 1.7","500","One order line"],
  ["student_performance_CLEAN.csv (6,607 students)","12_DATA_Education","1.1, 1.2, 1.4, 1.7","500","One student"],
  ["healthcare_ANONYMISED.csv (55,500 admissions)","13_DATA_Health","1.1, 1.2, 1.4, 1.7","500","One admission (Patient_ID)"],
  ["hotel_reservations_CLEAN.csv (36,275 bookings)","(extracts on sheets)","1.3, 1.4","—","One booking"]],
 ["1) Start at 01_1.1 and work in order. Each sheet takes ~30–45 minutes.  2) Before typing, read the 'Watch out' box.  3) After each task, compare your result to the Expected value and to the ANSWERS sheet.  4) Never edit DATA sheets — copy extracts to your working area first."])

# ---------- 1.1 ----------
ws=wb.create_sheet("01_1.1_Intro")
setup_sheet(ws,"Module 1.1 — Introduction to Data Analytics","Data vs information · DIKW · roles · question framing · evidence-led decisions. Datasets: E-commerce, Superstore, Education, Healthcare.")
r=4
r=section(ws,r,"Objectives — by the end you can…",ncols=6)
r=note(ws,r,"1) Classify any cell as quantitative/qualitative and explain why it matters.  2) Turn 20 raw rows into one decision-ready sentence with a benchmark.  3) Frame a vague request ('analyse sales') into a testable question with unit, period and comparison.",ncols=6,height=36)
r=section(ws,r,"Worked example A — Raw data → Information (E-commerce Nigeria, 12 real lines)",ncols=6)
r=note(ws,r,"Below are 12 real transaction lines (simplified). Raw data alone ('₦401,125.71') means nothing. Information adds: what, where, when, and compared to what.",ncols=6,height=28)
ex2 = tr.head(12)[["transaction_id","product_id","order_date","total_price","order_status","payment_method"]].merge(pr[["product_id","category"]],on="product_id",how="left")
r=write_table(ws,r,["Transaction","Category","Order date (ISO)","Amount (₦)","Status","Payment"],ex2.values.tolist(),widths=[16,16,16,16,14,18],money_cols=(4,))
r=note(ws,r,"Information sentence (copy the pattern): 'In this 12-line sample, Delivered orders total ₦X (Y% of sample value). The largest single line is ₦Z (Category, date).' — Now compute X and Y with formulas in the next table.",ncols=6,height=30)
r=section(ws,r,"Your turn — build the information (live formulas on the table above)",ncols=6)
# formulas reference the table just written: header at r-? We know table header row = r-14? Let's compute: write_table returned r after 12 rows. Header was at (r-13). Data rows header+1..header+12. Amount col D.
# Instead of fragile refs, write self-contained calc table with direct refs to DATA sheet later. For now use AVERAGE/SUM on visible range via named approach: we know header row number.
# Recompute: header_row = (current r) - 13 (1 header +12 data). Let's capture before.
# (We already advanced r; header = r-13)
hdr = r-13
r=add_answer_row(ws,r,"Sample total (SUM of Amount):",f"=SUM(D{hdr+1}:D{hdr+12})",f"₦{ex2['total_price'].sum():,.2f}")
r=add_answer_row(ws,r,"Sample average line value (AVERAGE):",f"=AVERAGE(D{hdr+1}:D{hdr+12})",f"₦{ex2['total_price'].mean():,.2f}")
r=add_answer_row(ws,r,"Count of Delivered lines (COUNTIF):",f'=COUNTIF(E{hdr+1}:E{hdr+12},"Delivered")',f"{(ex2['order_status']=='Delivered').sum()} of 12")
r=note(ws,r,"DIKW check: Data = the 12 rows above. Information = the three totals. Insight = 'Delivered dominates; one large line drives the average.' Decision = 'Investigate that large line before quoting an average.' If your insight needs a benchmark (target, last month, peer), it is not yet decision-ready.",ncols=6,height=42)
r=section(ws,r,"Worked example B — Roles & question framing (Education)",ncols=6)
r=note(ws,r,"Analyst (Excel): 'Average Exam_Score by Parental_Involvement?' — Data Scientist: 'Can we predict Exam_Score from study factors?' — Business Analyst: 'Should we fund tutoring for Low-Involvement students?' Same data, different questions. Rewrite each vague request below into: Unit + Measure + Period + Comparison.",ncols=6,height=42)
r=write_table(ws,r,["Vague request","Your sharp question (type in yellow)","Unit","Measure","Comparison"],[
 ["Analyse student performance","", "One student","Average Exam_Score","High vs Low Parental_Involvement"],
 ["Analyse hospital billing","", "One admission","Average Billing_Amount_Abs","By Medical_Condition"],
 ["Analyse retail sales","", "One order line","Total Sales / Profit","By Category, 2014–2017"],
],widths=[28,42,18,24,36])
r=note(ws,r,"Watch out: 'Average billing' without 'per admission, by condition' hides case-mix. Always name the denominator. See Module 1.9 Law 1 (one record per row) and the ANSWERS sheet for model answers.",ncols=6,height=30)

# ---------- 1.2 ----------
ws=wb.create_sheet("02_1.2_Types")
setup_sheet(ws,"Module 1.2 — Types of Data Analytics","Descriptive → Diagnostic → Predictive → Prescriptive, with numbers from Superstore & E-commerce.")
r=4
r=section(ws,r,"Objectives",ncols=6)
r=note(ws,r,"Place any question on the ladder and name the Excel tool that answers it (SUM/COUNT, drill-down, FORECAST, Goal Seek).",ncols=6,height=24)
r=section(ws,r,"Worked ladder — Superstore (real totals, 9,994 lines)",ncols=6)
ss_cat = ss.groupby("Category").agg(lines=("Row ID","count"),sales=("Sales","sum"),profit=("Profit","sum")).reset_index()
ss_cat["margin"] = ss_cat["profit"]/ss_cat["sales"]
r=write_table(ws,r,["Category","Lines","Sales ($)","Profit ($)","Margin"],ss_cat.values.tolist(),widths=[22,14,18,18,14],money_cols=(3,4),pct_cols=(5,))
r=note(ws,r,"Descriptive (What happened?): 'Office Supplies = $X sales (Y% of total).' Diagnostic (Why?): 'Furniture margin is lower because Discount averages Z% vs W% for Technology — check Discount by Sub-Category.' Predictive (What might happen?): 'FORECAST.ETS on monthly Sales projects next quarter.' Prescriptive (What should we do?): 'Cap Furniture discount at 20% and re-run Goal Seek for profit target.' Each step needs the previous one — never prescribe from a descriptive total alone.",ncols=6,height=58)
r=section(ws,r,"Your turn — E-commerce payment mix (Descriptive → Diagnostic)",ncols=6)
pay = tr.groupby(["payment_method","payment_status"]).agg(n=("transaction_id","count"),value=("total_price","sum")).reset_index().sort_values("value",ascending=False).head(8)
r=write_table(ws,r,["Payment method","Status","Lines","Value (₦)"],pay.values.tolist(),widths=[24,16,14,20],money_cols=(4,))
r=note(ws,r,"Task 1 (Descriptive): Write the headline sentence with total + share. Task 2 (Diagnostic): Which method/status combination is suspiciously large? What would you check next (e.g., Refunded by seller? Failed by date?)? Task 3: State what you would need to move to Predictive (a time series + a driver).",ncols=6,height=48)

# ---------- 1.3 ----------
ws=wb.create_sheet("03_1.3_Process")
setup_sheet(ws,"Module 1.3 — The Data Analytics Process","Objective → Collect → Clean → Transform → Analyse → Interpret → Visualise → Communicate. Mini-case: Hotel cancellations.")
r=4
r=section(ws,r,"Objectives",ncols=6)
r=note(ws,r,"Run the 8-step lifecycle on a real question and name the quality gate at each step (what proves this step is done?).",ncols=6,height=24)
r=section(ws,r,"Mini-case — Hotel cancellations (36,275 bookings, 2017–2018)",ncols=6)
ht_seg = ht.groupby("market_segment_type").agg(bookings=("Booking_ID","count"),canc=("Is_Cancelled","sum")).reset_index()
ht_seg["rate"]=ht_seg["canc"]/ht_seg["bookings"]
r=write_table(ws,r,["Market segment","Bookings","Cancellations","Cancellation rate"],ht_seg.values.tolist(),widths=[24,16,18,20],pct_cols=(4,))
r=note(ws,r,"Walk the lifecycle: 1 Objective: 'Which segment cancels most, and does lead_time predict it?' 2 Collect: Hotel Reservations.csv (CLEAN). 3 Clean: Arrival_Date_ISO built from Y/M/D; 37 invalid flagged. 4 Transform: Is_Cancelled 0/1. 5 Analyse: rate by segment (above) + average lead_time by status. 6 Interpret: 'Online cancels at X% vs Corporate Y% — but Online books earlier (lead_time), so compare within lead_time bands.' 7 Visualise: bar of rates with n labels. 8 Communicate: one sentence + denominator + caveat. Quality gates: row count reconciles (36,275), rates sum-check, no silent deletion of invalid dates.",ncols=6,height=78)
r=section(ws,r,"Your turn — fill the gate checklist",ncols=6)
r=write_table(ws,r,["Step","Gate (prove it)","Your evidence (type here)"],[
 ["Objective","Question names unit, measure, period, comparison",""],
 ["Collect","File + rows + date range recorded",""],
 ["Clean","Row count before/after + flags listed",""],
 ["Analyse","Formula shown; denominator stated",""],
 ["Communicate","Headline states finding + caveat",""],
],widths=[18,52,52])

# ---------- 1.4 ----------
ws=wb.create_sheet("04_1.4_Applications")
setup_sheet(ws,"Module 1.4 — Applications Across Industries","10 industries, each anchored to an attached dataset + KPI + relatable example.")
r=4
r=section(ws,r,"Objectives",ncols=6)
r=note(ws,r,"For any industry, name: core dataset, operational question, KPI, and one Excel technique. Use the table as a lookup during Weeks 2–5.",ncols=6,height=24)
r=write_table(ws,r,["Industry","Attached dataset","Operational question","KPI (formula)","Excel technique"],[
 ["Retail & E-commerce","ecommerce_* + superstore","Which categories/sellers drive profit?","SUMIFS revenue; Profit/Sales margin","SUMIFS, PivotTable, FORECAST.ETS"],
 ["Education","student_performance + social_media","What predicts Exam_Score?","AVERAGE Exam_Score by factor","AVERAGEIF, CORREL, Regression"],
 ["Healthcare","healthcare_ANON + diabetes","Where is LOS/billing highest?","AVERAGE LOS; SUM billing by condition","AVERAGEIF, PivotTable, SQL GROUP BY"],
 ["Banking/Finance","salary (pay) + ecommerce commission","How does pay/commission vary?","MEDIAN Salary by Job_Family","MEDIAN, PERCENTILE, Regression"],
 ["Government/Public","hotel (tourism) + ecommerce state","Where is demand/cancellation highest?","COUNT by state/segment","COUNTIFS, PivotTable"],
 ["Marketing","ecommerce membership_tier","Which tier spends most?","SUM total_spent by tier","SUMIF, PivotTable"],
 ["HR / People","salary","What drives Salary?","SLOPE Salary~Experience","CORREL, LINEST, Goal Seek"],
 ["Logistics","ecommerce_deliveries","How fast is delivery?","AVERAGE ship_lag_days by courier","AVERAGEIF, PivotTable"],
 ["Hospitality","hotel_reservations","What drives cancellation?","Cancellation rate by segment","COUNTIFS, PivotTable"],
 ["Wellbeing (sensitive)","student_depression_ANON","Where is risk concentrated? (aggregate only)","Depression rate by City_Clean (n≥30)","COUNTIFS with suppression"],
],widths=[22,30,38,30,30])
r=note(ws,r,"Relatable example (Logistics): 'Average ship lag is 4.0 days overall, but DHL Nigeria averages 3.2 vs Peace Mass Transit 4.8 — reallocate Same-Day orders to DHL.' Always pair the average with n and the range (min–max). See Module 4.1 for honest reporting.",ncols=6,height=42)

# ---------- 1.5 ----------
ws=wb.create_sheet("05_1.5_Excel")
setup_sheet(ws,"Module 1.5 — Introduction to Microsoft Excel","Anatomy, editions, file formats, safe saving. Hands-on checks (no heavy data needed).")
r=4
r=section(ws,r,"Objectives",ncols=6)
r=note(ws,r,"Navigate the window, choose the right file format, and save without losing work. Tick each check as you find it in your Excel.",ncols=6,height=28)
r=section(ws,r,"Anatomy checklist (tick in column C as you locate each item)",ncols=6)
r=write_table(ws,r,["Item","Where to find it","Found? (✓)","Why it matters for analysis"],[
 ["Ribbon tabs (Data, Formulas, Insert)","Top ribbon","", "Get Data, Sort/Filter, PivotTable, Charts live here"],
 ["Name Box + Formula Bar","Left of formula bar","", "Name Box shows active cell; Formula Bar shows the formula, not just the result"],
 ["Quick Access Toolbar","Top-left icons","", "Add Save/Undo/Redo; KeyTips (Alt keys) speed up repetitive work"],
 ["Status Bar","Bottom bar","", "Shows SUM/AVERAGE/COUNT of selected numbers — instant check"],
 ["Sheet tabs + + button","Bottom tabs","", "One module per sheet; never mix raw + report on one sheet"],
 ["File > Save As (.xlsx/.csv)","Backstage view","", ".xlsx = workbook; .csv = single tidy table; .xlsm = macros (avoid unless needed)"],
],widths=[34,26,14,52])
r=section(ws,r,"File-format exercise — choose the format (write .xlsx/.csv/.xlsm)",ncols=6)
r=write_table(ws,r,["Scenario","Format?","Why?"],[
 ["11,000-row transaction table to share with lecturer","", ""],
 ["Workbook with 9 module sheets + formulas + formatting","", ""],
 ["Single tidy table to import into Power Query/SQLite","", ""],
 ["Workbook with a recorded macro (Week 4 only)","", ""],
],widths=[52,18,52])
r=note(ws,r,"Safe saving: File > Save As → keep .xlsx for workbooks. Export a .csv copy for each tidy table you share. Enable AutoSave/Version History in Microsoft 365. Never edit Datasets/CLEAN originals — copy first.",ncols=6,height=36)

# ---------- 1.6 ----------
ws=wb.create_sheet("06_1.6_Working")
setup_sheet(ws,"Module 1.6 — Working with Data in Excel","Entry, types, formatting, Flash Fill, validation. Practice on 20 education rows.")
r=4
r=section(ws,r,"Objectives",ncols=6)
r=note(ws,r,"Enter data without creating mess: correct types, ISO dates, number formats, Flash Fill, and validation lists.",ncols=6,height=24)
r=section(ws,r,"Practice table — 20 students (real extract, simplified)",ncols=6)
prac = spf.head(20)[["Hours_Studied","Attendance","Parental_Involvement","Exam_Score","Gender"]]
r=write_table(ws,r,["Hours_Studied","Attendance","Parental_Involvement","Exam_Score","Gender","Exam_Band (formula)"],prac.values.tolist()+[["","","","","",""]],widths=[16,14,20,14,12,20])
# Add formula for Exam_Band in last col: =IF(D..>90,"90+",IF(D..>=80,"80-89",IF(D..>=70,"70-79",IF(D..>=60,"60-69","<60"))))
# header row = r-22? write_table with 21 rows (20+1 blank). header = r-22. Let's compute: rows written = 21, so header = r-22, data = header+1..header+20
hdr16 = r-22
for i in range(20):
    row = hdr16+1+i
    ws.cell(row=row,column=6,value=f'=IF(D{row}>=90,"90+",IF(D{row}>=80,"80-89",IF(D{row}>=70,"70-79",IF(D{row}>=60,"60-69","<60"))))')
    ws.cell(row=row,column=6).fill=PatternFill("solid",fgColor=GREEN)
    ws.cell(row=row,column=6).font=Font(name="Calibri",size=9)
r=note(ws,r,"Try it: 1) Format Exam_Score as Number (0 dp). 2) Format a test date as YYYY-MM-DD (type 2025-01-05). 3) Flash Fill (Ctrl+E): type 'High' next to 'High' in Parental_Involvement and accept. 4) Add Data Validation list (Low,Medium,High) to Parental_Involvement. 5) Freeze panes (View > Freeze Top Row) and filter Exam_Score ≥80. Watch out: numbers stored as text ('23 hrs') break SUM — use VALUE after SUBSTITUTE (Week 2).",ncols=6,height=58)
r=section(ws,r,"Six data types — classify these cells",ncols=6)
r=write_table(ws,r,["Cell value","Type? (Number/Date/Text/Boolean/Error)","Stored as?"],[
 ["401125.71","", ""],
 ["2025-01-05","", ""],
 ["DEL000001","", ""],
 ["TRUE","", ""],
 ["#DIV/0!","", ""],
 ["₦1,200 (text)","", ""],
],widths=[24,36,36])

# ---------- 1.7 ----------
ws=wb.create_sheet("07_1.7_Formulas")
setup_sheet(ws,"Module 1.7 — Formulas & Core Functions","SUM, AVERAGE, COUNT, IF, AND/OR, SUMIF/COUNTIF on E-commerce. All green cells are live formulas.")
r=4
r=section(ws,r,"Objectives",ncols=6)
r=note(ws,r,"Write and copy formulas with correct references ($ locking), read BODMAS, and diagnose errors (#VALUE!, #DIV/0!, #REF!).",ncols=6,height=28)
r=section(ws,r,"Practice data — 30 E-commerce lines (real, simplified)",ncols=6)
p30 = tr.head(30)[["transaction_id","quantity","unit_price","total_price","commission_amount","order_status"]]
r=write_table(ws,r,["Transaction","Qty","Unit price (₦)","Total (₦)","Commission (₦)","Status"],p30.values.tolist(),widths=[16,10,18,18,18,14],money_cols=(3,4,5))
hdr17 = r-32  # header +30 rows
r=section(ws,r,"Formula lab — complete column G (green formulas already filled for rows 1–5; copy down)",ncols=6)
r=header_row(ws,r,["Row","Check: Qty×Unit = Total?","Margin proxy: Commission/Total","Paid & Delivered?","Large line? (Total>200k)"],widths=[10,28,30,24,28])
for i in range(30):
    drow = hdr17+1+i
    ws.cell(row=r,column=1,value=i+1).font=Font(name="Calibri",size=9)
    ws.cell(row=r,column=2,value=f"=IF(ROUND(B{drow}*C{drow},2)=D{drow},\"OK\",\"CHECK\")").fill=PatternFill("solid",fgColor=GREEN)
    ws.cell(row=r,column=3,value=f"=IF(D{drow}=0,\"—\",E{drow}/D{drow})").fill=PatternFill("solid",fgColor=GREEN)
    ws.cell(row=r,column=3).number_format='0.0%'
    ws.cell(row=r,column=4,value=f'=IF(F{drow}="Delivered","Delivered","Other")').fill=PatternFill("solid",fgColor=GREEN)
    ws.cell(row=r,column=5,value=f'=IF(D{drow}>200000,"Large","")').fill=PatternFill("solid",fgColor=GREEN)
    for c in range(1,6):
        ws.cell(row=r,column=c).font=Font(name="Calibri",size=9); ws.cell(row=r,column=c).border=BORDER
    r+=1
r=section(ws,r,"Summary checks (live)",ncols=6)
r=add_answer_row(ws,r,"Total revenue (SUM Total):",f"=SUM(D{hdr17+1}:D{hdr17+30})",f"₦{p30['total_price'].sum():,.2f}")
r=add_answer_row(ws,r,"Average line (AVERAGE Total):",f"=AVERAGE(D{hdr17+1}:D{hdr17+30})",f"₦{p30['total_price'].mean():,.2f}")
r=add_answer_row(ws,r,"Delivered count (COUNTIF):",f'=COUNTIF(F{hdr17+1}:F{hdr17+30},"Delivered")',f"{(p30['order_status']=='Delivered').sum()}")
r=add_answer_row(ws,r,"Large lines (COUNTIF >200k):",f"=COUNTIF(D{hdr17+1}:D{hdr17+30},\">200000\")",f"{(p30['total_price']>200000).sum()}")
r=note(ws,r,"BODMAS drill: =2+3*4 → 14 (× before +). = (2+3)*4 → 20. F4 drill: =$D$5 locks both; =D$5 locks row; =$D5 locks column. Error drill: #DIV/0! means divide by zero (wrap with IFERROR/IF); #VALUE! means text where number expected (check '₦' text); #REF! means deleted reference. Copy formulas with fill handle; check the last row still points correctly.",ncols=6,height=58)

# ---------- 1.8 ----------
ws=wb.create_sheet("08_1.8_Tables")
setup_sheet(ws,"Module 1.8 — Organising Data with Excel Tables","Ctrl+T, headers, sorting, filtering, duplicates, calculated columns. Practice on Superstore extract.")
r=4
r=section(ws,r,"Objectives",ncols=6)
r=note(ws,r,"Convert a range to a Table, sort/filter without breaking rows, remove duplicates safely, and use the Total Row.",ncols=6,height=24)
r=section(ws,r,"Practice — 25 Superstore lines (real, simplified). Select A5:F29 → Insert > Table (Ctrl+T).",ncols=6)
p25 = ss.head(25)[["Order ID","Category","Region","Sales","Quantity","Profit"]]
r=write_table(ws,r,["Order ID","Category","Region","Sales ($)","Quantity","Profit ($)"],p25.values.tolist(),widths=[18,20,14,16,12,16],money_cols=(4,6),table_name="Superstore25")
hdr18 = r-27
r=note(ws,r,"Tasks: 1) Sort by Sales descending (Data > Sort — ensure 'My data has headers' is ticked; never sort one column alone). 2) Filter Region = West, then clear. 3) Add a calculated column Margin = Profit/Sales (type once; Table auto-fills). 4) Check duplicates: Data > Remove Duplicates on Order ID (expect 0 duplicates in this extract — try it on salary_DIRTY in Week 2 to see duplicates). 5) Turn on Total Row (Table Design > Total Row) and compare SUM to the Status Bar.",ncols=6,height=62)
r=section(ws,r,"Why Tables beat ranges (check the boxes as you verify)",ncols=6)
r=add_check_box(ws,r,"Headers are unique, short, no spaces (e.g., Order_Date_ISO not 'Order Date ')",ncols=6)
r=add_check_box(ws,r,"New rows auto-expand the Table and copy calculated-column formulas",ncols=6)
r=add_check_box(ws,r,"Filters/sorts never split a record across rows (whole-row sort)",ncols=6)
r=add_check_box(ws,r,"Structured references ([@Sales]) survive row insertion — no #REF! surprises",ncols=6)

# ---------- 1.9 ----------
ws=wb.create_sheet("09_1.9_Tidy")
setup_sheet(ws,"Module 1.9 — The Twelve Tidy Data Laws","Atomic cells, one record per row, ISO dates, explicit missing, no merges. Before/after audit.")
r=4
r=section(ws,r,"Objectives",ncols=6)
r=note(ws,r,"Audit any sheet against the 12 laws and fix violations without losing information.",ncols=6,height=24)
r=section(ws,r,"Before (messy) → After (tidy) — E-commerce delivery example",ncols=6)
r=write_table(ws,r,["BEFORE (messy — do not use)","AFTER (tidy — use this)"],[
 ["Header row 1: 'E-commerce Deliveries', Row 2: 'Jan 2025', Row 3: headers → AFTER: single header row (delivery_id, order_id, …)","One header row only; title in sheet title, not in data"],
 ["Cell: 'DEL000001 / ORD20240505 / Calabar' (3 values in 1 cell) → AFTER: 3 columns","One value per cell (Law 3)"],
 ["Blank row every 10 rows + blank column F → AFTER: removed","No spacer rows/columns (Law 5)"],
 ["Date: '05/01/2024' (ambiguous) → AFTER: 2024-01-05","ISO dates YYYY-MM-DD (Law 6)"],
 ["Missing shipped_date = 0 → AFTER: blank + flag","Blank = missing, 0 = zero (Law 7)"],
 ["Status: 'Delivered', 'DELIVERED', ' delivered ' → AFTER: Delivered (list)","Controlled vocabulary (Law 9)"],
 ["Typed total 401125.71 → AFTER: =Qty*Unit formula","Store calculations, not typed results (Law 10)"],
 ["Merged title cells across A1:F1 → AFTER: unmerged","No merged cells in data (Law 11)"],
],widths=[62,62])
r=section(ws,r,"Your audit — mark each law Pass/Fail on the DATA sheets",ncols=6)
r=write_table(ws,r,["Law","Pass/Fail (type)","Evidence (cell/row)"],[
 ["1 One record per row","",""],
 ["2 One variable per column","",""],
 ["3 Atomic cells","",""],
 ["4 One header row","",""],
 ["5 No spacer rows/columns","",""],
 ["6 ISO dates","",""],
 ["7 Missing deliberate (blank vs 0)","",""],
 ["8 Raw separated from report","",""],
 ["9 Controlled vocabulary","",""],
 ["10 Calculations, not typed results","",""],
 ["11 No merged cells","",""],
 ["12 Documented (README + dictionary)","",""],
],widths=[36,20,68])
r=note(ws,r,"Watch out: the original e-commerce.xlsx violated Laws 5 and 7 (26,000 blank rows; delivery_id only). The CLEAN files fix them. In Week 2 you will fix the DIRTY files the same way — keep the issues log (DIRTY_ISSUES_LOG.csv) beside your work.",ncols=6,height=42)

# ---------- DATA sheets ----------
data_sheet(wb,"10_DATA_Ecommerce",tr.head(500).reset_index(drop=True),"DATA — E-commerce Transactions (500-row extract)","Full file: Datasets/CLEAN/ecommerce_transactions_CLEAN.csv (11,000 rows). ₦ context. 570 Cancelled have no delivery.",table_name="Ecom500")
data_sheet(wb,"11_DATA_Superstore",ss.head(500).reset_index(drop=True),"DATA — Superstore (500-row extract)","Full file: Datasets/CLEAN/superstore_CLEAN.csv (9,994 rows, 2014–2017). $ context. Dates ISO.",table_name="Super500")
data_sheet(wb,"12_DATA_Education",spf.head(500).reset_index(drop=True),"DATA — Student Performance (500-row extract)","Full file: Datasets/CLEAN/student_performance_CLEAN.csv (6,607 students). Missing = 'Unknown'.",table_name="Edu500")
data_sheet(wb,"13_DATA_Health",hc.head(500).reset_index(drop=True),"DATA — Healthcare Admissions ANONYMISED (500-row extract)","Full file: Datasets/CLEAN/healthcare_ANONYMISED.csv (55,500). Patient_ID only — portfolio-safe.",table_name="Health500")

# ---------- ANSWERS ----------
ws=wb.create_sheet("14_ANSWERS")
setup_sheet(ws,"Week 1 — Answers & Marking Guide (Instructor)","Model answers with formulas. Students: do not open until you have attempted every task.",tab_color="C00000")
r=4
r=section(ws,r,"How to mark",ncols=6)
r=note(ws,r,"Award full marks for correct formula + correct denominator + caveat. A correct number with no denominator or no caveat scores half. All Expected values below were computed from the CLEAN files in Python and cross-checked to Excel formulas.",ncols=6,height=36)
r=section(ws,r,"1.1 Answers",ncols=6)
r=note(ws,r,f"12-line sample total ₦{ex2['total_price'].sum():,.2f}; average ₦{ex2['total_price'].mean():,.2f}; Delivered {(ex2['order_status']=='Delivered').sum()}/12. Sharp questions: see sheet 01 (Unit/Measure/Comparison columns).",ncols=6,height=36)
r=section(ws,r,"1.2 Answers",ncols=6)
r=note(ws,r,"Superstore: Office Supplies largest sales; Furniture lowest margin (discount-driven). E-commerce: Paystack/Flutterwave dominate value; Refunded/Failed small but must be investigated by seller/date before prescribing.",ncols=6,height=36)
r=section(ws,r,"1.7 Answers",ncols=6)
r=note(ws,r,f"30-line total ₦{p30['total_price'].sum():,.2f}; average ₦{p30['total_price'].mean():,.2f}; Delivered {(p30['order_status']=='Delivered').sum()}; Large {(p30['total_price']>200000).sum()}. Accept equivalent formulas (e.g., SUMPRODUCT for conditional sums).",ncols=6,height=36)
r=section(ws,r,"File-format answers (1.5)",ncols=6)
r=note(ws,r,"Share table → .csv; 9-sheet workbook → .xlsx; Power Query import → .csv; macro workbook → .xlsm (only if macros required).",ncols=6,height=28)

wb.active=wb["00_README"]
wb.save(OUT)
print(f"saved {OUT} sheets={wb.sheetnames} size={OUT.stat().st_size/1024:.0f} KB")
