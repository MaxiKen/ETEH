#!/usr/bin/env python3
"""
ETEH Bootcamp - Workbook builder (the four operation folders)
Builds, per week, a Section-Examples workbook (one worksheet per module) plus
Assignments, Instructor Projects and Student Projects workbooks. All example
tables are small, relatable extracts drawn from the supplied/standardised
datasets (ETEH Retail, Lifecare, Commerce, University, Hotels, Salary).
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = "/home/user/ETEH/ETEH_Bootcamp"
SEC  = os.path.join(ROOT, "section_examples")
ASG  = os.path.join(ROOT, "assignments")
PIN  = os.path.join(ROOT, "projects_instructor")
PST  = os.path.join(ROOT, "projects_students")
for p in (SEC, ASG, PIN, PST): os.makedirs(p, exist_ok=True)

NAVY  = "0B3D61"; LT = "DCE6EF"; GREY="EEF3F7"; GOLD="C9962B"
HFONT = Font(bold=True, color="FFFFFF", size=10)
TFONT = Font(bold=True, color="FFFFFF", size=13)
SFONT = Font(bold=True, color=NAVY, size=11)
BFONT = Font(size=10)
thin  = Side(style="thin", color="B8C4CC")
BORDER= Border(left=thin,right=thin,top=thin,bottom=thin)
fill_navy = PatternFill("solid", fgColor=NAVY)
fill_lt   = PatternFill("solid", fgColor=LT)
fill_grey = PatternFill("solid", fgColor=GREY)
fill_gold = PatternFill("solid", fgColor=GOLD)

def title(ws, text, ncols, sub=None):
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=ncols)
    c=ws.cell(1,1,text); c.font=TFONT; c.fill=fill_navy; c.alignment=Alignment(vertical="center",horizontal="left",indent=1)
    ws.row_dimensions[1].height=26
    r=2
    if sub:
        ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=ncols)
        c=ws.cell(2,1,sub); c.font=Font(italic=True,size=9,color="555555"); c.alignment=Alignment(indent=1)
        r=3
    return r

def auto_total_col(table):
    kw=('sales','profit','total','billing','revenue','price','salary','amount')
    for idx,h in enumerate(table[0]):
        if any(k in str(h).lower() for k in kw): return idx+1
    return None

def write_block(ws, row, heading, table, note=None, total=None):
    """table: list of rows, first is header. total: column index (1-based) to sum, or None (auto)."""
    ncols=len(table[0])
    if total is None: total=auto_total_col(table)
    ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=ncols)
    c=ws.cell(row,1,heading); c.font=SFONT; c.fill=fill_lt
    row+=1
    # header
    for j,h in enumerate(table[0],1):
        c=ws.cell(row,j,h); c.font=HFONT; c.fill=fill_navy; c.border=BORDER; c.alignment=Alignment(horizontal="center",wrap_text=True)
    row+=1
    data_start=row
    for ri,rw in enumerate(table[1:], start=0):
        for j,v in enumerate(rw,1):
            c=ws.cell(row,j,v); c.font=BFONT; c.border=BORDER
            if ri%2==1: c.fill=fill_grey
            if isinstance(v,(int,float)) and not isinstance(v,bool): c.number_format='#,##0.00' if isinstance(v,float) else '#,##0'
        row+=1
    if total is not None and len(table)>1:
        trow=row
        ws.cell(trow,1,"TOTAL").font=Font(bold=True)
        col=get_column_letter(total)
        f=f"=SUM({col}{data_start}:{col}{trow-1})"
        c=ws.cell(trow,total,f); c.font=Font(bold=True); c.number_format='#,##0.00'; c.fill=fill_gold
        for j in range(1,ncols+1):
            ws.cell(trow,j).border=BORDER; 
            if j!=total: ws.cell(trow,j).fill=fill_gold
        row=trow+1
    if note:
        ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=ncols)
        c=ws.cell(row,1,"💡 "+note); c.font=Font(italic=True,size=9,color="33475B"); c.alignment=Alignment(wrap_text=True,vertical="top")
        ws.row_dimensions[row].height=30
        row+=1
    return row+1

# ---------------------------------------------------------------------------
# Toy / relatable example tables (drawn from dataset ideas)
# ---------------------------------------------------------------------------
T_RETAIL=[["Order_ID","Region","Category","Sales (₦)","Profit (₦)"],
 ["ORD-1001","West","Technology",1200,320],["ORD-1002","East","Furniture",850,-40],
 ["ORD-1003","South","Office Supplies",430,95],["ORD-1004","West","Furniture",1500,410],
 ["ORD-1005","Central","Technology",2200,760],["ORD-1006","East","Office Supplies",310,60],
 ["ORD-1007","South","Technology",980,250],["ORD-1008","West","Furniture",1750,520],
 ["ORD-1009","Central","Office Supplies",540,120],["ORD-1010","East","Technology",1330,390]]

T_RETAIL_PIVOT=[["Month","Region","Category","Sales (₦)"],
 ["Jan","West","Technology",3200],["Jan","East","Furniture",1850],["Jan","South","Office Supplies",980],
 ["Feb","West","Technology",4100],["Feb","East","Furniture",2200],["Feb","South","Office Supplies",1320],
 ["Mar","West","Technology",3650],["Mar","East","Furniture",1980],["Mar","South","Office Supplies",1150]]

T_HOSPITAL=[["Patient","Condition","Admission","Discharge","Billing (₦)"],
 ["A. Bello","Hypertension","2024-01-04","2024-01-09",245000],
 ["C. Eze","Diabetes","2024-02-11","2024-02-15",188500],
 ["M. Okoro","Cancer","2024-03-02","2024-03-20",1240000],
 ["T. Sani","Asthma","2024-03-18","2024-03-21",96000],
 ["N. Ade","Arthritis","2024-04-05","2024-04-09",132000],
 ["K. Musa","Hypertension","2024-04-22","2024-04-25",78000],
 ["F. Lawal","Diabetes","2024-05-10","2024-05-16",210000],
 ["S. Idris","Cancer","2024-06-01","2024-06-12",980000]]

T_COMMERCE=[["Transaction","Customer","City","Category","Qty","Unit_Price (₦)","Total (₦)","Seller","Payment"],
 ["TRX00012","Funmilayo Buhari","Kano","Electronics",2,49121,98242,"Sokoto Suppliers","Flutterwave"],
 ["TRX00045","Amara Attah","Ikorodu","Health",1,55024,55024,"Eko Mall 3","Bank Transfer"],
 ["TRX00071","Ade Eze","Onitsha","Automotive",3,29000,87000,"Oshogbo Stores","Flutterwave"],
 ["TRX00103","Ngozi Okafor","Abuja","Fashion",1,116752,116752,"Kano Mart","Card"],
 ["TRX00140","Yusuf Bello","Ibadan","Electronics",2,33000,66000,"Sokoto Suppliers","Flutterwave"],
 ["TRX00188","Halima Sule","Kaduna","Home",4,18000,72000,"Oshogbo Stores","Bank Transfer"],
 ["TRX00220","Emeka Nwosu","Enugu","Health",2,27500,55000,"Eko Mall 3","Card"],
 ["TRX00265","Zainab Bakare","Minna","Automotive",1,60400,60400,"Kano Mart","Flutterwave"]]

T_UNI=[["Student","Hours_Studied","Attendance %","Prev_Score","Exam_Score"],
 ["S01",23,84,73,67],["S02",15,70,65,58],["S03",30,95,88,91],["S04",8,55,50,44],
 ["S05",19,78,70,63],["S06",26,90,82,85],["S07",12,60,58,52],["S08",34,98,90,94],
 ["S09",17,72,64,60],["S10",21,80,75,71]]

T_SALARY=[["Education","Years_Experience","Salary (₦)"],
 ["High School",1,2800000],["Bachelor's",2,4200000],["Bachelor's",5,6500000],
 ["Master's",3,7800000],["Master's",8,12500000],["PhD",4,11000000],["PhD",10,19000000],
 ["Bachelor's",12,14000000]]

T_HOTELS=[["Booking","Room_Type","Segment","Lead_Time","Avg_Price (₦)","Status"],
 ["BK-501","Room_Type 1","Online",45,65000,"Canceled"],["BK-502","Room_Type 3","Offline",12,82000,"Not Cancelled"],
 ["BK-503","Room_Type 2","Corporate",30,74000,"Not Cancelled"],["BK-504","Room_Type 4","Online",60,91000,"Canceled"],
 ["BK-505","Room_Type 1","Aviation",8,68000,"Not Cancelled"],["BK-506","Room_Type 5","Online",22,88000,"Not Cancelled"],
 ["BK-507","Room_Type 3","Offline",50,80000,"Canceled"],["BK-508","Room_Type 2","Corporate",18,76000,"Not Cancelled"]]

T_TIDY_MESSY=[["Name","Math,Science,English"],["Ada","70,82,65"],["Ben","55,61,73"],["Chi","88,90,84"]]
T_TIDY=[["Name","Subject","Score"],["Ada","Math",70],["Ada","Science",82],["Ada","English",65],
 ["Ben","Math",55],["Ben","Science",61],["Ben","English",73],["Chi","Math",88],["Chi","Science",90],["Chi","English",84]]

T_METRICS=[["Grain","Count"],["Transaction lines",35712],["Distinct receipts",28940],["Distinct customers",1240],["Distinct products",2120]]
T_TS=[["Month","Revenue (₦)"],["Jan",1820000],["Feb",2100000],["Mar",1950000],["Apr",2300000],
 ["May",2480000],["Jun",2620000],["Jul",2550000],["Aug",2710000]]
T_SURVEY=[["Student","Study_Hours","Sleep_Hours","Exam_Score","Internet"],
 ["S01",23,7,67,"Yes"],["S02",15,6,58,"Yes"],["S03",30,8,91,"No"],["S04",8,5,44,"Yes"],
 ["S05",19,7,63,"Yes"],["S06",26,8,85,"No"],["S07",12,6,52,"Yes"],["S08",34,9,94,"No"]]

# ---------------------------------------------------------------------------
# Module definitions: (code, name, goal, [tables], note override optional)
# ---------------------------------------------------------------------------
def M(code,name,goal,tables): return (code,name,goal,tables)

W1 = [
 M("1.1","Introduction to Data Analytics",
   "Distinguish raw data, information, insight and decision; read quantitative vs qualitative values.",
   [(T_RETAIL,"Raw line vs decision-ready summary",None,
     "ORD-1002 | 2 | ₦850 is raw data; 'East Furniture averaged ₦1,020 profit in Q1' is information.")]),
 M("1.2","Types of Data Analytics",
   "Describe descriptive, diagnostic, predictive and prescriptive analytics with a retail example.",
   [(T_RETAIL_PIVOT,"Descriptive → Diagnostic → Predictive",
     "Describe WHAT sold (pivot), diagnose WHY (East furniture dipped), predict NEXT month.")]),
 M("1.3","The Data Analytics Process",
   "Walk the ask→prepare→process→analyse→share→act loop on a hospital waitlist.",
   [(T_HOSPITAL,"One analytics cycle on ETEH Lifecare",5,
     "Ask: which condition costs most? Prepare billing → analyse totals → share → act on Cancer spend.")]),
 M("1.4","Applications Across Industries",
   "Map a business question to a dataset in retail, health, education and hospitality.",
   [(T_RETAIL,"Retail"),(T_HOSPITAL,"Health"),(T_UNI,"Education"),(T_HOTELS,"Hospitality")]),
 M("1.5","Introduction to Microsoft Excel",
   "Learn the grid, ribbon, cell references and the formula bar with a tiny table.",
   [(T_SALARY,"Enter & reference cells",3,
     "Type =B2 in a helper cell to pull Years_Experience for the first row.")]),
 M("1.6","Working with Data in Excel",
   "Select, sort, filter and freeze panes on the ETEH Retail extract.",
   [(T_RETAIL,"Sort by Profit, filter Region=West",5)]),
 M("1.7","Excel Formulas and Core Functions",
   "Use SUM, AVERAGE, IF, VLOOKUP/XLOOKUP on relatable data.",
   [(T_RETAIL,"SUM & AVERAGE practice",4,
     "Add =SUM(D2:D11) and =AVERAGE(E2:E11); use IF(Profit<0,'Loss','Gain').")]),
 M("1.8","Organising Data with Excel Tables",
   "Turn a range into a Table (Ctrl+T) and use structured references.",
   [(T_RETAIL,"Format as Table, then =SUM(Table1[Sales (₦)])",4)]),
 M("1.9","The Twelve Tidy Data Laws",
   "Each variable a column, each observation a row, one value per cell — messy vs tidy.",
   [(T_TIDY_MESSY,"Messy wide form (do not analyse like this)",None,
     "Scores bundled in one cell break sorting & maths."),
    (T_TIDY,"Tidy long form (one value per cell)",3)]),
 M("1.11","Measurement, Quality and Responsible Analysis",
   "State the grain before you sum; lines ≠ receipts.",
   [(T_METRICS,"Count at the correct grain",2,
     "Sum Total_Price for sales; COUNT(DISTINCT receipt) for tickets.")]),
]

W2 = [
 M("2.1","Understanding External Data Sources",
   "Know file types (CSV/XLSX/PDF/DB) and where ETEH datasets live (datasets/clean).",
   [(T_COMMERCE,"A relational extract to import",7)]),
 M("2.2","Introduction to Power Query",
   "Meet the Query Editor: applied steps record every transformation.",
   [(T_RETAIL,"Source → transform → load",4)]),
 M("2.3","Connecting and Importing with Power Query",
   "Import eteh_retail.xlsx and eteh_commerce_flat.csv; set data types.",
   [(T_RETAIL,"Imported retail extract",5)]),
 M("2.4","Data Cleaning with Excel",
   "Trim spaces, fix cases, remove duplicates with Excel tools.",
   [(T_HOTELS,"Clean mixed-case Status & Segment",6,
     "Standardise 'canceled'→'Canceled'; trim Room_Type spacing.")]),
 M("2.5","Cleaning and Transforming in Power Query",
   "Use PQ to unpivot, split columns and replace values reproducibly.",
   [(T_TIDY_MESSY,"Unpivot the bundled scores",None),
    (T_TIDY,"Result: tidy long table",3)]),
 M("2.6","Excel Cleaning vs Power Query Cleaning",
   "When to use each; PQ wins for repeatable pipelines.",
   [(T_RETAIL,"Same data, two paths",4)]),
 M("2.7","Text, Date, Lookup and Dynamic Array Functions",
   "TEXT, DATEVALUE, XLOOKUP, FILTER, UNIQUE on relatable rows.",
   [(T_COMMERCE,"XLOOKUP a seller's total",8,
     "=XLOOKUP(\"Sokoto Suppliers\", Sellers, Totals) returns ₦164,242.")]),
 M("2.8","Introduction to the Power Query M Language",
   "Read and tweak M: Table.SelectRows, each, let..in.",
   [(T_COMMERCE,"Filter M: keep Payment='Flutterwave'",7)]),
 M("2.9","Introduction to PivotTables",
   "Build a pivot: Region & Category → SUM of Sales.",
   [(T_RETAIL_PIVOT,"Pivot source (month/region/category/sales)",4,
     "Rows=Region, Cols=Category, Values=SUM(Sales).")]),
 M("2.10","Introduction to PivotCharts",
   "Turn the pivot into a chart; choose the right visual.",
   [(T_RETAIL_PIVOT,"Chart the pivot",4)]),
 M("2.13","Repeatable Power Query and Reconciliation",
   "Name steps, refresh predictably, and reconcile to source totals.",
   [(T_COMMERCE,"Refresh & reconcile to ₦660,418",7,
     "After cleaning, totals must still match the raw ₦660,418.")]),
]

W3 = [
 M("3.1","Database Fundamentals",
   "Tables, rows, columns, keys and relationships (the ETEH Commerce model).",
   [(T_COMMERCE,"A joined fact row (transactions↔products↔customers)",7)]),
 M("3.2","Introduction to SQL and SQLite",
   "SELECT, WHERE, GROUP BY, JOIN on eteh_commerce.db.",
   [(T_COMMERCE,"Rows to query",7,
     "SELECT Category, SUM(Total) FROM ... GROUP BY Category.")]),
 M("3.3","Connecting Excel to Databases with ODBC",
   "Pull a SQL result into Excel via ODBC / Power Query.",
   [(T_COMMERCE,"Result set to load",7)]),
 M("3.4","Merging and Appending in Power Query",
   "Merge transactions with products; append monthly files.",
   [(T_COMMERCE,"Merge on Product key",7)]),
 M("3.5","The Excel Data Model and Star Schema",
   "Fact (transactions) + dimensions (customers, products, sellers).",
   [(T_COMMERCE,"Star: fact surrounded by dimensions",7)]),
 M("3.6","Power Pivot and DAX (Time Intelligence, KPIs)",
   "Write measures: Total Sales, MoM %, a KPI.",
   [(T_RETAIL_PIVOT,"DAX over this matrix",4,
     "Total Sales = SUM('Sales'[Sales]); MoM% = DIVIDE([This],[Prev])-1.")]),
 M("3.7","Chart Anatomy and the 17 Excel Chart Types",
   "Pick charts by question; label axes; avoid chart junk.",
   [(T_RETAIL,"Bar for category comparison",4)]),
 M("3.8","Dashboard Development",
   "Lay out KPIs, a pivot chart and a slicer on one canvas.",
   [(T_RETAIL_PIVOT,"Dashboard feed",4),(T_COMMERCE,"Commerce KPI feed",7)]),
 M("3.9","Three Dashboard Blueprints",
   "KPI scorecard, trend+breakdown, geographic/segment view.",
   [(T_RETAIL_PIVOT,"Blueprint A: trend+breakdown",4)]),
 M("3.10","External Data in SQL, DAX and Dashboards",
   "Mix DB, CSV and API sources in one model; retargeted to ETEH datasets.",
   [(T_COMMERCE,"Multi-source model",7)]),
 M("3.12","Database Design, Query Reliability and BI Governance",
   "Declare grain & keys; test dimension uniqueness before modelling.",
   [(T_COMMERCE,"Keys that must be unique",7,
     "Product_ID and Customer_ID must be unique or totals inflate.")]),
]

W4 = [
 M("4.1","Descriptive Statistics and the Analysis ToolPak",
   "Mean, median, std-dev, min/max on exam scores.",
   [(T_UNI,"Describe Exam_Score",5,
     "Mean≈68.5, median≈65 — a right-skew from top students.")]),
 M("4.2","Distributions, Pareto, Percentiles and Z-Scores",
   "Histogram of billing; Pareto of conditions; z-score a patient.",
   [(T_HOSPITAL,"Billing distribution & Pareto",5)]),
 M("4.3","Correlation Analysis",
   "Correlate study hours / sleep with exam score.",
   [(T_SURVEY,"Correlation matrix practice",4,
     "Study_Hours vs Exam_Score is strongly positive; build =CORREL().")]),
 M("4.4","Linear Regression",
   "Predict Salary from Years_Experience (simple regression).",
   [(T_SALARY,"X=Experience, Y=Salary",3,
     "Fit y = a + b·x; each extra year ≈ +₦1.1m salary.")]),
 M("4.5","Time Series and Forecasting",
   "Plot revenue, add trendline / moving average, forecast next month.",
   [(T_TS,"8 months → forecast month 9",2,
     "Linear trend forecasts ≈ ₦2.78m for August.")]),
 M("4.6","Goal Seek, Data Tables, Scenarios and Solver",
   "Goal Seek the price that hits a profit target; Solver optimise mix.",
   [(T_RETAIL,"Optimise the basket",4)]),
 M("4.7","Data Storytelling",
   "Lead with the insight, not the chart; annotate the 'so what'.",
   [(T_RETAIL_PIVOT,"Tell the East-Furniture story",4)]),
 M("4.8","Report Writing and PDF Production",
   "Structure: exec summary → method → findings → appendix; export PDF.",
   [(T_HOSPITAL,"Findings to write up",5)]),
 M("4.9","Slide Presentations",
   "One idea per slide; the 10-minute executive defence.",
   [(T_COMMERCE,"Deck outline from this data",7)]),
 M("4.10","Formula Auditing, Protection and Light Automation",
   "Trace precedents, lock inputs, record a simple macro.",
   [(T_SALARY,"Audit the salary model",3)]),
 M("4.11","AI Co-pilots for Analysts",
   "Use Copilot/chat to draft a pivot or SQL; always verify.",
   [(T_COMMERCE,"Prompt → SQL on Commerce",7)]),
 M("4.12","Verifying AI Output and Handling Hallucinations",
   "Reconcile AI totals to source; catch fabricated rows.",
   [(T_COMMERCE,"Check the ₦660,418 total",7)]),
 M("4.13","Data Ethics, PII and PHI",
   "Anonymise patient/student files; follow Nigeria Data Protection Act 2023.",
   [(T_HOSPITAL,"PHI to mask",5,
     "Replace Patient name with a hashed ID before sharing.")]),
 M("4.14","External Benchmarks, PDF Extraction and Reconciliation",
   "Extract a PDF table, benchmark, and reconcile to source.",
   [(T_RETAIL,"Benchmark vs industry",4)]),
 M("4.16","Statistical Reasoning, Forecast Evaluation and Model Selection",
   "State grain/unit before summarising; evaluate a forecast honestly.",
   [(T_TS,"Evaluate the August forecast",2,
     "Compare forecast ₦2.78m to actual; report MAPE, not just R².")]),
]

WEEKS = [("1",W1),("2",W2),("3",W3),("4",W4)]

# ---------------------------------------------------------------------------
# Build Section Examples workbooks
# ---------------------------------------------------------------------------
def build_section_examples():
    for wk, mods in WEEKS:
        wb=Workbook(); wb.remove(wb.active)
        cover=wb.create_sheet("Cover")
        r=title(cover, f"ETEH Bootcamp — Week {wk} · Section Examples", 2,
                "One worksheet per module: a relatable, worked example drawn from the standardised ETEH datasets.")
        cover.cell(r,1,"Module").font=SFONT; cover.cell(r,2,"Name").font=SFONT
        rr=r+1
        for code,name,goal,tables in mods:
            ws=wb.create_sheet(f"M{code}")
            row=title(ws, f"Week {wk} · Module {code} — {name}", max(len(t[0]) for t in tables))
            ws.cell(row,1,"Goal: "+goal).font=Font(italic=True,size=9,color="33475B")
            ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=max(len(t[0]) for t in tables))
            row+=2
            for item in tables:
                table, head = item[0], item[1]
                note = None
                for x in item[2:]:           # legacy: ignore stray int total-hints, keep string note
                    if isinstance(x, str): note = x
                row=write_block(ws,row,head,table,note=note)
            # auto width
            for col in ws.columns:
                ml=max((len(str(c.value)) for c in col if c.value is not None), default=8)
                ws.column_dimensions[get_column_letter(col[0].column)].width=min(max(ml+2,10),34)
            cover.cell(rr,1,code).font=BFONT; cover.cell(rr,2,name).font=BFONT; rr+=1
        cover.cell(rr+1,1,"Datasets referenced: eteh_retail, eteh_lifecare, eteh_commerce, eteh_university, eteh_hotels, eteh_salary").font=Font(italic=True,size=8,color="777777")
        cover.column_dimensions['A'].width=10; cover.column_dimensions['B'].width=70
        out=os.path.join(SEC, f"Week{wk}_Section_Examples.xlsx")
        wb.save(out); print("wrote", out)

# ---------------------------------------------------------------------------
# Build Assignments workbooks (Cover + Brief + Practice_Data + Rubric)
# ---------------------------------------------------------------------------
ASSIGN = {
 "1":[("A1.1","Tidy the ETEH Retail extract","Reformat T_RETAIL_MESSY to one-value-per-cell; compute totals.",T_TIDY_MESSY),
      ("A1.2","Build a profit summary","SUM/AVERAGE profit by region from T_RETAIL.",T_RETAIL),
      ("A1.3","Grain audit","Explain why 35,712 lines ≠ 35,712 receipts (use T_METRICS).",T_METRICS)],
 "2":[("A2.1","Clean the Hotels export","Standardise Status/Segment; trim Room_Type (T_HOTELS).",T_HOTELS),
      ("A2.2","Pivot the retail extract","Region × Category SUM(Sales) from T_RETAIL_PIVOT.",T_RETAIL_PIVOT),
      ("A2.3","Unpivot scores","Convert T_TIDY_MESSY to tidy long form.",T_TIDY_MESSY)],
 "3":[("A3.1","SQL: revenue by category","Write the GROUP BY on eteh_commerce.db.",T_COMMERCE),
      ("A3.2","Star schema","Diagram fact+dimensions for Commerce.",T_COMMERCE),
      ("A3.3","DAX measure","Total Sales & MoM% over T_RETAIL_PIVOT.",T_RETAIL_PIVOT)],
 "4":[("A4.1","Regression","Fit Salary~Experience (T_SALARY); report slope.",T_SALARY),
      ("A4.2","Correlation","CORREL study hours vs score (T_SURVEY).",T_SURVEY),
      ("A4.3","Forecast","Trendline + 1-month forecast on T_TS.",T_TS)],
}
def build_assignments():
    for wk in ["1","2","3","4"]:
        wb=Workbook(); wb.remove(wb.active)
        cover=wb.create_sheet("Brief")
        r=title(cover, f"ETEH Bootcamp — Week {wk} · Assignments", 3,
                "Complete every task; submit the worked workbook. Practice data lives in the 'Practice_Data' sheet.")
        for j,h in enumerate(["Task","Brief","Source table"]): cover.cell(r,j+1,h).font=HFONT; cover.cell(r,j+1).fill=fill_navy
        rr=r+1
        for tid,brief,desc,src in ASSIGN[wk]:
            cover.cell(rr,1,tid).font=BFONT; cover.cell(rr,2,brief).font=BFONT; cover.cell(rr,3,desc).font=BFONT
            rr+=1
        # practice data sheet
        pdws=wb.create_sheet("Practice_Data")
        pr=title(pdws,"Practice data (work here)", max((len(t[0]) for _,_,_,t in ASSIGN[wk]), default=5))
        for tid,brief,desc,src in ASSIGN[wk]:
            pr=write_block(pdws,pr,f"{tid} — {desc}",src,total=(4 if len(src[0])>=4 else None))
        # rubric
        rb=wb.create_sheet("Rubric")
        rb.title="Rubric"
        r2=title(rb,"Marking rubric (per task)",4)
        for j,h in enumerate(["Criterion","Weight %","Excellent (5)","Poor (1)"]):
            rb.cell(r2,j+1,h).font=HFONT; rb.cell(r2,j+1).fill=fill_navy
        rows=[["Correctness of result",40,"Matches reconciled total","Wrong totals"],
              ["Tidy / clean structure",25,"One value per cell","Messy merged cells"],
              ["Formula/SQL reuse",20,"Live formulas, no hardcode","Hardcoded numbers"],
              ["Presentation & notes",15,"Clear labels & insight","Unlabelled"]]
        for i,row in enumerate(rows):
            for j,v in enumerate(row):
                c=rb.cell(r2+1+i,j+1,v); c.font=BFONT; c.border=BORDER; c.fill=fill_grey if i%2 else fill_grey
        rb.column_dimensions['A'].width=26; rb.column_dimensions['B'].width=10
        rb.column_dimensions['C'].width=34; rb.column_dimensions['D'].width=20
        cover.column_dimensions['A'].width=10; cover.column_dimensions['B'].width=60; cover.column_dimensions['C'].width=22
        out=os.path.join(ASG, f"Week{wk}_Assignments.xlsx")
        wb.save(out); print("wrote", out)

# ---------------------------------------------------------------------------
# Instructor & Student project workbooks
# ---------------------------------------------------------------------------
def proj_sheet(ws, title_txt, brief, steps, deliverables, rubric, ncols=6):
    r=title(ws, title_txt, ncols, brief)
    ws.cell(r,1,"Worked steps / Solution").font=SFONT; r+=1
    for s in steps:
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=ncols)
        c=ws.cell(r,1,"• "+s); c.font=BFONT; c.alignment=Alignment(wrap_text=True,vertical="top"); ws.row_dimensions[r].height=28; r+=1
    r+=1
    ws.cell(r,1,"Deliverables").font=SFONT; r+=1
    for d in deliverables:
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=ncols)
        c=ws.cell(r,1,"– "+d); c.font=BFONT; r+=1
    r+=1
    ws.cell(r,1,"Rubric").font=SFONT; r+=1
    for j,h in enumerate(["Criterion","Weight %","Excellent","Poor"]):
        ws.cell(r,j+1,h).font=HFONT; ws.cell(r,j+1).fill=fill_navy
    r+=1
    for i,row in enumerate(rubric):
        for j,v in enumerate(row):
            c=ws.cell(r,j+1,v); c.font=BFONT; c.border=BORDER; c.fill=fill_grey
        r+=1
    return r

def build_projects():
    # INSTRUCTOR
    wb=Workbook(); wb.remove(wb.active)
    cov=wb.create_sheet("Index")
    title(cov,"ETEH Bootcamp — Instructor Project Pack (worked solutions)",2,
          "Full solutions the instructor models live. Students get the same briefs in projects_students/.")
    r=3
    for code,name in [("P1","Retail Pivot & Dashboard"),("P2","Commerce SQL Joins"),("P3","Hospital Stats & Report")]:
        cov.cell(r,1,code).font=BFONT; cov.cell(r,2,name).font=BFONT; r+=1
    cov.column_dimensions['A'].width=8; cov.column_dimensions['B'].width=40
    ws=wb.create_sheet("P1_Retail_Pivot")
    proj_sheet(ws,"Project 1 — ETEH Retail Pivot & Dashboard (INSTRUCTOR SOLUTION)",
      "Build a PivotTable of Sales by Region & Category, a PivotChart, and a one-screen dashboard with a slicer. Source: datasets/clean/eteh_retail.xlsx.",
      ["Load eteh_retail.xlsx as a Table.","Insert PivotTable: Rows=Region, Cols=Category, Values=SUM(Sales), SUM(Profit).",
       "Add a PivotChart (clustered column).","Add KPI cards (total sales, total profit, order count) with CUBE/GETPIVOTDATA.",
       "Insert a Slicer on Region; format with the ETEH navy palette.","Reconcile Grand Total Sales to source (should match exactly)."],
      ["Dashboard .xlsx with slicer","2 KPI cards + 1 chart","1-paragraph insight on East Furniture"],
      [["Correct totals",40,"Matches source","Wrong"],["Visual design",25,"Clean, labelled","Cluttered"],["Insight",20,"Actionable","Generic"],["Reuse",15,"Live pivot","Hardcoded"]])
    ws=wb.create_sheet("P2_Commerce_SQL")
    proj_sheet(ws,"Project 2 — ETEH Commerce SQL Joins (INSTRUCTOR SOLUTION)",
      "Answer 3 questions with SQL on datasets/sql/eteh_commerce.db. Joins across transactions↔products↔customers↔sellers↔deliveries.",
      ["Q1 Monthly revenue: SELECT substr(order_date,1,7), SUM(total_price) ... GROUP BY month.",
       "Q2 Top category: JOIN transactions↔products, GROUP BY category ORDER BY revenue DESC.",
       "Q3 CLV: JOIN transactions↔customers, SUM(total_price) per customer.",
       "Verify each total reconciles to the flat extract (eteh_commerce_flat.csv)."],
      ["queries.sql with 3 queries","Result tables (CSV)","1-page interpretation"],
      [["Correct SQL",45,"Valid JOINs","Errors"],["Reconciles",30,"Matches flat","Off"],["Interpretation",25,"Clear","Missing"]])
    ws=wb.create_sheet("P3_Hospital_Stats")
    proj_sheet(ws,"Project 3 — ETEH Lifecare Statistics & Report (INSTRUCTOR SOLUTION)",
      "Summarise billing by condition, build a Pareto, and write a 1-page PDF report. Source: datasets/clean/eteh_lifecare.xlsx.",
      ["Compute total & mean billing per Medical Condition.","Rank conditions; compute cumulative % for Pareto.",
       "Mask Patient names (PII) before any shareable output.","Write report: exec summary → method → findings → recommendation."],
      ["Summary table","Pareto chart","Anonymised 1-page PDF"],
      [["Stats correct",40,"Right aggregates","Wrong"],["Privacy",25,"Anonymised","Leaks PII"],["Report",25,"Structured","Sloppy"],["Visual",10,"Clear","Missing"]])
    wb.save(os.path.join(PIN,"Instructor_Projects.xlsx")); print("wrote Instructor_Projects.xlsx")

    # STUDENT
    wb=Workbook(); wb.remove(wb.active)
    cov=wb.create_sheet("Index")
    title(cov,"ETEH Bootcamp — Student Capstone Projects",2,
          "Pick ONE to submit as your capstone. Datasets are in datasets/clean. Reference DATASET_GUIDE.md.")
    r=3
    for code,name in [("C1","Retail Sales Analysis"),("C2","Commerce SQL Deep-Dive"),("C3","University & Health Insight")]:
        cov.cell(r,1,code).font=BFONT; cov.cell(r,2,name).font=BFONT; r+=1
    cov.column_dimensions['A'].width=8; cov.column_dimensions['B'].width=42
    ws=wb.create_sheet("C1_Retail")
    proj_sheet(ws,"Capstone C1 — ETEH Retail Sales Analysis (STUDENT BRIEF)",
      "Analyse datasets/clean/eteh_retail.xlsx: profitability by region/category, discount effect, and a 3-month forecast. Deliver a dashboard + 1-page brief.",
      ["Clean the extract (mixed dates/cases).","Pivot profit by Region × Category.","Quantify discount vs profit relationship.",
       "Forecast next quarter with a trendline.","Build a 1-screen dashboard.","Write a 1-page executive brief."],
      ["Cleaned workbook","Dashboard","1-page brief + forecast"],
      [["Analysis depth",35,"Insightful","Shallow"],["Dashboard",25,"Polished","Basic"],["Forecast",20,"Justified","Missing"],["Brief",20,"Clear","Vague"]])
    ws=wb.create_sheet("C2_Commerce")
    proj_sheet(ws,"Capstone C2 — ETEH Commerce SQL Deep-Dive (STUDENT BRIEF)",
      "Use datasets/sql/eteh_commerce.db. Answer: monthly revenue trend, category mix, top-10 sellers, and customer CLV. Export results to CSV.",
      ["Run the 8 queries in datasets/sql/queries.sql and adapt 3.","Join ≥3 tables in at least one query.",
       "Reconcile your totals to eteh_commerce_flat.csv.","Export 4 result tables to CSV.","Present one dashboard-ready extract."],
      ["queries.sql (your edits)","4 CSV result tables","1-page findings"],
      [["SQL skill",40,"Multi-join","Single table"],["Reconcile",30,"Matches","Off"],["Findings",30,"Clear","Weak"]])
    ws=wb.create_sheet("C3_Uni_Health")
    proj_sheet(ws,"Capstone C3 — ETEH University & Health Insight (STUDENT BRIEF)",
      "Combine datasets/clean/eteh_university.xlsx and eteh_health_diabetes.csv. Test whether study hours / BMI relate to outcomes; anonymise; report ethically.",
      ["Profile both datasets; state grain.","Correlation: study hours vs exam score.","Compare diabetes rate by BMI band."],
      ["Anonymised analysis","2 charts","Ethics note + 1-page report"],
      [["Method",35,"Sound","Flawed"],["Privacy",30,"Anonymised","Leaks"],["Report",35,"Ethical","Weak"]])
    wb.save(os.path.join(PST,"Student_Projects.xlsx")); print("wrote Student_Projects.xlsx")

if __name__=="__main__":
    build_section_examples()
    build_assignments()
    build_projects()
    print("DONE workbooks")
