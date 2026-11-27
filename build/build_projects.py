# -*- coding: utf-8 -*-
"""
ETEH Bootcamp - Projects generator.

Folder 3 - Projects_Instructor: ONE complete, solved capstone
  (E-Commerce & Retail Analytics) with a real worked SQL/Excel solution.

Folder 4 - Projects_Students: FOUR student projects, each with a brief,
  a starter data extract, and a marking rubric.

Datasets used: e-commerce (SQL), Superstore, Healthcare, Student Performance.
"""
import os, sqlite3, datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CLEAN = os.path.join(ROOT, "Datasets", "Clean")
SQL   = os.path.join(ROOT, "Datasets", "SQL")
PIN   = os.path.join(ROOT, "Projects_Instructor")
PST   = os.path.join(ROOT, "Projects_Students")
os.makedirs(PIN, exist_ok=True); os.makedirs(PST, exist_ok=True)

C_HEAD="11487A"; C_SUB="1F6FB2"
THIN = Side(style="thin", color="BBBBBB"); BORDER=Border(THIN,THIN,THIN,THIN)
def hdr(ws, row, headers, start=1):
    for j,h in enumerate(headers, start):
        c=ws.cell(row,j,h); c.font=Font(bold=True,color="FFFFFF",size=10)
        c.fill=PatternFill("solid",fgColor=C_HEAD); c.border=BORDER
        c.alignment=Alignment(horizontal="center",wrap_text=True)

def write_df(ws, df, start_row=1, start_col=1):
    hdr(ws, start_row, list(df.columns), start_col)
    for i,(_,r) in enumerate(df.iterrows(), start_row+1):
        for j,col in enumerate(df.columns, start_col):
            v=r[col]
            if pd.isna(v): v=""
            c=ws.cell(i,j,v); c.border=BORDER; c.font=Font(size=9)
            if isinstance(v,(int,float)): c.alignment=Alignment(horizontal="right")
    return start_row+len(df)+1

# ----------------------------------------------------------------- instructor
def build_instructor():
    db = os.path.join(SQL, "ecommerce.db")
    conn = sqlite3.connect(db)
    q = {
        "Revenue by Category": """
            SELECT p.category, ROUND(SUM(t.total_price),2) revenue, COUNT(*) orders
            FROM transactions t JOIN products p ON p.product_id=t.product_id
            GROUP BY p.category ORDER BY revenue DESC;""",
        "Top 10 Sellers": """
            SELECT s.store_name, s.store_state, ROUND(SUM(t.total_price),2) revenue, COUNT(*) txns
            FROM transactions t JOIN sellers s ON s.seller_id=t.seller_id
            GROUP BY s.seller_id ORDER BY revenue DESC LIMIT 10;""",
        "Customer LTV (Top 15)": """
            SELECT c.customer_id, c.first_name||' '||c.last_name customer, c.membership_tier,
                   ROUND(SUM(t.total_price),2) lifetime_spend, COUNT(*) orders
            FROM customers c JOIN transactions t ON t.customer_id=c.customer_id
            GROUP BY c.customer_id ORDER BY lifetime_spend DESC LIMIT 15;""",
        "Monthly Revenue Trend": """
            SELECT substr(order_date,1,7) month, ROUND(SUM(total_price),2) revenue,
                   ROUND(SUM(commission_amount),2) commission
            FROM transactions GROUP BY month ORDER BY month;""",
        "Delivery Performance": """
            SELECT courier_name, ROUND(AVG(julianday(delivered_date)-julianday(shipped_date)),2) avg_days, COUNT(*) parcels
            FROM deliveries WHERE delivered_date IS NOT NULL
            GROUP BY courier_name ORDER BY avg_days;""",
        "Payment Mix": """
            SELECT payment_method, COUNT(*) txns,
                   ROUND(100.0*COUNT(*)/SUM(COUNT(*)) OVER (),1) pct
            FROM transactions GROUP BY payment_method ORDER BY txns DESC;""",
    }
    results = {name: pd.read_sql_query(sql, conn) for name, sql in q.items()}
    conn.close()
    # superstore region profit
    sup = pd.read_csv(os.path.join(CLEAN, "superstore_sales.csv"))
    reg = sup.groupby("region")[["sales","profit"]].sum().round(2).reset_index().sort_values("profit", ascending=False)
    results["Superstore Region Profit"] = reg

    # ---- Brief workbook
    wb = Workbook(); ws = wb.active; ws.title="Brief"
    ws.cell(1,1,"ETEH Ltd - Instructor-Led Capstone Project").font=Font(bold=True,size=14,color=C_HEAD)
    ws.cell(2,1,"E-Commerce & Retail Analytics Deep-Dive").font=Font(bold=True,size=12,color=C_SUB)
    brief = [
      ("Scenario","ETEH's marketplace and a sister retail chain need a single analytics pack: who buys, what sells, which sellers and regions earn the most, and where logistics lag."),
      ("Datasets","ecommerce.db (relational: customers, sellers, products, deliveries, transactions) + Sample - Superstore (superstore_sales.csv)."),
      ("Deliverable 1","A SQL script (instructor_capstone_queries.sql) answering 6 business questions."),
      ("Deliverable 2","An Excel model: raw -> Power Query clean -> Data Model -> 3 dashboards (Revenue, Sellers, Logistics)."),
      ("Deliverable 3","A 2-page PDF report with 3 KPIs, 2 charts and one recommendation."),
      ("Success criteria","Every total ties to source; no #REF!; interactive slicers; report leads with the decision."),
      ("Timeline","Days 1-2 SQL, Day 3 Power Query + Model, Day 4 Dashboards, Day 5 Report & defence."),
      ("Marks","SQL 30 | Model 30 | Dashboard 25 | Report 15 (100 total)."),
    ]
    r=4
    for k,v in brief:
        ws.cell(r,1,k).font=Font(bold=True,size=10); ws.cell(r,2,v).font=Font(size=10)
        ws.cell(r,2).alignment=Alignment(wrap_text=True); r+=1
    ws.column_dimensions["A"].width=18; ws.column_dimensions["B"].width=95
    for name, df in results.items():
        ws2 = wb.create_sheet(name[:31])
        write_df(ws2, df)
        for j in range(1, df.shape[1]+1):
            ws2.column_dimensions[get_column_letter(j)].width=22
    wb.save(os.path.join(PIN,"Instructor_Capstone_Brief.xlsx"))

    # ---- Solution workbook (queries as text + results)
    swb = Workbook(); sws = swb.active; sws.title="Solution Notes"
    sws.cell(1,1,"Worked Solution - E-Commerce & Retail Analytics").font=Font(bold=True,size=13,color=C_HEAD)
    sws.cell(2,1,"Each question below is answered in its own sheet with the exact result from ecommerce.db / superstore.").font=Font(italic=True,size=9)
    r=4
    for name, sql in q.items():
        sws.cell(r,1,f"Q: {name}").font=Font(bold=True,color=C_SUB); r+=1
        for line in sql.strip().splitlines():
            sws.cell(r,1,line).font=Font(size=9, name="Consolas"); r+=1
        r+=1
    sws.column_dimensions["A"].width=110
    for name, df in results.items():
        ws2 = swb.create_sheet("R - "+name[:26])
        write_df(ws2, df)
        for j in range(1, df.shape[1]+1):
            ws2.column_dimensions[get_column_letter(j)].width=22
    swb.save(os.path.join(PIN,"Instructor_Capstone_Solution.xlsx"))

    # ---- SQL file
    with open(os.path.join(PIN,"instructor_capstone_queries.sql"),"w") as f:
        f.write("-- ETEH Instructor Capstone - E-Commerce & Retail Analytics\n")
        f.write(f"-- Generated {datetime.date.today().isoformat()}\n\n")
        for name, sql in q.items():
            f.write(f"-- {name}\n{sql.strip()}\n\n")

    # ---- README
    with open(os.path.join(PIN,"README.md"),"w") as f:
        f.write("# Projects - Instructor (Capstone)\n\n")
        f.write("**Project:** E-Commerce & Retail Analytics Deep-Dive\n\n")
        f.write("This is the single instructor-led capstone. It is fully solved.\n\n")
        f.write("## Files\n")
        f.write("- `Instructor_Capstone_Brief.xlsx` - scenario, datasets, deliverables, marking.\n")
        f.write("- `Instructor_Capstone_Solution.xlsx` - worked SQL queries + real result tables.\n")
        f.write("- `instructor_capstone_queries.sql` - the runnable SQL solution.\n\n")
        f.write("## How to use\n1. Teach the SQL in Week 3 against `Datasets/SQL/ecommerce.db`.\n")
        f.write("2. Demonstrate the Power Query + Data Model build in Week 2-3.\n")
        f.write("3. Hand the brief to strong students as a stretch; use the solution for live marking.\n")
    print("[Instructor] capstone built:", os.listdir(PIN))

# ----------------------------------------------------------------- students
STUDENT_PROJECTS = [
  {"id":"01","name":"Retail Analytics with PivotTables",
   "folder":"project_01_retail_analytics",
   "dataset":"superstore_sales.csv", "subset_rows":800, "seed":11,
   "brief":("Build a retail performance pack for the Superstore. Use PivotTables to report Sales and Profit by Region and Category, identify the worst sub-category, and recommend one action. Deliver a workbook + 1-page PDF."),
   "tasks":["Clean the extract (fix the dirty ship dates if used).","Pivot Sales & Profit by Region x Category.","Rank sub-categories by profit; flag the loss-maker.","Write one data-backed recommendation."],
   "rubric":[("Data cleaning",25,"No errors, types correct"),("PivotTables",30,"Correct design, 2 pivots"),("Insight",30,"Right worst sub-category identified"),("Communication",15,"Clear 1-page write-up")]},
  {"id":"02","name":"Hospital Operations Review",
   "folder":"project_02_healthcare_operations",
   "dataset":"healthcare_records.csv", "subset_rows":1500, "seed":22,
   "brief":("Act as a hospital COO. From the admissions extract, report average Length of Stay and Billing by Medical Condition, and the busiest admission type. Deliver a workbook + chart."),
   "tasks":["Compute length_of_stay from admission/discharge dates.","Summarise billing & LOS by condition.","Count admissions by admission type.","Recommend where to add capacity."],
   "rubric":[("Calculations",30,"LOS & billing correct"),("Analysis",30,"Condition comparison accurate"),("Visual",20,"At least one clear chart"),("Recommendation",20,"Specific, data-backed")]},
  {"id":"03","name":"Marketplace Seller Performance (SQL)",
   "folder":"project_03_marketplace_sql",
   "dataset":None,
   "brief":("Using Datasets/SQL/ecommerce.db, write SQL to rank sellers by revenue, list the top 10 customers by lifetime spend, and report monthly revenue. Deliver a .sql file + result export."),
   "tasks":["Rank sellers by total revenue (JOIN sellers + transactions).","Top 10 customers by lifetime spend (JOIN customers + transactions).","Monthly revenue trend (GROUP BY month).","Export results to CSV."],
   "rubric":[("SQL correctness",40,"Joins & aggregates correct"),("Results",30,"Numbers tie to source"),("Export",15,"Clean CSV delivered"),("Readability",15,"Commented, formatted")]},
  {"id":"04","name":"Student Success Predictor",
   "folder":"project_04_student_success",
   "dataset":"student_performance.csv", "subset_rows":1200, "seed":33,
   "brief":("Help a school improve exam results. From the student factors dataset, find which factor correlates most with Exam Score, build a simple linear regression of Score on Study Hours, and state one intervention."),
   "tasks":["Correlate Exam Score with each factor.","Identify the strongest driver.","Fit Score ~ Hours Studied (regression).","Recommend one intervention with expected impact."],
   "rubric":[("Correlation",30,"Correct, honest interpretation"),("Regression",30,"Slope & R-squared reported"),("Insight",25,"Strongest driver named"),("Communication",15,"Plain-language conclusion")]},
]

def build_students():
    idx_rows=[]
    for p in STUDENT_PROJECTS:
        d = os.path.join(PST, p["folder"]); os.makedirs(d, exist_ok=True)
        # brief
        wb=Workbook(); ws=wb.active; ws.title="Brief"
        ws.cell(1,1,f"Student Project {p['id']}: {p['name']}").font=Font(bold=True,size=13,color=C_HEAD)
        ws.cell(3,1,"Brief").font=Font(bold=True,color=C_SUB); ws.cell(3,2,p["brief"]).font=Font(size=10)
        ws.cell(3,2).alignment=Alignment(wrap_text=True)
        ws.cell(4,1,"Tasks").font=Font(bold=True,color=C_SUB)
        r=5
        for i,t in enumerate(p["tasks"],1):
            ws.cell(r,1,f"{i}."); ws.cell(r,2,t).font=Font(size=10); ws.cell(r,2).alignment=Alignment(wrap_text=True); r+=1
        ws.cell(r+1,1,"Datasets").font=Font(bold=True,color=C_SUB)
        ws.cell(r+1,2, p["dataset"] if p["dataset"] else "Datasets/SQL/ecommerce.db (provided)").font=Font(size=10)
        ws.column_dimensions["A"].width=14; ws.column_dimensions["B"].width=95
        wb.save(os.path.join(d,"Brief.xlsx"))
        # rubric
        wb2=Workbook(); ws2=wb2.active; ws2.title="Rubric"
        ws2.cell(1,1,f"Marking Rubric - Project {p['id']}").font=Font(bold=True,size=12,color=C_HEAD)
        hdr(ws2,3,["Criterion","Weight","Distinction descriptor"])
        rr=4
        total=0
        for crit,w,desc in p["rubric"]:
            ws2.cell(rr,1,crit); ws2.cell(rr,2,w); ws2.cell(rr,3,desc)
            for c in range(1,4): ws2.cell(rr,c).border=BORDER; ws2.cell(rr,c).font=Font(size=10); ws2.cell(rr,c).alignment=Alignment(wrap_text=True)
            total+=w; rr+=1
        ws2.cell(rr,1,"TOTAL").font=Font(bold=True); ws2.cell(rr,2,total).font=Font(bold=True)
        ws2.column_dimensions["A"].width=20; ws2.column_dimensions["B"].width=10; ws2.column_dimensions["C"].width=70
        wb2.save(os.path.join(d,"Rubric.xlsx"))
        # starter data
        if p["dataset"]:
            df = pd.read_csv(os.path.join(CLEAN, p["dataset"]))
            sub = df.sample(n=min(p["subset_rows"], len(df)), random_state=p["seed"]).reset_index(drop=True)
            sub.to_csv(os.path.join(d,"starter_data.csv"), index=False)
            data_note = f"starter_data.csv ({len(sub)} rows sampled from {p['dataset']})"
        else:
            # copy schema note for SQL project
            with open(os.path.join(d,"starter_data.txt"),"w") as f:
                f.write("Use the relational database at: ../../Datasets/SQL/ecommerce.db\n")
                f.write("Tables: customers, sellers, products, deliveries, transactions (all joined by *_id keys).\n")
                f.write("See Datasets/SQL/ecommerce_schema.sql for the full schema.\n")
            data_note = "starter_data.txt (points to ecommerce.db)"
        idx_rows.append([p["id"], p["name"], p["dataset"] or "ecommerce.db (SQL)", data_note, os.path.join(p["folder"],"Brief.xlsx")])
        print(f"[Student] {p['folder']}: brief + rubric + {data_note}")
    # overview
    ov=Workbook(); o=ov.active; o.title="Overview"
    o.cell(1,1,"ETEH Bootcamp - Student Projects Index").font=Font(bold=True,size=14,color=C_HEAD)
    hdr(o,3,["ID","Project","Source dataset","Starter file","Brief"])
    for i,row in enumerate(idx_rows,4):
        for j,v in enumerate(row,1):
            c=o.cell(i,j,v); c.border=BORDER; c.font=Font(size=10); c.alignment=Alignment(wrap_text=True)
    for col,w in zip("ABCDE",[6,34,22,40,40]): o.column_dimensions[col].width=w
    ov.save(os.path.join(PST,"Student_Projects_Overview.xlsx"))
    # README
    with open(os.path.join(PST,"README.md"),"w") as f:
        f.write("# Projects - Students\n\nFour independent projects, one per dataset theme.\n\n")
        for p in STUDENT_PROJECTS:
            f.write(f"## {p['id']} - {p['name']}\n")
            f.write(f"- Brief: `{p['folder']}/Brief.xlsx`\n- Rubric: `{p['folder']}/Rubric.xlsx`\n")
            if p["dataset"]:
                f.write(f"- Starter data: `{p['folder']}/starter_data.csv`\n")
            else:
                f.write(f"- Starter: `{p['folder']}/starter_data.txt` (uses ecommerce.db)\n")
            f.write("\n")
        f.write("Each project is marked out of 100 using its rubric. Pair students or run individually.\n")
    print("[Students] overview + README written.")

if __name__ == "__main__":
    build_instructor()
    build_students()
