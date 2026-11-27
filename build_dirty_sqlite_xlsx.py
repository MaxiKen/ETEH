#!/usr/bin/env python3
"""Build DIRTY teaching files, SQLite DBs, XLSX extracts."""
import pandas as pd, numpy as np, sqlite3, re
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT = Path("/home/user/ETEH")
CLEAN = ROOT/"ETEH_Bootcamp/Datasets/CLEAN"
DIRTY = ROOT/"ETEH_Bootcamp/Datasets/DIRTY_FOR_WEEK2"
SQLITE_DIR = ROOT/"ETEH_Bootcamp/Datasets/SQLITE"
XLSX_DIR = ROOT/"ETEH_Bootcamp/Datasets/XLSX_EXTRACTS"
rng = np.random.default_rng(7)

NAVY="1F3864"; TEAL="0E6B7A"; LIGHT="D9E2F3"; GOLD="FFF2CC"; GREY="F2F2F2"
hdr_font=Font(name="Calibri",size=10,bold=True,color="FFFFFF")
hdr_fill=PatternFill("solid",fgColor=NAVY)
title_font=Font(name="Calibri",size=14,bold=True,color=NAVY)
sub_font=Font(name="Calibri",size=10,italic=True,color="595959")
thin=Side(style="thin",color="B4C6E7")
border=Border(left=thin,right=thin,top=thin,bottom=thin)

def style_header(ws, row=1, ncols=None):
    ncols = ncols or ws.max_column
    for c in range(1, ncols+1):
        cell=ws.cell(row=row,column=c)
        cell.font=hdr_font; cell.fill=hdr_fill
        cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        cell.border=border
    ws.row_dimensions[row].height=30
    ws.freeze_panes=ws.cell(row=row+1,column=1).coordinate
    ws.auto_filter.ref=ws.dimensions

def autofit(ws, maxw=32):
    for col in ws.columns:
        letter=get_column_letter(col[0].column)
        mx=max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[letter].width=min(max(10,mx+2),maxw)

# ============================================================
# DIRTY FILES — intentionally inject 12 quality problems
# ============================================================
print("=== DIRTY FILES ===")

# --- Dirty 1: Salary (from RAW, re-inject + add new mess) ---
sal_raw = pd.read_csv(ROOT/"Salary_Data.csv", low_memory=False, encoding="utf-8-sig")
sal_dirty = sal_raw.copy()
# take 1,500-row extract for manageable lab (stratified-ish: first 1500 after shuffle)
sal_dirty = sal_dirty.sample(n=1500, random_state=11).reset_index(drop=True)
sal_dirty["Salary"] = sal_dirty["Salary"].astype(object)
sal_dirty["Age"] = sal_dirty["Age"].astype(object)
sal_dirty["Education Level"] = sal_dirty["Education Level"].astype(object)
sal_dirty["Gender"] = sal_dirty["Gender"].astype(object)
sal_dirty["Job Title"] = sal_dirty["Job Title"].astype(object)
# inject: inconsistent education labels
sal_dirty.loc[sal_dirty.index[::7], "Education Level"] = sal_dirty.loc[sal_dirty.index[::7], "Education Level"].replace({"Bachelor's Degree":"BACHELORS","Master's Degree":"masters ","PhD":"phd"})
# inject: numbers as text with currency/commas
idx = sal_dirty.index[::11]
sal_dirty.loc[idx, "Salary"] = sal_dirty.loc[idx, "Salary"].apply(lambda x: f"₦{int(x):,}" if pd.notna(x) else x)
# inject: extra spaces + case variants in Gender/Job
sal_dirty.loc[sal_dirty.index[::13], "Gender"] = sal_dirty.loc[sal_dirty.index[::13], "Gender"].replace({"Male":" MALE ","Female":"female"})
sal_dirty.loc[sal_dirty.index[::17], "Job Title"] = sal_dirty.loc[sal_dirty.index[::17], "Job Title"].apply(lambda x: f"  {x} " if pd.notna(x) else x)
# inject: blanks (missing)
sal_dirty.loc[sal_dirty.index[::19], "Salary"] = np.nan
sal_dirty.loc[sal_dirty.index[::23], "Age"] = np.nan
# inject: duplicates (append 80 exact copies)
dups = sal_dirty.sample(n=80, random_state=12)
sal_dirty = pd.concat([sal_dirty, dups], ignore_index=True).sample(frac=1, random_state=13).reset_index(drop=True)
# inject: impossible values
sal_dirty.loc[sal_dirty.index[5], "Salary"] = -5000
sal_dirty.loc[sal_dirty.index[9], "Age"] = 150
sal_dirty.loc[sal_dirty.index[14], "Years of Experience"] = 99
sal_dirty.to_csv(DIRTY/"salary_DIRTY.csv", index=False, encoding="utf-8")
print(f" salary_DIRTY {sal_dirty.shape}")

# --- Dirty 2: Superstore extract (dates mixed, numbers as text, spacer rows) ---
ss = pd.read_csv(CLEAN/"superstore_CLEAN.csv", low_memory=False)
ss_d = ss.sample(n=1200, random_state=21).copy().reset_index(drop=True)
ss_d["Sales"] = ss_d["Sales"].astype(object)
ss_d["Profit"] = ss_d["Profit"].astype(object)
ss_d["Category"] = ss_d["Category"].astype(object)
ss_d["Region"] = ss_d["Region"].astype(object)
# mixed date formats in Order_Date_ISO -> rewrite some as M/D/YYYY and D-Mon-YY
def mix_date(s):
    try:
        d = pd.to_datetime(s)
        r = rng.random()
        if r<0.4: return d.strftime("%m/%d/%Y")
        elif r<0.7: return d.strftime("%d-%b-%y")
        else: return s
    except: return s
ss_d["Order_Date_ISO"] = ss_d["Order_Date_ISO"].apply(mix_date)
# numbers as text with $ and commas
ss_d["Sales"] = ss_d["Sales"].apply(lambda x: f"${x:,.2f}" if rng.random()<0.3 else x)
# inconsistent category case
ss_d.loc[ss_d.index[::9], "Category"] = ss_d.loc[ss_d.index[::9], "Category"].str.upper()
ss_d.loc[ss_d.index[::15], "Region"] = ss_d.loc[ss_d.index[::15], "Region"].apply(lambda x: f" {x} ")
# blanks
ss_d.loc[ss_d.index[::25], "Profit"] = np.nan
# spacer rows: insert blank rows every ~100
ss_d = ss_d.iloc[:100].to_dict("records") + [{}] + ss_d.iloc[100:].to_dict("records")
ss_d = pd.DataFrame(ss_d)
ss_d.to_csv(DIRTY/"superstore_DIRTY.csv", index=False, encoding="utf-8")
print(f" superstore_DIRTY {ss_d.shape}")

# --- Dirty 3: E-commerce Transactions extract (key padding + blanks + mixed status) ---
tr = pd.read_csv(CLEAN/"ecommerce_transactions_CLEAN.csv", low_memory=False)
tr_d = tr.sample(n=2000, random_state=31).copy().reset_index(drop=True)
tr_d["total_price"] = tr_d["total_price"].astype(object)
tr_d["quantity"] = tr_d["quantity"].astype(object)
tr_d["order_status"] = tr_d["order_status"].astype(object)
tr_d["payment_status"] = tr_d["payment_status"].astype(object)
# break delivery_id padding (remove leading zeros) for 50%
tr_d["delivery_id"] = tr_d["delivery_id"].apply(lambda x: "DEL"+str(int(re.search(r"(\d+)",str(x)).group(1))) if rng.random()<0.5 else x)
# mixed status case
tr_d.loc[tr_d.index[::8], "order_status"] = tr_d.loc[tr_d.index[::8], "order_status"].str.upper()
tr_d.loc[tr_d.index[::12], "payment_status"] = tr_d.loc[tr_d.index[::12], "payment_status"].apply(lambda x: f" {x} ")
# numbers as text
tr_d["total_price"] = tr_d["total_price"].apply(lambda x: f"₦{x:,.2f}" if rng.random()<0.25 else x)
# blanks + duplicates
tr_d.loc[tr_d.index[::20], "quantity"] = np.nan
tr_d = pd.concat([tr_d, tr_d.sample(n=60, random_state=32)], ignore_index=True).sample(frac=1, random_state=33).reset_index(drop=True)
# append 40 fully blank rows (like original file's trap)
blank = pd.DataFrame([{}]*40, columns=tr_d.columns)
tr_d = pd.concat([tr_d, blank], ignore_index=True)
tr_d.to_csv(DIRTY/"ecommerce_transactions_DIRTY.csv", index=False, encoding="utf-8")
print(f" ecommerce_transactions_DIRTY {tr_d.shape}")

# --- Dirty 4: Student Performance extract (ordinal mess + missing + typos) ---
spf = pd.read_csv(CLEAN/"student_performance_CLEAN.csv", low_memory=False)
spf_d = spf.sample(n=1500, random_state=41).copy().reset_index(drop=True)
spf_d["Hours_Studied"] = spf_d["Hours_Studied"].astype(object)
spf_d["Exam_Score"] = spf_d["Exam_Score"].astype(object)
spf_d["Parental_Involvement"] = spf_d["Parental_Involvement"].astype(object)
spf_d["Teacher_Quality"] = spf_d["Teacher_Quality"].astype(object)
spf_d.loc[spf_d.index[::6], "Parental_Involvement"] = spf_d.loc[spf_d.index[::6], "Parental_Involvement"].replace({"Low":"low","Medium":"MEDIUM","High":" High "})
spf_d.loc[spf_d.index[::10], "Teacher_Quality"] = np.nan
spf_d.loc[spf_d.index[::14], "Exam_Score"] = spf_d.loc[spf_d.index[::14], "Exam_Score"] + 50  # push some >100
spf_d["Hours_Studied"] = spf_d["Hours_Studied"].apply(lambda x: f"{x} hrs" if rng.random()<0.2 else x)
spf_d.to_csv(DIRTY/"student_performance_DIRTY.csv", index=False, encoding="utf-8")
print(f" student_performance_DIRTY {spf_d.shape}")

# --- Dirty 5: Healthcare extract (case mess + negative billing + date formats) ---
hc = pd.read_csv(CLEAN/"healthcare_CLEAN.csv", low_memory=False)
hc_d = hc.sample(n=1500, random_state=51).copy().reset_index(drop=True)
hc_d["Billing Amount"] = hc_d["Billing Amount"].astype(object)
hc_d["Blood_Type"] = hc_d["Blood_Type"].astype(object)
hc_d["Name"] = hc_d["Name"].astype(object)
hc_d["Date_of_Admission_ISO"] = hc_d["Date_of_Admission_ISO"].astype(object)
hc_d["Name"] = hc_d["Name"].apply(lambda x: x.upper() if rng.random()<0.3 else x.lower() if rng.random()<0.3 else x)
hc_d.loc[hc_d.index[::18], "Billing Amount"] = hc_d.loc[hc_d.index[::18], "Billing Amount"] * -1
hc_d["Date_of_Admission_ISO"] = hc_d["Date_of_Admission_ISO"].apply(lambda s: pd.to_datetime(s).strftime("%m/%d/%Y") if rng.random()<0.4 else s)
hc_d.loc[hc_d.index[::22], "Blood_Type"] = np.nan
hc_d.to_csv(DIRTY/"healthcare_DIRTY.csv", index=False, encoding="utf-8")
print(f" healthcare_DIRTY {hc_d.shape}")

# Issues log
issues = pd.DataFrame([
    ["salary_DIRTY.csv","Duplicates (80 exact copies appended)","Remove duplicates; verify with COUNT vs COUNTA"],
    ["salary_DIRTY.csv","Inconsistent categories (BACHELORS/masters/phd, MALE/female)","Standardise via lookup table; TRIM/PROPER"],
    ["salary_DIRTY.csv","Numbers stored as text (₦1,200)","VALUE/SUBSTITUTE to strip currency; set Number format"],
    ["salary_DIRTY.csv","Extra spaces in Job Title","TRIM; Text to Columns if needed"],
    ["salary_DIRTY.csv","Missing Salary/Age (blanks)","Decide: blank vs 0 vs Unknown; document; do not average blanks as 0"],
    ["salary_DIRTY.csv","Impossible values (-5000 salary, Age 150, Exp 99)","Flag with IF; quarantine; do not silently delete"],
    ["superstore_DIRTY.csv","Mixed date formats (ISO vs M/D/YYYY vs D-Mon-YY)","Parse to ISO YYYY-MM-DD; DATEVALUE with locale check"],
    ["superstore_DIRTY.csv","Sales as text ($1,234.56)","Strip $ and commas; VALUE; check SUM reconciles"],
    ["superstore_DIRTY.csv","Category case (FURNITURE) + Region spaces","Standardise case; TRIM"],
    ["superstore_DIRTY.csv","Blank Profit + spacer blank row","Filter blanks; delete spacer rows; check row counts"],
    ["ecommerce_transactions_DIRTY.csv","delivery_id padding broken (DEL1 vs DEL000001)","Standardise key: 'DEL' & TEXT(VALUE(...),'000000')"],
    ["ecommerce_transactions_DIRTY.csv","Status case + spaces (DELIVERED, ' Paid ')","TRIM/UPPER standardise; data validation list"],
    ["ecommerce_transactions_DIRTY.csv","total_price as text (₦...) + blank quantity + 60 dups + 40 blank rows","Clean numbers; handle blanks; dedupe; remove blank rows"],
    ["student_performance_DIRTY.csv","Ordinal case mess (low/MEDIUM/' High ')","Map to 1/2/3 via lookup; standardise"],
    ["student_performance_DIRTY.csv","Missing Teacher_Quality + Exam_Score >100 + '23 hrs' text","Explicit Unknown; flag >100; VALUE after SUBSTITUTE"],
    ["healthcare_DIRTY.csv","Name case (UPPER/lower) + negative billing + mixed dates + blank Blood_Type","PROPER; flag negatives; ISO dates; explicit missing"],
], columns=["File","Problem injected","Fix (Week 2)"])
issues.to_csv(DIRTY/"DIRTY_ISSUES_LOG.csv", index=False, encoding="utf-8")
print(f" issues log {issues.shape}")

# ============================================================
# SQLITE — e-commerce flagship + superstore + education + health
# ============================================================
print("\n=== SQLITE ===")
def build_db(db_path, tables: dict, indexes: list):
    if db_path.exists(): db_path.unlink()
    con = sqlite3.connect(db_path)
    cur = con.cursor()
    for name, df in tables.items():
        # sanitise col names for sqlite (keep readable)
        df2 = df.copy()
        df2.columns = [re.sub(r"[^A-Za-z0-9_]+","_", str(c)).strip("_") for c in df2.columns]
        df2.to_sql(name, con, index=False, if_exists="replace")
        print(f"  {db_path.name} :: {name} {df2.shape}")
    for stmt in indexes:
        try: cur.execute(stmt)
        except Exception as e: print(f"  index warn: {e} :: {stmt[:80]}")
    con.commit()
    # integrity
    cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
    print(f"  tables: {[r[0] for r in cur.fetchall()]}")
    con.close()

tr_c = pd.read_csv(CLEAN/"ecommerce_transactions_CLEAN.csv")
se_c = pd.read_csv(CLEAN/"ecommerce_sellers_CLEAN.csv")
pr_c = pd.read_csv(CLEAN/"ecommerce_products_CLEAN.csv")
de_c = pd.read_csv(CLEAN/"ecommerce_deliveries_CLEAN.csv")
cu_c = pd.read_csv(CLEAN/"ecommerce_customers_CLEAN.csv")
build_db(SQLITE_DIR/"eteh_ecommerce.sqlite3",
         {"customers":cu_c,"sellers":se_c,"products":pr_c,"deliveries":de_c,"transactions":tr_c},
         ["CREATE INDEX IF NOT EXISTS idx_trx_customer ON transactions(customer_id)",
          "CREATE INDEX IF NOT EXISTS idx_trx_product ON transactions(product_id)",
          "CREATE INDEX IF NOT EXISTS idx_trx_seller ON transactions(seller_id)",
          "CREATE INDEX IF NOT EXISTS idx_trx_order ON transactions(order_id)",
          "CREATE INDEX IF NOT EXISTS idx_trx_date ON transactions(order_date)",
          "CREATE INDEX IF NOT EXISTS idx_del_order ON deliveries(order_id)",
          "CREATE INDEX IF NOT EXISTS idx_prod_seller ON products(seller_id)"])

ss_c = pd.read_csv(CLEAN/"superstore_CLEAN.csv")
build_db(SQLITE_DIR/"eteh_superstore.sqlite3", {"superstore":ss_c},
         ["CREATE INDEX IF NOT EXISTS idx_ss_date ON superstore(Order_Date_ISO)",
          "CREATE INDEX IF NOT EXISTS idx_ss_cat ON superstore(Category)",
          "CREATE INDEX IF NOT EXISTS idx_ss_region ON superstore(Region)"])

spf_c = pd.read_csv(CLEAN/"student_performance_CLEAN.csv")
build_db(SQLITE_DIR/"eteh_education.sqlite3", {"student_performance":spf_c},
         ["CREATE INDEX IF NOT EXISTS idx_spf_exam ON student_performance(Exam_Score)"])

hc_c = pd.read_csv(CLEAN/"healthcare_ANONYMISED.csv")
build_db(SQLITE_DIR/"eteh_healthcare.sqlite3", {"admissions":hc_c},
         ["CREATE INDEX IF NOT EXISTS idx_hc_date ON admissions(Date_of_Admission_ISO)",
          "CREATE INDEX IF NOT EXISTS idx_hc_cond ON admissions(Medical_Condition)"])

# SQL schema + queries file
sql_text = """-- ============================================================
-- ETEH Bootcamp — E-commerce SQLite schema & teaching queries
-- Database: eteh_ecommerce.sqlite3  (Week 3, Module 3.2 / 3.4)
-- Run in DB Browser for SQLite or sqlite3 CLI.
-- ============================================================

-- Tables (grain declared):
--   customers    1 row per customer_id (PK)
--   sellers      1 row per seller_id (PK)
--   products     1 row per product_id (PK, FK seller_id -> sellers)
--   deliveries   1 row per delivery_id (PK, FK order_id, customer_id)
--   transactions 1 row per transaction_id (PK, FKs to all dims)
-- NOTE: 570 Cancelled transactions have NO delivery row (cancelled before
-- shipment). Use LEFT JOIN from transactions to deliveries to keep them.

-- ---------- Q1: Revenue by category (INNER JOIN) ----------
SELECT p.category, COUNT(*) AS lines, SUM(t.total_price) AS revenue
FROM transactions t
JOIN products p ON t.product_id = p.product_id
GROUP BY p.category ORDER BY revenue DESC;

-- ---------- Q2: Revenue by state (JOIN via deliveries) ----------
SELECT d.state, COUNT(*) AS deliveries, SUM(t.total_price) AS revenue
FROM transactions t
JOIN deliveries d ON t.order_id = d.order_id
GROUP BY d.state ORDER BY revenue DESC;

-- ---------- Q3: Keep cancelled orders (LEFT JOIN) ----------
SELECT t.order_status, COUNT(*) AS n, SUM(t.total_price) AS value
FROM transactions t
LEFT JOIN deliveries d ON t.order_id = d.order_id
GROUP BY t.order_status;

-- ---------- Q4: Top sellers by commission ----------
SELECT s.store_name, s.store_city, COUNT(*) AS lines,
       SUM(t.commission_amount) AS commission
FROM transactions t
JOIN sellers s ON t.seller_id = s.seller_id
GROUP BY s.seller_id ORDER BY commission DESC LIMIT 10;

-- ---------- Q5: Monthly revenue trend (for forecasting) ----------
SELECT order_month, SUM(total_price) AS revenue, COUNT(*) AS lines
FROM transactions GROUP BY order_month ORDER BY order_month;

-- ---------- Q6: Orphan check (should return 0 for customers/products) ----------
SELECT COUNT(*) AS orphan_customer_lines
FROM transactions t LEFT JOIN customers c ON t.customer_id = c.customer_id
WHERE c.customer_id IS NULL;

-- ---------- Q7: Payment mix ----------
SELECT payment_method, payment_status, COUNT(*) AS n, SUM(total_price) AS value
FROM transactions GROUP BY payment_method, payment_status ORDER BY value DESC;
"""
(SQLITE_DIR/"ecommerce_queries.sql").write_text(sql_text, encoding="utf-8")
print(" wrote ecommerce_queries.sql")

# ============================================================
# XLSX EXTRACTS — multi-sheet workbooks per dataset
# ============================================================
print("\n=== XLSX EXTRACTS ===")
def write_xlsx(path, sheets: dict, title: str, subtitle: str):
    wb = Workbook()
    # README first
    ws0 = wb.active; ws0.title="00_README"; ws0.sheet_properties.tabColor=NAVY
    ws0["A1"]=title; ws0["A1"].font=title_font
    ws0["A2"]=subtitle; ws0["A2"].font=sub_font
    ws0["A4"]="How to use: each sheet below is a tidy table (one header row, one record per row). Do not insert blank rows/columns. Filter/sort freely. Formulas in teaching workbooks reference these sheets."
    ws0["A4"].alignment=Alignment(wrap_text=True,vertical="top")
    ws0.row_dimensions[4].height=45
    r=6
    for name, df in sheets.items():
        ws0.cell(row=r,column=1,value=f"{name}  —  {df.shape[0]:,} rows × {df.shape[1]} cols").font=Font(name="Calibri",size=10,bold=True)
        r+=1
    ws0.column_dimensions["A"].width=110
    ws0.sheet_properties.pageSetUpPr.fitToPage=True
    ws0.page_setup.orientation="portrait"; ws0.page_setup.fitToWidth=1; ws0.page_setup.fitToHeight=1
    ws0.print_title_rows="1:2"
    for name, df in sheets.items():
        ws = wb.create_sheet(title=name[:31])
        ws.sheet_properties.tabColor=TEAL
        # header
        for j, col in enumerate(df.columns, 1):
            c=ws.cell(row=1,column=j,value=str(col))
            c.font=hdr_font; c.fill=hdr_fill; c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); c.border=border
        ws.row_dimensions[1].height=32
        # data (cap at 5000 rows per sheet for file size; note if truncated)
        n = min(len(df), 5000)
        for i in range(n):
            row = df.iloc[i]
            for j, v in enumerate(row, 1):
                if pd.isna(v): val=None
                elif isinstance(v,(np.integer,)): val=int(v)
                elif isinstance(v,(np.floating,)): val=float(v)
                else: val=v
                c=ws.cell(row=i+2,column=j,value=val)
                c.font=Font(name="Calibri",size=9); c.border=border
                if i%2==1: c.fill=PatternFill("solid",fgColor=GREY)
        if len(df)>5000:
            ws.cell(row=n+2,column=1,value=f"NOTE: showing first 5,000 of {len(df):,} rows. Full data in CSV.").font=Font(name="Calibri",size=9,italic=True,color="C00000")
        ws.freeze_panes="A2"
        ws.auto_filter.ref=f"A1:{get_column_letter(len(df.columns))}{n+1}"
        ws.sheet_properties.pageSetUpPr.fitToPage=True
        ws.page_setup.orientation="landscape"; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0
        ws.page_setup.paperSize=ws.PAPERSIZE_A4
        ws.print_title_rows="1:1"
        autofit(ws)
    wb.save(path)
    print(f"  wrote {path.name}  sheets={list(sheets.keys())}  size={path.stat().st_size/1024:.0f} KB")

# E-commerce workbook (samples: full dims + 3,000-row TRX sample + monthly summary)
tr_sample = tr_c.sample(n=3000, random_state=99).sort_values("order_date").reset_index(drop=True)
monthly = tr_c.groupby("order_month").agg(lines=("transaction_id","count"), revenue=("total_price","sum"), commission=("commission_amount","sum")).reset_index()
write_xlsx(XLSX_DIR/"ecommerce_EXTRACT.xlsx",
           {"Transactions_3000":tr_sample,"Customers":cu_c,"Products":pr_c,"Sellers":se_c,"Deliveries_3000":de_c.sample(n=3000,random_state=99).reset_index(drop=True),"Monthly_Summary":monthly},
           "E-commerce Nigeria — Teaching Extract",
           "Source: e-commerce.xlsx (5 sheets). Transactions cleaned: 11,000 real rows (26,000 blank rows removed). delivery_id standardised to DEL+6 digits. 570 Cancelled orders have no delivery (expected).")
# Superstore workbook
write_xlsx(XLSX_DIR/"superstore_EXTRACT.xlsx",
           {"Superstore_5000":ss_c.head(5000).reset_index(drop=True),
            "Monthly_Sales": ss_c.groupby("Order_Month").agg(lines=("Row ID","count"),sales=("Sales","sum"),profit=("Profit","sum")).reset_index(),
            "Category_Summary": ss_c.groupby("Category").agg(lines=("Row ID","count"),sales=("Sales","sum"),profit=("Profit","sum")).reset_index()},
           "Superstore (US Retail) — Teaching Extract",
           "Source: Sample - Superstore.csv (9,994 lines, 2014–2017). Dates converted to ISO. Margin = Profit/Sales. Postal code stored as 5-char text.")
# Education workbook
write_xlsx(XLSX_DIR/"student_performance_EXTRACT.xlsx",
           {"Students_5000":spf_c.head(5000).reset_index(drop=True),
            "Exam_Bands": spf_c.groupby("Exam_Band").size().reset_index(name="students")},
           "Student Performance Factors — Teaching Extract",
           "Source: StudentPerformanceFactors.csv (6,607 students). Missing Teacher_Quality / Parental_Education / Distance coded as 'Unknown'. Exam_Score >100 flagged.")
# Health workbook (anonymised)
write_xlsx(XLSX_DIR/"healthcare_EXTRACT_ANON.xlsx",
           {"Admissions_5000":hc_c.head(5000).reset_index(drop=True),
            "Condition_Summary": hc_c.groupby("Medical_Condition").agg(admissions=("Patient_ID","count"), avg_billing=("Billing_Amount_Abs","mean"), avg_los=("Length_of_Stay_Days","mean")).reset_index()},
           "Healthcare Admissions (ANONYMISED) — Teaching Extract",
           "Source: healthcare_dataset.csv (55,500 admissions). Names replaced with Patient_ID. Negative billing flagged (108 rows). LOS = Discharge − Admission.")
# Salary workbook
sal_c = pd.read_csv(CLEAN/"salary_CLEAN.csv")
write_xlsx(XLSX_DIR/"salary_EXTRACT.xlsx",
           {"Salary_Clean":sal_c,
            "Job_Family_Summary": sal_c.groupby("Job_Family").agg(n=("Salary","count"), median_salary=("Salary","median"), mean_salary=("Salary","mean")).reset_index()},
           "Salary Data (Clean) — Teaching Extract",
           "Source: Salary_Data.csv (6,704 rows → 1,791 after exact-dedupe). Education standardised; Job_Family added. Outliers flagged, not deleted.")
# Hotel + Diabetes sample + Social + Depression-anon extracts (single-sheet + summary)
ht_c = pd.read_csv(CLEAN/"hotel_reservations_CLEAN.csv")
write_xlsx(XLSX_DIR/"hotel_EXTRACT.xlsx",
           {"Bookings_5000":ht_c.head(5000).reset_index(drop=True),
            "Cancellation_Summary": ht_c.groupby("market_segment_type").agg(bookings=("Booking_ID","count"), cancellations=("Is_Cancelled","sum")).reset_index()},
           "Hotel Reservations — Teaching Extract",
           "Source: Hotel Reservations.csv (36,275 bookings, 2017–2018). Arrival_Date_ISO built from Y/M/D; 37 invalid Feb-29-2018 flagged.")
db_s = pd.read_csv(CLEAN/"diabetes_SAMPLE_5000.csv")
write_xlsx(XLSX_DIR/"diabetes_EXTRACT.xlsx",
           {"Screening_5000":db_s,
            "Age_Group_Summary": db_s.groupby("age_group").agg(n=("diabetes","count"), diabetes_rate=("diabetes","mean")).reset_index()},
           "Diabetes Screening — Teaching Extract (5,000-row sample)",
           "Source: diabetes_prediction_dataset.csv (100,000 rows). Sample preserves 8.5% diabetes rate. 'No Info' smoking = missing.")
sm_c = pd.read_csv(CLEAN/"social_media_addiction_CLEAN.csv")
write_xlsx(XLSX_DIR/"social_media_EXTRACT.xlsx",
           {"Students":sm_c,
            "Platform_Summary": sm_c.groupby("Most_Used_Platform").agg(n=("Student_ID","count"), avg_addiction=("Addicted_Score","mean")).reset_index()},
           "Social Media Addiction — Teaching Extract (full 705 rows)",
           "Source: Students Social Media Addiction.csv (705 students). Small clean file — ideal for correlation/regression practice.")
dep_a = pd.read_csv(CLEAN/"student_depression_ANONYMISED.csv")
write_xlsx(XLSX_DIR/"student_depression_EXTRACT_ANON.xlsx",
           {"Respondents_5000":dep_a.head(5000).reset_index(drop=True),
            "City_Summary": dep_a.groupby("City_Clean").agg(n=("Respondent_ID","count"), depression_rate=("Depression","mean")).reset_index()},
           "Student Depression (ANONYMISED) — Teaching Extract",
           "Source: Student Depression Dataset.csv (27,901 respondents). Sensitive data — anonymised IDs only. Work Pressure = blank for Students (not applicable).")

print("\nDirty/SQLite/XLSX build complete.")
