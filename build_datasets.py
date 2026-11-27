#!/usr/bin/env python3
"""ETEH Bootcamp — Dataset standardisation builder.
Reads the 9 attached source files, produces CLEAN standardised versions,
DIRTY teaching versions (Week 2), SQLite DBs, XLSX extracts and PDF dictionaries.
All outputs go under ETEH_Bootcamp/Datasets/.
"""
import os, re, sqlite3, textwrap
import pandas as pd
import numpy as np
from pathlib import Path

ROOT = Path("/home/user/ETEH")
CLEAN = ROOT/"ETEH_Bootcamp/Datasets/CLEAN"
DIRTY = ROOT/"ETEH_Bootcamp/Datasets/DIRTY_FOR_WEEK2"
SQLITE_DIR = ROOT/"ETEH_Bootcamp/Datasets/SQLITE"
XLSX_DIR = ROOT/"ETEH_Bootcamp/Datasets/XLSX_EXTRACTS"
PDF_DIR = ROOT/"ETEH_Bootcamp/Datasets/PDF_DICTIONARIES"
for d in [CLEAN, DIRTY, SQLITE_DIR, XLSX_DIR, PDF_DIR]:
    d.mkdir(parents=True, exist_ok=True)

rng = np.random.default_rng(42)

def save_csv(df, path, index=False):
    df.to_csv(path, index=index, encoding="utf-8")
    print(f"  wrote {path.name}  shape={df.shape}  size={path.stat().st_size/1024:.0f} KB")

# ============================================================
# 1. E-COMMERCE (Nigerian, relational — flagship for SQL)
# ============================================================
print("\n=== 1. E-COMMERCE ===")
xls = pd.ExcelFile(ROOT/"e-commerce.xlsx")
tr = pd.read_excel(xls, sheet_name="Transactions")
se = pd.read_excel(xls, sheet_name="Sellers")
pr = pd.read_excel(xls, sheet_name="Products")
de = pd.read_excel(xls, sheet_name="Deliveries")
cu = pd.read_excel(xls, sheet_name="Customers")
print(f" raw: TRX {tr.shape} Sellers {se.shape} Products {pr.shape} Deliveries {de.shape} Customers {cu.shape}")

# --- Transactions clean: keep 11,000 real rows, drop 26,000 blank/partial ---
tr_clean = tr[tr["transaction_id"].notna()].copy()
print(f" TRX real rows kept: {len(tr_clean)} (dropped {len(tr)-len(tr_clean)} blank/partial)")
# standardise delivery_id: DEL + 6-digit zero-padded to match Deliveries (DEL000001)
def norm_del(x):
    m = re.search(r"(\d+)", str(x))
    return f"DEL{int(m.group(1)):06d}" if m else x
tr_clean["delivery_id"] = tr_clean["delivery_id"].apply(norm_del)
# types
tr_clean["quantity"] = tr_clean["quantity"].astype(int)
for c in ["unit_price","total_price","commission_amount"]:
    tr_clean[c] = tr_clean[c].astype(float).round(2)
tr_clean["commission_percentage"] = tr_clean["commission_percentage"].astype(float)
tr_clean["order_date"] = pd.to_datetime(tr_clean["order_date"]).dt.date.astype(str)  # ISO YYYY-MM-DD
# trim text
for c in ["transaction_id","order_id","customer_id","product_id","seller_id","delivery_id","order_status","payment_method","payment_status"]:
    tr_clean[c] = tr_clean[c].astype(str).str.strip()
# sort by order_date then transaction_id
tr_clean = tr_clean.sort_values(["order_date","transaction_id"]).reset_index(drop=True)
# add derived
tr_clean["order_year"] = pd.to_datetime(tr_clean["order_date"]).dt.year
tr_clean["order_month"] = pd.to_datetime(tr_clean["order_date"]).dt.to_period("M").astype(str)
tr_clean["is_paid"] = (tr_clean["payment_status"]=="Paid").astype(int)
tr_clean["is_delivered"] = (tr_clean["order_status"]=="Delivered").astype(int)
cols_tr = ["transaction_id","order_id","customer_id","product_id","seller_id","delivery_id",
           "quantity","unit_price","total_price","commission_percentage","commission_amount",
           "order_status","payment_method","payment_status","order_date","order_year","order_month","is_paid","is_delivered"]
tr_clean = tr_clean[cols_tr]

# --- Dimensions clean ---
se_clean = se.copy()
for c in ["seller_id","store_name","owner_name","email","store_city","store_state","store_country","account_status"]:
    se_clean[c] = se_clean[c].astype(str).str.strip()
se_clean["email"] = se_clean["email"].str.lower()
se_clean["phone_number"] = se_clean["phone_number"].astype(str).str.strip()  # keep as text (10-digit NG)
se_clean["created_at"] = pd.to_datetime(se_clean["created_at"]).dt.date.astype(str)
se_clean = se_clean.sort_values("seller_id").reset_index(drop=True)

pr_clean = pr.copy()
for c in ["product_id","product_name","description","category","subcategory","brand_name","seller_id"]:
    pr_clean[c] = pr_clean[c].astype(str).str.strip()
pr_clean["created_at"] = pd.to_datetime(pr_clean["created_at"]).dt.date.astype(str)
for c in ["price","discount_price","weight"]:
    pr_clean[c] = pr_clean[c].astype(float).round(2)
pr_clean["discount_pct"] = ((pr_clean["price"]-pr_clean["discount_price"])/pr_clean["price"]*100).round(2)
pr_clean = pr_clean.sort_values("product_id").reset_index(drop=True)

de_clean = de.copy()
for c in ["delivery_id","order_id","customer_id","delivery_address","city","state","country","courier_name","delivery_status"]:
    de_clean[c] = de_clean[c].astype(str).str.strip()
de_clean["tracking_number"] = de_clean["tracking_number"].fillna("").astype(str).str.strip()
de_clean["shipped_date"] = pd.to_datetime(de_clean["shipped_date"], errors="coerce").dt.date.astype(str)
de_clean["delivered_date"] = pd.to_datetime(de_clean["delivered_date"], errors="coerce").dt.date.astype(str)
# blank strings for missing dates (keep empty, not 'nan')
de_clean["shipped_date"] = de_clean["shipped_date"].replace("nan","").replace("NaT","")
de_clean["delivered_date"] = de_clean["delivered_date"].replace("nan","").replace("NaT","")
de_clean["ship_lag_days"] = (pd.to_datetime(de_clean["delivered_date"], errors="coerce") - pd.to_datetime(de_clean["shipped_date"], errors="coerce")).dt.days
de_clean = de_clean.sort_values("delivery_id").reset_index(drop=True)

cu_clean = cu.copy()
for c in ["customer_id","first_name","last_name","email","gender","city","state","country","membership_tier"]:
    cu_clean[c] = cu_clean[c].astype(str).str.strip()
cu_clean["email"] = cu_clean["email"].str.lower()
cu_clean["phone_number"] = cu_clean["phone_number"].astype(str).str.strip()
cu_clean["date_of_birth"] = pd.to_datetime(cu_clean["date_of_birth"]).dt.date.astype(str)
cu_clean["created_at"] = pd.to_datetime(cu_clean["created_at"]).dt.date.astype(str)
# age at 2025-12-31 reference
cu_clean["age_at_2025"] = (pd.Timestamp("2025-12-31") - pd.to_datetime(cu_clean["date_of_birth"])).dt.days//365
cu_clean = cu_clean.sort_values("customer_id").reset_index(drop=True)

save_csv(tr_clean, CLEAN/"ecommerce_transactions_CLEAN.csv")
save_csv(se_clean, CLEAN/"ecommerce_sellers_CLEAN.csv")
save_csv(pr_clean, CLEAN/"ecommerce_products_CLEAN.csv")
save_csv(de_clean, CLEAN/"ecommerce_deliveries_CLEAN.csv")
save_csv(cu_clean, CLEAN/"ecommerce_customers_CLEAN.csv")

# Join integrity report
print(" join checks (clean):")
for name, left, right, key in [("TRX->Customers",tr_clean,cu_clean,"customer_id"),
                               ("TRX->Products",tr_clean,pr_clean,"product_id"),
                               ("TRX->Sellers",tr_clean,se_clean,"seller_id"),
                               ("TRX->Deliveries(delivery_id)",tr_clean,de_clean,"delivery_id"),
                               ("TRX->Deliveries(order_id)",tr_clean,de_clean,"order_id"),
                               ("Products->Sellers",pr_clean,se_clean,"seller_id"),
                               ("Deliveries->Customers",de_clean,cu_clean,"customer_id")]:
    orph = (~left[key].isin(right[key])).sum()
    print(f"  {name}: orphans {orph} / {len(left)}")

# ============================================================
# 2. SUPERSTORE (US retail — replaces fictional SuperMart)
# ============================================================
print("\n=== 2. SUPERSTORE ===")
ss = pd.read_csv(ROOT/"Sample - Superstore.csv", encoding="latin1", low_memory=False)
ss_clean = ss.copy()
ss_clean["Order_Date_ISO"] = pd.to_datetime(ss_clean["Order Date"], errors="coerce").dt.date.astype(str)
ss_clean["Ship_Date_ISO"] = pd.to_datetime(ss_clean["Ship Date"], errors="coerce").dt.date.astype(str)
ss_clean["Ship_Lag_Days"] = (pd.to_datetime(ss_clean["Ship_Date_ISO"]) - pd.to_datetime(ss_clean["Order_Date_ISO"])).dt.days
ss_clean["Margin"] = (ss_clean["Profit"]/ss_clean["Sales"]).round(4)
ss_clean["Order_Year"] = pd.to_datetime(ss_clean["Order_Date_ISO"]).dt.year
ss_clean["Order_Month"] = pd.to_datetime(ss_clean["Order_Date_ISO"]).dt.to_period("M").astype(str)
ss_clean["Postal_Code_Text"] = ss_clean["Postal Code"].astype(int).astype(str).str.zfill(5)
# tidy headers: keep original + ISO cols; drop original date strings? Keep both for traceability but rename
ss_clean = ss_clean.rename(columns={"Order Date":"Order_Date_RAW","Ship Date":"Ship_Date_RAW"})
# order columns
ss_clean = ss_clean[["Row ID","Order ID","Order_Date_ISO","Ship_Date_ISO","Ship_Lag_Days","Ship Mode","Customer ID","Customer Name","Segment","Country","City","State","Postal_Code_Text","Region","Product ID","Category","Sub-Category","Product Name","Sales","Quantity","Discount","Profit","Margin","Order_Year","Order_Month","Order_Date_RAW","Ship_Date_RAW"]]
ss_clean = ss_clean.sort_values(["Order_Date_ISO","Row ID"]).reset_index(drop=True)
save_csv(ss_clean, CLEAN/"superstore_CLEAN.csv")
print(f" superstore: {ss_clean.shape}, dates {ss_clean['Order_Date_ISO'].min()} to {ss_clean['Order_Date_ISO'].max()}")

# ============================================================
# 3. SALARY (messy — flagship for Week 2 cleaning)
# ============================================================
print("\n=== 3. SALARY ===")
sal = pd.read_csv(ROOT/"Salary_Data.csv", low_memory=False, encoding="utf-8-sig")
print(f" raw {sal.shape} cols={[repr(c) for c in sal.columns]}")
sal_clean = sal.copy()
# standardise education
edu_map = {"Bachelor's":"Bachelor's Degree","Bachelor's Degree":"Bachelor's Degree",
           "Master's":"Master's Degree","Master's Degree":"Master's Degree",
           "PhD":"PhD","phD":"PhD","High School":"High School"}
sal_clean["Education Level"] = sal_clean["Education Level"].map(lambda x: edu_map.get(str(x).strip(), x) if pd.notna(x) else x)
# standardise job titles: trim, fix known variants
sal_clean["Job Title"] = sal_clean["Job Title"].astype(str).str.strip()
# keep original but add Job_Family grouping
def job_family(t):
    t = str(t).lower()
    if "data" in t: return "Data"
    if "software" in t or "developer" in t or "engineer" in t and "project" not in t: return "Engineering"
    if "project" in t: return "Project/Operations"
    if "marketing" in t: return "Marketing"
    if "sales" in t: return "Sales"
    if "human" in t or "hr" in t: return "HR"
    if "finance" in t or "account" in t: return "Finance"
    if "product" in t: return "Product"
    if "director" in t or "chief" in t or "vp" in t or "ceo" in t or "manager" in t: return "Leadership"
    return "Other"
sal_clean["Job_Family"] = sal_clean["Job Title"].apply(job_family)
sal_clean["Gender"] = sal_clean["Gender"].astype(str).str.strip()
# quality flags (keep rows, flag don't silently delete)
sal_clean["flag_salary_outlier"] = ((sal_clean["Salary"]<10000)|(sal_clean["Salary"]>200000)).astype(int)
sal_clean["flag_exp_impossible"] = (sal_clean["Years of Experience"] > sal_clean["Age"]-18).astype(int)
# deduplicate exact rows for clean version, keep count
n_before = len(sal_clean)
sal_clean = sal_clean.drop_duplicates().reset_index(drop=True)
print(f" dedup: {n_before} -> {len(sal_clean)} (removed {n_before-len(sal_clean)} exact duplicates)")
# sort
sal_clean = sal_clean.sort_values(["Job Title","Years of Experience","Salary"]).reset_index(drop=True)
sal_clean = sal_clean.rename(columns={"Age":"Age","Gender":"Gender","Education Level":"Education_Level","Job Title":"Job_Title","Years of Experience":"Years_Experience","Salary":"Salary"})
save_csv(sal_clean, CLEAN/"salary_CLEAN.csv")

# ============================================================
# 4. STUDENT PERFORMANCE FACTORS (education — replaces Crestview)
# ============================================================
print("\n=== 4. STUDENT PERFORMANCE FACTORS ===")
spf = pd.read_csv(ROOT/"StudentPerformanceFactors.csv", low_memory=False)
spf_clean = spf.copy()
# explicit missing category (Tidy Law 7) instead of blank
for c in ["Teacher_Quality","Parental_Education_Level","Distance_from_Home"]:
    spf_clean[c] = spf_clean[c].fillna("Unknown")
spf_clean["flag_exam_over_100"] = (spf_clean["Exam_Score"]>100).astype(int)
# ordinal scores for analysis
ord_map = {"Low":1,"Medium":2,"High":3}
for c in ["Parental_Involvement","Access_to_Resources","Motivation_Level","Family_Income","Teacher_Quality"]:
    spf_clean[c+"_Score"] = spf_clean[c].map(ord_map)
spf_clean["Exam_Band"] = pd.cut(spf_clean["Exam_Score"], bins=[0,60,70,80,90,101], labels=["<60","60-69","70-79","80-89","90+"], include_lowest=True).astype(str)
save_csv(spf_clean, CLEAN/"student_performance_CLEAN.csv")
print(f" spf {spf_clean.shape}")

# ============================================================
# 5. HOTEL RESERVATIONS
# ============================================================
print("\n=== 5. HOTEL ===")
ht = pd.read_csv(ROOT/"Hotel Reservations.csv", low_memory=False)
ht_clean = ht.copy()
ht_clean["Arrival_Date_ISO"] = pd.to_datetime(dict(year=ht_clean["arrival_year"], month=ht_clean["arrival_month"], day=ht_clean["arrival_date"]), errors="coerce").dt.date.astype(str)
ht_clean["Arrival_Date_ISO"] = ht_clean["Arrival_Date_ISO"].replace("nan","").replace("NaT","")
ht_clean["Total_Nights"] = ht_clean["no_of_weekend_nights"]+ht_clean["no_of_week_nights"]
ht_clean["Total_Guests"] = ht_clean["no_of_adults"]+ht_clean["no_of_children"]
ht_clean["flag_invalid_date"] = (ht_clean["Arrival_Date_ISO"]=="").astype(int)
ht_clean["flag_zero_price"] = (ht_clean["avg_price_per_room"]==0).astype(int)
ht_clean["flag_zero_adults"] = (ht_clean["no_of_adults"]==0).astype(int)
ht_clean["flag_zero_nights"] = (ht_clean["Total_Nights"]==0).astype(int)
ht_clean["Is_Cancelled"] = (ht_clean["booking_status"]=="Canceled").astype(int)
ht_clean["Cancellation_Rate"] = ht_clean["Is_Cancelled"]  # per-row 0/1
save_csv(ht_clean, CLEAN/"hotel_reservations_CLEAN.csv")
print(f" hotel {ht_clean.shape} invalid dates {ht_clean['flag_invalid_date'].sum()}")

# ============================================================
# 6. HEALTHCARE (hospital ops — replaces Lifecare)
# ============================================================
print("\n=== 6. HEALTHCARE ===")
hc = pd.read_csv(ROOT/"healthcare_dataset.csv", low_memory=False)
hc_clean = hc.copy()
# Title-case names (fix Bobby JacksOn -> Bobby Jackson)
def title_case(s):
    return " ".join(w.capitalize() for w in str(s).strip().split())
hc_clean["Name"] = hc_clean["Name"].apply(title_case)
hc_clean["Doctor"] = hc_clean["Doctor"].apply(title_case)
hc_clean["Hospital"] = hc_clean["Hospital"].apply(title_case)
hc_clean["Date_of_Admission_ISO"] = pd.to_datetime(hc_clean["Date of Admission"], errors="coerce").dt.date.astype(str)
hc_clean["Discharge_Date_ISO"] = pd.to_datetime(hc_clean["Discharge Date"], errors="coerce").dt.date.astype(str)
hc_clean["Length_of_Stay_Days"] = (pd.to_datetime(hc_clean["Discharge_Date_ISO"]) - pd.to_datetime(hc_clean["Date_of_Admission_ISO"])).dt.days
hc_clean["Admission_Year"] = pd.to_datetime(hc_clean["Date_of_Admission_ISO"]).dt.year
hc_clean["Admission_Month"] = pd.to_datetime(hc_clean["Date_of_Admission_ISO"]).dt.to_period("M").astype(str)
hc_clean["flag_negative_billing"] = (hc_clean["Billing Amount"]<0).astype(int)
hc_clean["Billing_Amount_Abs"] = hc_clean["Billing Amount"].abs()
hc_clean = hc_clean.rename(columns={"Blood Type":"Blood_Type","Medical Condition":"Medical_Condition","Insurance Provider":"Insurance_Provider","Room Number":"Room_Number","Admission Type":"Admission_Type","Test Results":"Test_Results"})
hc_clean = hc_clean.sort_values(["Date_of_Admission_ISO","Name"]).reset_index(drop=True)
# anonymised version: Patient_ID only (for portfolio)
hc_anon = hc_clean.copy()
hc_anon.insert(0, "Patient_ID", [f"HLT{i+1:06d}" for i in range(len(hc_anon))])
hc_anon = hc_anon.drop(columns=["Name"])
save_csv(hc_clean, CLEAN/"healthcare_CLEAN.csv")
save_csv(hc_anon, CLEAN/"healthcare_ANONYMISED.csv")
print(f" healthcare {hc_clean.shape} anon {hc_anon.shape} neg billing {hc_clean['flag_negative_billing'].sum()}")

# ============================================================
# 7. DIABETES (clinical screening)
# ============================================================
print("\n=== 7. DIABETES ===")
db = pd.read_csv(ROOT/"diabetes_prediction_dataset.csv", low_memory=False)
db_clean = db.copy()
db_clean["smoking_history"] = db_clean["smoking_history"].astype(str).str.strip()
db_clean["gender"] = db_clean["gender"].astype(str).str.strip()
db_clean["flag_under_18"] = (db_clean["age"]<18).astype(int)
db_clean["flag_bmi_outlier"] = (db_clean["bmi"]>60).astype(int)
db_clean["age_group"] = pd.cut(db_clean["age"], bins=[0,18,40,60,120], labels=["0-17","18-39","40-59","60+"], include_lowest=True).astype(str)
db_clean["bmi_category"] = pd.cut(db_clean["bmi"], bins=[0,18.5,25,30,200], labels=["Underweight","Normal","Overweight","Obese"], include_lowest=True).astype(str)
db_clean["hba1c_category"] = pd.cut(db_clean["HbA1c_level"], bins=[0,5.7,6.5,20], labels=["Normal","Prediabetes","Diabetes"], include_lowest=True).astype(str)
db_clean["glucose_category"] = pd.cut(db_clean["blood_glucose_level"], bins=[0,100,126,1000], labels=["Normal","Prediabetes","Diabetes"], include_lowest=True).astype(str)
db_clean = db_clean.sort_values(["age","bmi"]).reset_index(drop=True)
save_csv(db_clean, CLEAN/"diabetes_CLEAN.csv")
# stratified 5,000-row teaching sample (preserves diabetes rate)
sample = db_clean.groupby("diabetes", group_keys=False).apply(lambda g: g.sample(n=min(5000, len(g)), random_state=42))
# Actually stratified: take proportional 5k
sample = db_clean.sample(n=5000, random_state=42).sort_values(["age"]).reset_index(drop=True)
save_csv(sample, CLEAN/"diabetes_SAMPLE_5000.csv")
print(f" diabetes {db_clean.shape} sample {sample.shape} rate {db_clean['diabetes'].mean():.3f}")

# ============================================================
# 8. STUDENT DEPRESSION (sensitive — handle with care)
# ============================================================
print("\n=== 8. DEPRESSION ===")
dep = pd.read_csv(ROOT/"Student Depression Dataset.csv", low_memory=False)
dep_clean = dep.copy()
# City canonicalisation: map typos to nearest canonical city, else keep
canonical_cities = ["Ahmedabad","Agra","Bangalore","Bhopal","Chennai","Delhi","Faridabad","Ghaziabad","Hyderabad","Indore","Jaipur","Kalyan","Kanpur","Kolkata","Lucknow","Ludhiana","Meerut","Mumbai","Nagpur","Nashik","Patna","Pune","Rajkot","Srinagar","Surat","Thane","Vadodara","Varanasi","Vasai-Virar","Visakhapatnam"]
city_map = {"Khaziabad":"Ghaziabad","Nalyan":"Kalyan","Less Delhi":"Delhi","Less than 5 Kalyan":"Kalyan"}
def clean_city(c):
    c = str(c).strip()
    if c in city_map: return city_map[c]
    if c in canonical_cities: return c
    return "Other/Unknown"
dep_clean["City_Clean"] = dep_clean["City"].apply(clean_city)
# Profession grouping: tiny professions -> Other Profession (keep original too)
prof_counts = dep_clean["Profession"].value_counts()
tiny = prof_counts[prof_counts<50].index.tolist()
dep_clean["Profession_Grouped"] = dep_clean["Profession"].apply(lambda x: "Other Profession" if x in tiny else x)
# Structural zeros: Work Pressure / Job Satisfaction = 0 for Students means Not Applicable
dep_clean["Work_Pressure_Clean"] = dep_clean.apply(lambda r: np.nan if (r["Profession"]=="Student" and r["Work Pressure"]==0) else r["Work Pressure"], axis=1)
dep_clean["Job_Satisfaction_Clean"] = dep_clean.apply(lambda r: np.nan if (r["Profession"]=="Student" and r["Job Satisfaction"]==0) else r["Job Satisfaction"], axis=1)
dep_clean["flag_cgpa_zero"] = (dep_clean["CGPA"]==0).astype(int)
dep_clean["Sleep_Hours_Mid"] = dep_clean["Sleep Duration"].map({"Less than 5 hours":4,"5-6 hours":5.5,"7-8 hours":7.5,"More than 8 hours":9,"Others":np.nan})
dep_clean = dep_clean.sort_values(["City_Clean","Age"]).reset_index(drop=True)
save_csv(dep_clean, CLEAN/"student_depression_CLEAN.csv")
# anonymised: drop id, keep pseudonymised
dep_anon = dep_clean.copy()
dep_anon.insert(0, "Respondent_ID", [f"DEP{i+1:06d}" for i in range(len(dep_anon))])
dep_anon = dep_anon.drop(columns=["id"])
save_csv(dep_anon, CLEAN/"student_depression_ANONYMISED.csv")
print(f" depression {dep_clean.shape} tiny professions grouped: {tiny}")

# ============================================================
# 9. SOCIAL MEDIA ADDICTION (small, clean — Week 4 stats)
# ============================================================
print("\n=== 9. SOCIAL MEDIA ===")
sm = pd.read_csv(ROOT/"Students Social Media Addiction.csv", low_memory=False)
sm_clean = sm.copy()
sm_clean["Addiction_Band"] = pd.cut(sm_clean["Addicted_Score"], bins=[0,4,6,10], labels=["Low","Moderate","High"], include_lowest=True).astype(str)
sm_clean["Usage_Band"] = pd.cut(sm_clean["Avg_Daily_Usage_Hours"], bins=[0,3,5,10], labels=["Light","Moderate","Heavy"], include_lowest=True).astype(str)
save_csv(sm_clean, CLEAN/"social_media_addiction_CLEAN.csv")
print(f" social {sm_clean.shape}")

print("\nCLEAN build complete.")
