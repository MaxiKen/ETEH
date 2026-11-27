#!/usr/bin/env python3
"""
ETEH Bootcamp - Dataset build pipeline (Part 3 of 3)
Produces:
  * datasets/DATASET_GUIDE.md          (markdown reference used by the decks)
  * datasets/pdf/DATASET_DICTIONARY.pdf (printable column dictionary)
  * datasets/pdf/ETEH_*_sample.pdf      (printable sample extracts)
"""
import os, datetime as dt
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, PageBreak)

ROOT = "/home/user/ETEH"
OUT  = os.path.join(ROOT, "ETEH_Bootcamp", "datasets")
CLEAN= os.path.join(OUT, "clean")
PDF  = os.path.join(OUT, "pdf")
os.makedirs(PDF, exist_ok=True)

# ---------------------------------------------------------------------------
# Curated metadata: which dataset, domain, files, grain, dirty issues, modules
# ---------------------------------------------------------------------------
DATASETS = [
 dict(id="eteh_retail", brand="ETEH Retail", domain="Retail sales (Superstore)",
   clean=["eteh_retail.csv","eteh_retail.xlsx"], dirty=["eteh_retail.csv","eteh_retail.xlsx"],
   grain="One row per product line on an order",
   used_in="W1.6 data entry · W2 Pivots · W3 Dashboards · W4 Forecasting",
   dirty_issues=[
     "Mixed date formats in order_date / ship_date (ISO, US m/d/Y, 'Mon D, Y').",
     "Category / sub_category have stray upper/lower-case and trailing spaces.",
     "Ship Mode has typos ('Stan dard Class', 'SameDay').",
     "City / State padded with leading/trailing whitespace.",
     "~60 fully duplicated rows.",
     "Blank discount and profit cells."]),
 dict(id="eteh_commerce", brand="ETEH Commerce (SQL)", domain="Relational e-commerce (5 joined tables)",
   clean=["eteh_commerce.db (SQLite)","eteh_commerce.xlsx (5 sheets)","eteh_commerce_flat.csv"],
   dirty=["eteh_commerce_transactions.csv","eteh_commerce_products.csv"],
   grain="One row per transaction line; keys: customer_id, seller_id, product_id, delivery_id",
   used_in="W3 SQL, DAX & Dashboards (primary relational dataset)",
   dirty_issues=[
     "transaction_id duplicated (~26k repeats) in the raw file.",
     "delivery_id key spaces do not align between tables (re-link via order_id).",
     "A few orphan customer/product/seller keys.",
     "order_status mixed case; commission_amount stored as text ('NGN 12,345.67').",
     "order_date in mixed d/m/Y vs Y-m-d; 40 appended duplicate rows."]),
 dict(id="eteh_lifecare", brand="ETEH Lifecare Hospital", domain="Healthcare operations & billing",
   clean=["eteh_lifecare.csv","eteh_lifecare.xlsx"], dirty=["eteh_lifecare.csv","eteh_lifecare.xlsx"],
   grain="One row per hospital admission",
   used_in="W1 examples · W4 Descriptive stats & distributions · W3 Dashboards",
   dirty_issues=[
     "Patient names in random case ('Bobby JacksOn').",
     "Admission / discharge dates in mixed formats.",
     "Billing amount stored as text with currency symbol ('₦18,856.28').",
     "Medical condition / test results mixed case; some blank test_results.",
     "~50 duplicated rows."]),
 dict(id="eteh_university", brand="ETEH University – Student Analytics", domain="Education / student outcomes",
   clean=["eteh_university.xlsx (Performance, Wellbeing, Social_Media)"],
   dirty=["eteh_university.xlsx","eteh_uni_performance.csv","eteh_uni_wellbeing.csv","eteh_uni_social.csv"],
   grain="One row per student (three related surveys)",
   used_in="W4 Correlation, Regression & distributions (Crestview replacement)",
   dirty_issues=[
     "Performance: gender / parental_involvement mixed case; blank attendance & scores.",
     "Wellbeing: sleep_duration recoded inconsistently ('5 – 6 hrs').",
     "Social: platform names mixed case; duplicate rows."]),
 dict(id="eteh_hotels", brand="ETEH Hotels", domain="Hospitality / bookings",
   clean=["eteh_hotels.csv","eteh_hotels.xlsx"], dirty=["eteh_hotels.csv","eteh_hotels.xlsx"],
   grain="One row per room booking",
   used_in="W4 Logistic-style cancellation analysis & pivot practice",
   dirty_issues=[
     "meal_plan / room_type mixed case and spacing ('meal plan 1', 'room type 4').",
     "booking_status mixed case ('canceled' vs 'Canceled').",
     "Blank price / special-request cells; ~45 duplicate rows."]),
 dict(id="eteh_salary", brand="ETEH Salary", domain="HR / compensation",
   clean=["eteh_salary.csv","eteh_salary.xlsx"], dirty=["eteh_salary.csv","eteh_salary.xlsx"],
   grain="One row per employee",
   used_in="W4 Linear Regression (simple worked example)",
   dirty_issues=[
     "Education Level inconsistent ('Bachelor's' vs \"Bachelor's Degree\", 'phD' vs 'PhD').",
     "Blank age / gender / salary cells.",
     "Duplicate rows (kept intentionally as the teaching point)."]),
 dict(id="eteh_health_diabetes", brand="ETEH Health (Diabetes)", domain="Public-health screening",
   clean=["eteh_health_diabetes.csv","eteh_health_diabetes.xlsx"],
   dirty=["eteh_health_diabetes.csv","eteh_health_diabetes.xlsx"],
   grain="One row per patient screening (10,000-row sample)",
   used_in="W4 Distributions, correlation & classification practice",
   dirty_issues=[
     "smoking_history mixed case; gender lower-cased in places.",
     "Duplicate rows; large original (100k) sampled to 10k for performance."]),
]

# Column descriptions (curated; anything missing falls back to a friendly guess)
DESC = {
 "row_id":"Unique line identifier","order_id":"Order key (groups product lines)",
 "order_date":"Order date (YYYY-MM-DD)","ship_date":"Ship date (YYYY-MM-DD)",
 "ship_mode":"Shipping service","customer_id":"Customer key","customer_name":"Customer name",
 "segment":"Customer segment","country":"Country","city":"City","state":"State",
 "postal_code":"Postal code","region":"Sales region","product_id":"Product key",
 "category":"Product category","sub_category":"Product sub-category","product_name":"Product name",
 "sales":"Sales value (₦)","quantity":"Units sold","discount":"Discount fraction","profit":"Profit (₦, may be negative)",
 "ship_days":"Days order→ship",
 "patient_name":"Patient name","age":"Age (years)","gender":"Gender","blood_type":"Blood group",
 "medical_condition":"Primary diagnosis","admission_date":"Admission date","discharge_date":"Discharge date",
 "length_of_stay":"Days admitted","doctor":"Attending doctor","hospital":"Hospital",
 "insurance_provider":"Insurer","billing_amount":"Billing (₦)","room_number":"Room number",
 "admission_type":"Urgent/Emergency/Elective","medication":"Primary medication","test_results":"Normal/Inconclusive/Abnormal",
 "booking_id":"Booking key","num_adults":"Adults","num_children":"Children",
 "weekend_nights":"Weekend nights","week_nights":"Weekday nights","meal_plan":"Meal plan",
 "car_parking":"Parking required (0/1)","room_type":"Room type","lead_time":"Days booking→arrival",
 "market_segment":"Booking segment","repeated_guest":"Returning guest (0/1)","prev_cancellations":"Prior cancellations",
 "prev_not_canceled":"Prior completed stays","avg_price_per_room":"Avg room price (₦)","special_requests":"Special requests",
 "booking_status":"Canceled / Not Cancelled","arrival_date":"Arrival date",
 "education_level":"Highest education","job_title":"Job title","years_experience":"Years experience","salary":"Annual salary (₦)",
 "hypertension":"Hypertension flag (0/1)","heart_disease":"Heart-disease flag (0/1)","smoking_history":"Smoking status",
 "bmi":"Body mass index","hba1c_level":"HbA1c level","blood_glucose_level":"Blood glucose","diabetes":"Diabetes flag (0/1)",
 "transaction_id":"Transaction line key","first_name":"Customer first name","last_name":"Customer last name",
 "membership_tier":"Bronze/Silver/Gold","brand_name":"Product brand","seller_id":"Seller key","store_name":"Store name",
 "store_state":"Seller state","unit_price":"Unit price (₦)","total_price":"Line total (₦)",
 "commission_amount":"Platform commission (₦)","order_status":"Order status","payment_method":"Payment method",
 "payment_status":"Payment status","delivery_status":"Delivery status","courier_name":"Courier",
 "hours_studied":"Weekly study hours","attendance":"Attendance %","parental_involvement":"Parental involvement",
 "access_to_resources":"Resource access","extracurricular_activities":"Extracurricular (Yes/No)","sleep_hours":"Sleep hours",
 "previous_scores":"Prior exam score","motivation_level":"Motivation","internet_access":"Internet access (Yes/No)",
 "tutoring_sessions":"Tutoring sessions","family_income":"Family income","teacher_quality":"Teacher quality",
 "school_type":"School type","peer_influence":"Peer influence","physical_activity":"Physical activity hrs",
 "learning_disabilities":"Learning disability (Yes/No)","parental_education_level":"Parent education","distance_from_home":"Distance from home",
 "exam_score":"Final exam score",
 "id":"Respondent id","city":"City","profession":"Profession","academic_pressure":"Academic pressure (1-5)",
 "work_pressure":"Work pressure (1-5)","cgpa":"CGPA","study_satisfaction":"Study satisfaction (1-5)",
 "job_satisfaction":"Job satisfaction (1-5)","sleep_duration":"Sleep duration band","dietary_habits":"Diet quality",
 "degree":"Degree programme","have_you_ever_had_suicidal_thoughts_":"Suicidal thoughts (Yes/No)",
 "work/study_hours":"Work/study hours","financial_stress":"Financial stress (1-5)","family_history_of_mental_illness":"Family mental-illness history",
 "depression":"Depression flag (1/0)",
 "student_id":"Student id","academic_level":"Academic level","country":"Country",
 "avg_daily_usage_hours":"Avg daily social-media hrs","most_used_platform":"Top platform",
 "affects_academic_performance":"Affects academics (Yes/No)","sleep_hours_per_night":"Sleep hrs/night",
 "mental_health_score":"Mental-health score (1-10)","relationship_status":"Relationship status",
 "conflicts_over_social_media":"Conflicts over social media","addicted_score":"Addiction score (1-10)",
}
def guess(col):
    return col.replace('_',' ').strip().capitalize()
def describe(col):
    return DESC.get(col, guess(col))

# ---------------------------------------------------------------------------
# Build markdown guide
# ---------------------------------------------------------------------------
md = []
md.append("# ETEH Bootcamp — Dataset Guide\n")
md.append(f"_Generated {dt.date.today().isoformat()} · all datasets standardised, with deliberately dirtied working copies for cleaning labs._\n")
md.append("## Folder layout\n")
md.append("```\ndatasets/\n  clean/   # standardised 'gold' copies (CSV + XLSX)\n  dirty/   # intentionally messy working copies to clean\n  sql/     # eteh_commerce.db (SQLite) + queries.sql + expected outputs\n  pdf/     # printable data dictionary & sample extracts\n```\n")
for d in DATASETS:
    md.append(f"## {d['brand']}  (`{d['id']}`)\n")
    md.append(f"- **Domain:** {d['domain']}")
    md.append(f"- **Grain:** {d['grain']}")
    md.append(f"- **Used in:** {d['used_in']}")
    md.append(f"- **Clean files:** {', '.join(d['clean'])}")
    md.append(f"- **Dirty files:** {', '.join(d['dirty'])}")
    md.append("- **Intentional dirty issues (cleaning challenges):**")
    for i in d['dirty_issues']:
        md.append(f"  - {i}")
    md.append("")
md.append("## Suggested module → dataset map\n")
md.append("- **Week 1 (Foundations):** eteh_retail, eteh_lifecare (tidy-data & entry examples)\n"
          "- **Week 2 (Power Query & Pivots):** eteh_retail (pivots), eteh_hotels (cleaning), eteh_commerce_flat (transforms)\n"
          "- **Week 3 (SQL, DAX & Dashboards):** eteh_commerce.db (SQL joins), eteh_retail & eteh_lifecare (dashboards)\n"
          "- **Week 4 (Stats, Forecast & AI):** eteh_salary (regression), eteh_university (correlation), eteh_health_diabetes & eteh_lifecare (distributions), eteh_hotels (cancellation analysis)\n")
with open(os.path.join(OUT, "DATASET_GUIDE.md"), "w") as f:
    f.write("\n".join(md))
print("wrote DATASET_GUIDE.md")

# ---------------------------------------------------------------------------
# Build printable PDF dictionary
# ---------------------------------------------------------------------------
styles = getSampleStyleSheet()
H1 = ParagraphStyle('H1', parent=styles['Heading1'], textColor=colors.HexColor('#0B3D61'), spaceAfter=6)
H2 = ParagraphStyle('H2', parent=styles['Heading2'], textColor=colors.HexColor('#0B3D61'), spaceBefore=10, spaceAfter=4)
BODY = ParagraphStyle('BODY', parent=styles['BodyText'], fontSize=8.5, leading=11)
SMALL= ParagraphStyle('SMALL', parent=styles['BodyText'], fontSize=7.5, leading=9.5)
story=[]
story.append(Paragraph("ETEH Bootcamp — Dataset Dictionary", H1))
story.append(Paragraph(f"Standardised datasets with deliberately dirtied working copies · {dt.date.today().isoformat()}", BODY))
story.append(Spacer(1,6))

def type_of(series):
    if pd.api.types.is_integer_dtype(series): return "Integer"
    if pd.api.types.is_float_dtype(series): return "Decimal"
    if pd.api.types.is_datetime64_any_dtype(series): return "Date"
    return "Text"

for d in DATASETS:
    story.append(Paragraph(f"{d['brand']}  ·  <font size=8>{d['domain']}</font>", H2))
    story.append(Paragraph(f"<b>Grain:</b> {d['grain']} &nbsp; <b>Used in:</b> {d['used_in']}", SMALL))
    # pick the first clean csv/xlsx to read columns; for commerce use flat
    sample_file = None
    for f in d['clean']:
        if f.endswith('.csv'): sample_file = f; break
    if d['id']=='eteh_commerce': sample_file='eteh_commerce_flat.csv'
    if d['id']=='eteh_university': sample_file='eteh_uni_performance.csv'
    if sample_file is None:
        # xlsx only
        xf = [f for f in d['clean'] if f.endswith('.xlsx')][0]
        df = pd.read_excel(os.path.join(CLEAN, xf), sheet_name=0)
    else:
        df = pd.read_csv(os.path.join(CLEAN, sample_file))
    rows=[["Column","Type","Description"]]
    for c in df.columns:
        rows.append([c, type_of(df[c]), describe(c)])
    t=Table(rows, colWidths=[4.6*cm, 2.0*cm, 10.4*cm], repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#0B3D61')),
        ('TEXTCOLOR',(0,0),(-1,0),colors.white),
        ('FONTSIZE',(0,0),(-1,-1),7.5),
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
        ('GRID',(0,0),(-1,-1),0.4,colors.HexColor('#B8C4CC')),
        ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#EEF3F7')]),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('LEFTPADDING',(0,0),(-1,-1),3),('RIGHTPADDING',(0,0),(-1,-1),3),
    ]))
    story.append(Spacer(1,4)); story.append(t)
    story.append(Spacer(1,3))
    story.append(Paragraph("<b>Dirty issues to clean:</b> " + " ".join("•"+i for i in d['dirty_issues']), SMALL))
    story.append(PageBreak())

doc=SimpleDocTemplate(os.path.join(PDF,"DATASET_DICTIONARY.pdf"), pagesize=A4,
                      leftMargin=1.6*cm,rightMargin=1.6*cm,topMargin=1.4*cm,bottomMargin=1.4*cm,
                      title="ETEH Bootcamp Dataset Dictionary")
doc.build(story)
print("wrote DATASET_DICTIONARY.pdf")

# ---------------------------------------------------------------------------
# Printable sample extracts (first 22 rows of a few datasets)
# ---------------------------------------------------------------------------
def sample_pdf(src_csv, title, out_name, n=22):
    df=pd.read_csv(os.path.join(CLEAN,src_csv)).head(n)
    s=[Paragraph(title,H1),Spacer(1,6)]
    rows=[list(df.columns)]
    for _,r in df.iterrows():
        rows.append([("" if pd.isna(v) else str(v))[:22] for v in r.tolist()])
    w=[max(1.4*cm, (17*cm)/len(df.columns)) for _ in df.columns]
    t=Table(rows, colWidths=w, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#0B3D61')),
        ('TEXTCOLOR',(0,0),(-1,0),colors.white),
        ('FONTSIZE',(0,0),(-1,-1),6.4),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
        ('GRID',(0,0),(-1,-1),0.3,colors.HexColor('#B8C4CC')),
        ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#EEF3F7')]),
    ]))
    s.append(t)
    d=SimpleDocTemplate(os.path.join(PDF,out_name),pagesize=A4,
        leftMargin=1.6*cm,rightMargin=1.6*cm,topMargin=1.4*cm,bottomMargin=1.4*cm,title=title)
    d.build(s)
    print("wrote",out_name)

sample_pdf("eteh_retail.csv","ETEH Retail — Sample Extract (clean)","ETEH_Retail_sample.pdf")
sample_pdf("eteh_lifecare.csv","ETEH Lifecare Hospital — Sample Extract (clean)","ETEH_Lifecare_sample.pdf")
sample_pdf("eteh_commerce_flat.csv","ETEH Commerce — Sample Extract (clean, joined)","ETEH_Commerce_sample.pdf")
sample_pdf("eteh_salary.csv","ETEH Salary — Sample Extract (clean)","ETEH_Salary_sample.pdf")
print("DONE part 3")
