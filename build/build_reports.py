"""
ETEH Bootcamp - Phase A4: Professional data documentation.

Outputs:
  Datasets/Reports/data_catalogue.pdf  -- a printed-ready catalogue of every dataset
                                         (overview + per-dataset schema, samples, notes).
  Datasets/data_dictionary.xlsx        -- master workbook: Overview sheet + one sheet
                                         per dataset (field / type / description / sample).
"""
import os, datetime
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
                                PageBreak, ListFlowable, ListItem)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CLEAN = os.path.join(ROOT, "Datasets", "Clean")
REP   = os.path.join(ROOT, "Datasets", "Reports")
os.makedirs(REP, exist_ok=True)

# Human-friendly metadata: dataset -> (title, domain, description, primary_use, column_descriptions)
META = {
 "ecommerce_transactions_reconciled.csv": ("E-Commerce — Transactions", "Retail / Marketplace",
    "Line-item sales transactions for a Nigerian online marketplace. Fully reconciled so every foreign key resolves to customers, products, sellers and deliveries.",
    "SQL joins, relational modelling, revenue/commission analytics (Week 3).",
    {"transaction_id":"Unique transaction id (PK)","order_id":"Marketplace order id","customer_id":"FK -> customers","product_id":"FK -> products",
     "seller_id":"FK -> sellers (derived from product)","delivery_id":"FK -> deliveries","quantity":"Units purchased","unit_price":"Price per unit (NGN)",
     "total_price":"quantity x unit_price","commission_percentage":"Platform commission %","commission_amount":"Commission earned (NGN)",
     "order_status":"Fulfilment status","payment_method":"e.g. Flutterwave","payment_status":"Paid / Pending","order_date":"Order date (ISO)"}),
 "ecommerce_customers.csv": ("E-Commerce — Customers", "Retail / Marketplace",
    "Shopper master records with location, membership tier and lifetime totals.",
    "Customer segmentation, LTV (Weeks 2-4).",
    {"customer_id":"Unique customer id (PK)","first_name":"Given name","last_name":"Family name","email":"Contact email","phone_number":"Phone",
     "gender":"Male / Female","date_of_birth":"DOB (ISO)","city":"City","state":"State","country":"Country","membership_tier":"Bronze/Silver/Gold",
     "loyalty_points":"Loyalty balance","total_orders":"Lifetime orders","total_spent":"Lifetime spend (NGN)","created_at":"Account created (ISO)"}),
 "ecommerce_sellers.csv": ("E-Commerce — Sellers", "Retail / Marketplace",
    "Marketplace vendor/store records with ratings and revenue.",
    "Seller performance ranking (Week 3 SQL).",
    {"seller_id":"Unique seller id (PK)","store_name":"Store/brand name","owner_name":"Owner name","email":"Contact email","phone_number":"Phone",
     "store_city":"Store city","store_state":"Store state","store_country":"Store country","rating":"Avg rating 1-5","total_sales":"Units sold",
     "total_revenue":"Revenue (NGN)","account_status":"Active/Inactive","created_at":"Joined (ISO)"}),
 "ecommerce_products.csv": ("E-Commerce — Products", "Retail / Marketplace",
    "Product catalogue with category, brand, price and stock.",
    "Catalogue analysis, sell-through (Weeks 2-3).",
    {"product_id":"Unique product id (PK)","product_name":"Product name","description":"Short description","category":"Top category","subcategory":"Sub-category",
     "brand_name":"Brand","seller_id":"FK -> sellers","price":"List price (NGN)","discount_price":"Discounted price","commission_percentage":"Commission %",
     "stock_quantity":"Units in stock","weight":"Weight (kg)","total_sold":"Units sold","average_rating":"Avg rating","created_at":"Listed (ISO)"}),
 "ecommerce_deliveries.csv": ("E-Commerce — Deliveries", "Retail / Marketplace",
    "Shipment records linking orders to couriers and addresses.",
    "Logistics / delivery-time analysis (Week 3).",
    {"delivery_id":"Unique delivery id (PK)","order_id":"Marketplace order id","customer_id":"FK -> customers","delivery_address":"Street address",
     "city":"City","state":"State","country":"Country","courier_name":"Courier","tracking_number":"Tracking no.","delivery_status":"Status",
     "shipped_date":"Shipped (ISO)","delivered_date":"Delivered (ISO)"}),
 "healthcare_records.csv": ("Healthcare Records", "Healthcare",
    "Hospital admission records: patient demographics, condition, billing, stay.",
    "Tidy-data, cleaning & stats exercises (Weeks 1, 4).",
    {"name":"Patient name","age":"Age (years)","gender":"Male/Female","blood_type":"ABO blood type","medical_condition":"Diagnosis",
     "date_of_admission":"Admission (ISO)","doctor":"Attending doctor","hospital":"Facility","insurance_provider":"Insurer",
     "billing_amount":"Billing (NGN)","room_number":"Room #","admission_type":"Urgent/Emergency/Elective","discharge_date":"Discharge (ISO)",
     "medication":"Medication","test_results":"Normal/Abnormal/Inconclusive","length_of_stay":"Days admitted"}),
 "salary_data.csv": ("Salary Survey", "HR / Compensation",
    "Employee compensation survey with experience and education.",
    "Regression & correlation (Week 4 stats/forecast).",
    {"age":"Age (years)","gender":"Male/Female/Other","education_level":"Highest education","job_title":"Role","years_of_experience":"Experience (yrs)","salary":"Annual salary"}),
 "superstore_sales.csv": ("Superstore Sales", "Retail",
    "US office-supplies superstore orders with ship mode, geography and profit.",
    "PivotTables, dashboards, forecasting (Weeks 2, 4).",
    {"row_id":"Row id","order_id":"Order id","order_date":"Order (ISO)","ship_date":"Shipped (ISO)","ship_mode":"Delivery mode","customer_id":"Customer id",
     "customer_name":"Customer","segment":"Consumer/Corporate/Home Office","country":"Country","city":"City","state":"State","postal_code":"ZIP",
     "region":"Region","product_id":"Product id","category":"Category","sub_category":"Sub-category","product_name":"Product","sales":"Sales",
     "quantity":"Qty","discount":"Discount %","profit":"Profit","ship_days":"Transit days","profit_margin":"Profit/Sales","year":"Order year"}),
 "student_performance.csv": ("Student Performance Factors", "Education",
    "Factors influencing exam scores for secondary students.",
    "Multiple regression, correlation (Week 4).",
    {"hours_studied":"Weekly study hours","attendance":"Attendance %","parental_involvement":"Low/Med/High","access_to_resources":"Low/Med/High",
     "extracurricular_activities":"Yes/No","sleep_hours":"Sleep hours","previous_scores":"Prior avg score","motivation_level":"Low/Med/High",
     "internet_access":"Yes/No","tutoring_sessions":"Sessions","family_income":"Low/Med/High","teacher_quality":"Low/Med/High","school_type":"Public/Private",
     "peer_influence":"Positive/Neutral/Negative","physical_activity":"Activity hrs","learning_disabilities":"Yes/No","parental_education_level":"Education",
     "distance_from_home":"Near/Moderate/Far","gender":"Male/Female","exam_score":"Final exam score (0-100)"}),
 "student_depression.csv": ("Student Depression Study", "Public Health",
    "Lifestyle, academic and mental-health survey of students & young workers.",
    "Classification, group comparison (Week 4).",
    {"id":"Respondent id","gender":"Male/Female","age":"Age","city":"City","profession":"Profession","academic_pressure":"0-5","work_pressure":"0-5",
     "cgpa":"CGPA","study_satisfaction":"0-5","job_satisfaction":"0-4","sleep_duration":"Sleep band","dietary_habits":"Healthy/Moderate/Unhealthy",
     "degree":"Highest degree","suicidal_thoughts":"Yes/No","work_study_hours":"Hours","financial_stress":"1-5","family_history_of_mental_illness":"Yes/No","depression":"1 if depressed"}),
 "social_media_addiction.csv": ("Social Media & Wellbeing", "Public Health",
    "Small survey linking social-media usage to wellbeing and academics.",
    "Simple relatable examples, correlation (Weeks 1, 4).",
    {"student_id":"Id","age":"Age","gender":"Male/Female","academic_level":"Level","country":"Country","avg_daily_usage_hours":"Usage hrs/day",
     "most_used_platform":"Platform","affects_academic_performance":"Yes/No","sleep_hours_per_night":"Sleep hrs","mental_health_score":"1-10",
     "relationship_status":"Status","conflicts_over_social_media":"Conflicts","addicted_score":"2-9"}),
 "diabetes_prediction.csv": ("Diabetes Prediction", "Healthcare / ML",
    "Biometric profile labelled for diabetes presence.",
    "Classification, feature analysis (Week 4 / AI).",
    {"gender":"Male/Female/Other","age":"Age (years)","hypertension":"0/1","heart_disease":"0/1","smoking_history":"Smoking status",
     "bmi":"Body-mass index","hbA1c_level":"HbA1c","blood_glucose_level":"Glucose","diabetes":"1 if diabetic"}),
 "hotel_reservations.csv": ("Hotel Reservations", "Hospitality",
    "Hotel booking records with cancellation outcome.",
    "Classification (will-cancel?), cleaning (Week 4).",
    {"booking_id":"Booking id","no_of_adults":"Adults","no_of_children":"Children","no_of_weekend_nights":"Weekend nights","no_of_week_nights":"Week nights",
     "type_of_meal_plan":"Meal plan","required_car_parking_space":"0/1","room_type_reserved":"Room type","lead_time":"Days to arrival","arrival_year":"Year",
     "arrival_month":"Month","arrival_date":"Day","market_segment_type":"Segment","repeated_guest":"0/1","no_of_previous_cancellations":"Prev cancels",
     "no_of_previous_bookings_not_canceled":"Prev kept","avg_price_per_room":"Avg price","no_of_special_requests":"Requests","booking_status":"Canceled/Not Canceled"}),
}

ORDER = ["ecommerce_transactions_reconciled.csv","ecommerce_customers.csv","ecommerce_sellers.csv",
         "ecommerce_products.csv","ecommerce_deliveries.csv","healthcare_records.csv","salary_data.csv",
         "superstore_sales.csv","student_performance.csv","student_depression.csv","social_media_addiction.csv",
         "diabetes_prediction.csv","hotel_reservations.csv"]

def dtype_name(s):
    try:
        pd.to_numeric(s.replace("", pd.NA).dropna())
        return "numeric"
    except Exception:
        return "text"

def build_pdf():
    styles = getSampleStyleSheet()
    h1 = ParagraphStyle("h1", parent=styles["Heading1"], textColor=colors.HexColor("#11487a"), spaceAfter=8)
    h2 = ParagraphStyle("h2", parent=styles["Heading2"], textColor=colors.HexColor("#1f6fb2"), spaceBefore=10, spaceAfter=4)
    body = ParagraphStyle("body", parent=styles["BodyText"], fontSize=9, leading=12)
    small = ParagraphStyle("small", parent=styles["BodyText"], fontSize=8, leading=10, textColor=colors.HexColor("#555555"))
    doc = SimpleDocTemplate(os.path.join(REP, "data_catalogue.pdf"), pagesize=A4,
                            topMargin=1.6*cm, bottomMargin=1.4*cm, leftMargin=1.6*cm, rightMargin=1.6*cm,
                            title="ETEH Bootcamp Data Catalogue")
    E = []
    E.append(Paragraph("ETEH Ltd &mdash; Four-Week Excel Data Analytics Bootcamp", h1))
    E.append(Paragraph("Data Catalogue &amp; Dictionary", h2))
    E.append(Paragraph(f"Generated {datetime.date.today().isoformat()} &middot; {len(ORDER)} curated datasets "
                       f"(standardised clean versions + intentionally dirty practice copies).", small))
    E.append(Spacer(1, 0.3*cm))

    # Overview table
    ov = [["#","Dataset","Domain","Rows","Cols","Primary use"]]
    for i, f in enumerate(ORDER, 1):
        df = pd.read_csv(os.path.join(CLEAN, f), nrows=0)
        # full counts
        n = sum(1 for _ in open(os.path.join(CLEAN, f))) - 1
        m = META[f]
        ov.append([str(i), m[0], m[1], f"{n:,}", str(len(df.columns)), m[3]])
    t = Table(ov, colWidths=[0.8*cm,4.4*cm,2.8*cm,2.0*cm,1.3*cm,5.0*cm], repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#11487a")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTSIZE",(0,0),(-1,-1),7.5),
        ("GRID",(0,0),(-1,-1),0.4,colors.HexColor("#bbbbbb")),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#eef3f8")]),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ]))
    E.append(t)
    E.append(PageBreak())

    for f in ORDER:
        title, domain, desc, use, cols = META[f]
        df = pd.read_csv(os.path.join(CLEAN, f), dtype=str, keep_default_na=False)
        E.append(Paragraph(title, h2))
        E.append(Paragraph(f"<b>Domain:</b> {domain} &nbsp;|&nbsp; <b>Shape:</b> {len(df):,} rows &times; {df.shape[1]} columns", small))
        E.append(Paragraph(desc, body))
        E.append(Paragraph(f"<b>Primary bootcamp use:</b> {use}", small))
        E.append(Spacer(1, 0.15*cm))
        # field table
        rows = [["Field","Type","Description","Sample"]]
        for c in df.columns:
            d = cols.get(c, "")
            sample = df[c].iloc[0] if len(df) else ""
            sample = (sample[:22] + "…") if len(str(sample)) > 23 else sample
            rows.append([c, dtype_name(df[c]), d, str(sample)])
        ft = Table(rows, colWidths=[3.6*cm,1.7*cm,7.4*cm,3.4*cm], repeatRows=1)
        ft.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#1f6fb2")),
            ("TEXTCOLOR",(0,0),(-1,0),colors.white),
            ("FONTSIZE",(0,0),(-1,-1),7),
            ("GRID",(0,0),(-1,-1),0.3,colors.HexColor("#cccccc")),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#f3f7fb")]),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ]))
        E.append(ft)
        E.append(PageBreak())
    doc.build(E)
    print("Wrote", os.path.join(REP, "data_catalogue.pdf"))

def build_xlsx():
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    wb = Workbook()
    ov = wb.active; ov.title = "Overview"
    head = Font(bold=True, color="FFFFFF"); fill = PatternFill("solid", fgColor="11487A")
    ov.append(["#","Dataset","Domain","Rows","Columns","Primary Bootcamp Use"])
    for c in range(1,7): ov.cell(1,c).font=head; ov.cell(1,c).fill=fill
    for i,f in enumerate(ORDER,1):
        df = pd.read_csv(os.path.join(CLEAN,f), dtype=str, keep_default_na=False)
        m = META[f]
        ov.append([i, m[0], m[1], len(df), df.shape[1], m[3]])
    for col,w in zip("ABCDEF",[4,34,18,10,10,52]): ov.column_dimensions[col].width=w

    for f in ORDER:
        title, domain, desc, use, cols = META[f]
        df = pd.read_csv(os.path.join(CLEAN, f), dtype=str, keep_default_na=False)
        ws = wb.create_sheet(title[:31])
        ws.append([title]); ws["A1"].font=Font(bold=True, size=13, color="11487A")
        ws.append([f"Domain: {domain}   |   Shape: {len(df):,} x {df.shape[1]}"])
        ws.append([desc])
        ws.append([f"Primary use: {use}"])
        ws.append([])
        ws.append(["Field","Type","Description","Sample Value"])
        for c in range(1,5):
            ws.cell(ws.max_row,c).font=head; ws.cell(ws.max_row,c).fill=fill
        for col in df.columns:
            sample = df[col].iloc[0] if len(df) else ""
            ws.append([col, dtype_name(df[col]), cols.get(col,""), str(sample)[:40]])
        for col,w in zip("ABCD",[34,12,46,30]): ws.column_dimensions[col].width=w
    out = os.path.join(ROOT, "Datasets", "data_dictionary.xlsx")
    wb.save(out)
    print("Wrote", out)

if __name__ == "__main__":
    build_pdf()
    build_xlsx()
