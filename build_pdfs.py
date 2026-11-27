#!/usr/bin/env python3
"""Build PDF data dictionaries + master DATASET_GUIDE.pdf"""
import pandas as pd
from pathlib import Path
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white, black
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, HRFlowable
from reportlab.lib import colors

ROOT = Path("/home/user/ETEH")
CLEAN = ROOT/"ETEH_Bootcamp/Datasets/CLEAN"
PDF_DIR = ROOT/"ETEH_Bootcamp/Datasets/PDF_DICTIONARIES"
PDF_DIR.mkdir(parents=True, exist_ok=True)

NAVY = HexColor("#1F3864"); TEAL = HexColor("#0E6B7A"); LIGHT = HexColor("#D9E2F3"); GREY = HexColor("#F2F2F2")
styles = getSampleStyleSheet()
sTitle = ParagraphStyle("T", parent=styles["Title"], fontSize=20, textColor=NAVY, spaceAfter=4)
sH1 = ParagraphStyle("H1", parent=styles["Heading1"], fontSize=13, textColor=NAVY, spaceBefore=12, spaceAfter=6)
sH2 = ParagraphStyle("H2", parent=styles["Heading2"], fontSize=11, textColor=TEAL, spaceBefore=10, spaceAfter=4)
sBody = ParagraphStyle("B", parent=styles["BodyText"], fontSize=9, leading=13)
sSmall = ParagraphStyle("S", parent=styles["BodyText"], fontSize=8, leading=11, textColor=HexColor("#404040"))
sCell = ParagraphStyle("C", parent=styles["BodyText"], fontSize=7.5, leading=10)
sCellH = ParagraphStyle("CH", parent=styles["BodyText"], fontSize=7.5, leading=10, textColor=white, fontName="Helvetica-Bold")

def header_footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 7); canvas.setFillColor(HexColor("#595959"))
    canvas.drawString(15*mm, 287*mm, "ETEH Ltd  ·  Eaglet2EagleHub Limited  —  Excel Data Analytics Bootcamp")
    canvas.drawRightString(195*mm, 287*mm, f"Page {doc.page}")
    canvas.setStrokeColor(TEAL); canvas.setLineWidth(0.6)
    canvas.line(15*mm, 285*mm, 195*mm, 285*mm)
    canvas.restoreState()

def dict_table(df, filename, source_note, grain, keys, cleaning):
    path = PDF_DIR/f"{filename}.pdf"
    doc = SimpleDocTemplate(str(path), pagesize=A4, leftMargin=15*mm, rightMargin=15*mm, topMargin=18*mm, bottomMargin=15*mm,
                            title=f"ETEH Data Dictionary — {filename}", author="ETEH Ltd")
    story = []
    story.append(Paragraph(f"ETEH Data Dictionary", sTitle))
    story.append(Paragraph(f"<b>{filename}</b> &nbsp;·&nbsp; {source_note}", sBody))
    story.append(HRFlowable(width="100%", color=TEAL, thickness=1))
    story.append(Paragraph(f"<b>Grain (one row = ):</b> {grain}", sBody))
    story.append(Paragraph(f"<b>Keys:</b> {keys}", sBody))
    story.append(Paragraph(f"<b>Rows:</b> {len(df):,} &nbsp;·&nbsp; <b>Columns:</b> {len(df.columns)}", sBody))
    story.append(Paragraph("<b>Standardisation applied:</b> " + cleaning, sSmall))
    story.append(Spacer(1,4))
    # column table
    rows = [[Paragraph("<b>Column</b>",sCellH), Paragraph("<b>Type</b>",sCellH), Paragraph("<b>Example</b>",sCellH), Paragraph("<b>Missing</b>",sCellH), Paragraph("<b>Notes</b>",sCellH)]]
    for c in df.columns:
        ex = df[c].dropna().iloc[0] if df[c].dropna().shape[0]>0 else ""
        exs = str(ex)[:42].replace("&","&amp;").replace("<","&lt;")
        miss = int(df[c].isna().sum())
        dtype = str(df[c].dtype)
        # notes heuristics
        note=""
        if "flag_" in c: note="Quality flag (1=investigate). Keep row; do not silently delete."
        elif c.endswith("_ISO") or "date" in c.lower(): note="ISO YYYY-MM-DD. Never M/D/YYYY."
        elif "id" in c.lower() or "ID" in c: note="Key. Must be unique at declared grain."
        elif "Score" in c or "score" in c: note="Ordinal 1–5 / 1–10. Treat as ordered, not interval, unless justified."
        rows.append([Paragraph(f"<b>{c}</b>",sCell), Paragraph(dtype,sCell), Paragraph(exs,sCell), Paragraph(str(miss),sCell), Paragraph(note,sCell)])
    t = Table(rows, colWidths=[52*mm,28*mm,42*mm,18*mm,40*mm], repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),NAVY), ("TEXTCOLOR",(0,0),(-1,0),white),
        ("GRID",(0,0),(-1,-1),0.4,HexColor("#B4C6E7")),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[white,GREY]),
        ("VALIGN",(0,0),(-1,-1),"TOP"), ("LEFTPADDING",(0,0),(-1,-1),3), ("RIGHTPADDING",(0,0),(-1,-1),3),
        ("TOPPADDING",(0,0),(-1,-1),2), ("BOTTOMPADDING",(0,0),(-1,-1),2),
    ]))
    story.append(t)
    story.append(Spacer(1,6))
    story.append(Paragraph("How to use this file: copy it to your working folder before editing. Never edit the CLEAN original. For Week 2, start from the DIRTY file and reconcile to this CLEAN truth. Quote totals as describing this sample.", sSmall))
    doc.build(story, onFirstPage=header_footer, onLaterPages=header_footer)
    print(f"  wrote {path.name} ({path.stat().st_size/1024:.0f} KB)")

print("=== PDF DICTIONARIES ===")
specs = [
    ("ecommerce_transactions_CLEAN.csv","ecommerce_transactions","Nigerian e-commerce (e-commerce.xlsx → Transactions)","one transaction line (transaction_id)","PK transaction_id; FKs customer_id, product_id, seller_id, delivery_id, order_id","Dropped 26,000 blank/partial rows (11,000 real kept). delivery_id standardised to DEL+6 digits to match Deliveries. Dates to ISO. Added order_year/month, is_paid/is_delivered. 570 Cancelled lines have no delivery (expected)."),
    ("ecommerce_customers_CLEAN.csv","ecommerce_customers","Nigerian e-commerce (Customers)","one customer (customer_id)","PK customer_id","Trimmed text; email lowercased; phone kept as 10-digit text; DOB/created to ISO; added age_at_2025. City/state pairing as supplied — validate before geographic analysis."),
    ("ecommerce_products_CLEAN.csv","ecommerce_products","Nigerian e-commerce (Products)","one product (product_id)","PK product_id; FK seller_id","Trimmed; dates to ISO; added discount_pct. unit_price in Transactions equals discount_price (verified)."),
    ("ecommerce_sellers_CLEAN.csv","ecommerce_sellers","Nigerian e-commerce (Sellers)","one seller store (seller_id)","PK seller_id","Trimmed; email lowercased; phone as text; dates to ISO."),
    ("ecommerce_deliveries_CLEAN.csv","ecommerce_deliveries","Nigerian e-commerce (Deliveries)","one delivery (delivery_id)","PK delivery_id; FK order_id, customer_id","Trimmed; dates to ISO (blank = Pending, 2,144 rows); added ship_lag_days (1–7 days)."),
    ("superstore_CLEAN.csv","superstore","US Superstore retail (Sample - Superstore.csv)","one order line (Row ID)","Order ID groups lines; Customer ID, Product ID are FKs","Latin-1 → UTF-8. Dates M/D/YYYY → ISO. Postal code as 5-char text. Added Ship_Lag_Days, Margin, Order_Year/Month. 1,871 loss-making lines kept (Margin<0)."),
    ("salary_CLEAN.csv","salary","Salary survey (Salary_Data.csv)","one person-job record","No ID — exact-duplicate grain","BOM fixed. Education standardised (Bachelor's→Bachelor's Degree etc.). Job_Family added. Exact duplicates removed (6,704→1,791). Outliers flagged, not deleted."),
    ("student_performance_CLEAN.csv","student_performance","Student performance factors","one student","No ID","Missing Teacher_Quality/Parental_Education/Distance → 'Unknown' (explicit). Ordinal _Score 1–3 added. Exam_Score>100 flagged."),
    ("hotel_reservations_CLEAN.csv","hotel_reservations","Hotel bookings (Hotel Reservations.csv)","one booking (Booking_ID)","PK Booking_ID","Arrival_Date_ISO from Y/M/D; 37 invalid (29 Feb 2018) flagged blank. Added Total_Nights/Guests, Is_Cancelled, quality flags."),
    ("healthcare_CLEAN.csv","healthcare","Hospital admissions (healthcare_dataset.csv)","one admission","No ID — admission grain","Names Title-Cased. Dates to ISO. Added Length_of_Stay_Days, Admission_Year/Month. Negative billing (108) flagged. Doctor/Hospital are synthetic names — do not treat as real staff."),
    ("healthcare_ANONYMISED.csv","healthcare_ANONYMISED","Hospital admissions — ANONYMISED (portfolio-safe)","one admission (Patient_ID)","Patient_ID pseudonym","Name removed; Patient_ID HLT000001… Only file that may leave the machine / enter AI tools."),
    ("diabetes_CLEAN.csv","diabetes","Diabetes screening (diabetes_prediction_dataset.csv)","one screening record","No ID","Smoking 'No Info' = missing (35.8%). Added age_group, bmi/hba1c/glucose categories, under-18 & BMI>60 flags. 100,000 rows; use 5,000 sample in Excel."),
    ("student_depression_CLEAN.csv","student_depression","Student wellbeing survey (SENSITIVE)","one respondent (id)","id (gapped)","City typos canonicalised; tiny professions grouped; Work Pressure/Job Satisfaction blank for Students (not applicable). Handle per NDPA 2023."),
    ("student_depression_ANONYMISED.csv","student_depression_ANONYMISED","Student wellbeing — ANONYMISED","one respondent (Respondent_ID)","Respondent_ID DEP000001…","id replaced. Only version for portfolios/AI."),
    ("social_media_addiction_CLEAN.csv","social_media_addiction","Social media & study (705 students)","one student (Student_ID)","PK Student_ID","Already tidy. Added Addiction_Band, Usage_Band."),
]
for csvf, name, src, grain, keys, clean_note in specs:
    df = pd.read_csv(CLEAN/csvf, low_memory=False)
    dict_table(df, name, src, grain, keys, clean_note)

# Master guide PDF
print("=== MASTER GUIDE ===")
path = PDF_DIR/"DATASET_GUIDE.pdf"
doc = SimpleDocTemplate(str(path), pagesize=A4, leftMargin=15*mm, rightMargin=15*mm, topMargin=18*mm, bottomMargin=15*mm,
                        title="ETEH DATASET_GUIDE", author="ETEH Ltd")
story=[]
story.append(Paragraph("ETEH Bootcamp — Dataset Guide", sTitle))
story.append(Paragraph("Read this before your first import. All nine attached datasets, standardised. No external/Wikipedia/PDF-bulletin data is used anywhere in the bootcamp.", sBody))
story.append(HRFlowable(width="100%", color=TEAL, thickness=1))
story.append(Paragraph("The nine datasets and when to open them", sH1))
rows=[[Paragraph("<b>Dataset (CLEAN file)</b>",sCellH),Paragraph("<b>Rows</b>",sCellH),Paragraph("<b>Use it when</b>",sCellH)]]
use_map={
 "ecommerce_transactions_CLEAN.csv":("11,000","Flagship for SQL joins (Week 3), formulas/pivots (Weeks 1–2), forecasting (Week 4). Nigerian ₦ context."),
 "ecommerce_customers_CLEAN.csv":("2,500","Joins, segmentation, loyalty analysis."),
 "ecommerce_products_CLEAN.csv":("800","Category/margin analysis; DAX measures."),
 "ecommerce_sellers_CLEAN.csv":("200","Seller scorecards; commission analysis."),
 "ecommerce_deliveries_CLEAN.csv":("10,430","Logistics: ship lag, courier, state analysis. LEFT JOIN teaching (570 Cancelled have no delivery)."),
 "superstore_CLEAN.csv":("9,994","Retail replaces fictional SuperMart. Pivots, Pareto, forecasting. US $ context."),
 "salary_CLEAN.csv":("1,791","Week 2 cleaning flagship (dirty version provided). Regression: Salary ~ Experience."),
 "student_performance_CLEAN.csv":("6,607","Education replaces fictional Crestview. Correlation, regression, dashboards."),
 "hotel_reservations_CLEAN.csv":("36,275","Cancellation analysis; date-building exercise (Y/M/D → ISO)."),
 "healthcare_CLEAN.csv":("55,500","Hospital ops replaces fictional Lifecare. LOS, billing, occupancy. Names are synthetic."),
 "healthcare_ANONYMISED.csv":("55,500","Only health file for portfolios/AI. Patient_ID only."),
 "diabetes_CLEAN.csv":("100,000","Clinical screening distributions; use 5,000 sample in Excel. 8.5% diabetes rate."),
 "student_depression_CLEAN.csv":("27,901","Sensitive wellbeing data. Ethics, structural zeros, small-group suppression. NDPA 2023."),
 "student_depression_ANONYMISED.csv":("27,901","Anonymised version for any sharing."),
 "social_media_addiction_CLEAN.csv":("705","Small clean file for first correlation/regression. Ideal Week 4 starter."),
}
for csvf,(n,use) in use_map.items():
    rows.append([Paragraph(f"<b>{csvf}</b>",sCell),Paragraph(n,sCell),Paragraph(use,sCell)])
t=Table(rows,colWidths=[62*mm,18*mm,100*mm],repeatRows=1)
t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),NAVY),("GRID",(0,0),(-1,-1),0.4,HexColor("#B4C6E7")),("ROWBACKGROUNDS",(0,1),(-1,-1),[white,GREY]),("VALIGN",(0,0),(-1,-1),"TOP")]))
story.append(t)
story.append(Paragraph("Three guarantees", sH1))
story.append(Paragraph("1) <b>Keys resolve.</b> Every e-commerce foreign key resolves except the 570 Cancelled transactions that correctly have no delivery — that is the LEFT JOIN lesson. &nbsp; 2) <b>Dates are ISO</b> (YYYY-MM-DD) in every CLEAN file. &nbsp; 3) <b>Missing is explicit:</b> blank means missing, 0 means zero, 'Unknown'/'No Info' means not recorded — never guess. Dirty files break these on purpose; the CLEAN file is the truth to reconcile to.", sBody))
story.append(Paragraph("Formats in this pack", sH1))
story.append(Paragraph("CSV (CLEAN + DIRTY + issues log) for Power Query/SQL import · XLSX extracts (multi-sheet, 5,000-row capped) for Excel work · SQLite (.sqlite3 + ecommerce_queries.sql) for Week 3 SQL · PDF dictionaries (this folder) for definitions. Copy raw files to your working folder; never edit CLEAN originals.", sBody))
story.append(Paragraph("Privacy", sH1))
story.append(Paragraph("Healthcare and depression files contain personal/health data. Use ANONYMISED versions for portfolios, presentations and any AI tool. Never paste names, emails or phone numbers into public AI. See Week 4 Module 4.13 and Week 5 Module P.8.", sBody))
doc.build(story, onFirstPage=header_footer, onLaterPages=header_footer)
print(f"  wrote {path.name} ({path.stat().st_size/1024:.0f} KB)")
print("PDF build complete.")
