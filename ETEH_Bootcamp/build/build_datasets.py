#!/usr/bin/env python3
"""
ETEH Bootcamp - Dataset build pipeline (Part 1 of 3)
Standardises the supplied raw datasets, produces a CLEAN (gold) copy and a
deliberately DIRTIED working copy for students to clean, and exports them as
CSV + XLSX. The relational e-commerce set is handled in build_sql.py.

Run:  python3 build_datasets.py
"""
import os, re, random, string, datetime as dt
import numpy as np
import pandas as pd

ROOT = "/home/user/ETEH"
SRC  = ROOT                      # raw supplied files live at repo root
OUT  = os.path.join(ROOT, "ETEH_Bootcamp", "datasets")
CLEAN = os.path.join(OUT, "clean")
DIRTY = os.path.join(OUT, "dirty")
os.makedirs(CLEAN, exist_ok=True)
os.makedirs(DIRTY, exist_ok=True)

random.seed(20240517)
np.random.seed(20240517)

# ----------------------------------------------------------------------------
#  Dirtying helpers
# ----------------------------------------------------------------------------
def ws(series, frac=0.15):
    """Add leading/trailing whitespace to a fraction of string cells."""
    def f(v):
        if isinstance(v, str) and v != "" and random.random() < frac:
            return random.choice(["  ", " "]) + v + random.choice(["  ", " "])
        return v
    return series.apply(f)

def jitter_case(series, frac=0.12):
    """Randomly lower/upper-case whole string values."""
    def f(v):
        if isinstance(v, str) and v != "" and random.random() < frac:
            return random.choice([v.lower(), v.upper()])
        return v
    return series.apply(f)

def blanks(df, cols, frac=0.04):
    for c in cols:
        mask = df[c].notna() & (np.random.rand(len(df)) < frac)
        df.loc[mask, c] = np.nan
    return df

def dups(df, n):
    if n <= 0:
        return df
    idx = np.random.choice(df.index, size=n, replace=False)
    return pd.concat([df, df.loc[idx]], ignore_index=True)

def mixed_dates(series, frac=0.4):
    """Keep some as ISO, scramble others into US / textual formats."""
    out = []
    fmts = ["%m/%d/%Y", "%d-%m-%Y", "%Y/%m/%d", "%b %d, %Y", "%d %b %Y"]
    for v in series:
        if pd.isna(v):
            out.append(v); continue
        try:
            d = pd.to_datetime(v)
        except Exception:
            out.append(v); continue
        if random.random() < frac:
            out.append(d.strftime(random.choice(fmts)))
        else:
            out.append(d.strftime("%Y-%m-%d"))
    return pd.Series(out, index=series.index)

def money_text(series, frac=0.25, sym="₦"):
    out = []
    for v in series:
        if pd.isna(v):
            out.append(v); continue
        if random.random() < frac:
            out.append(f"{sym}{v:,.2f}")
        else:
            out.append(v)
    return pd.Series(out, index=series.index)

def title_proper(s):
    if not isinstance(s, str): return s
    return " ".join(w[:1].upper() + w[1:].lower() if w else w for w in s.split())

# ============================================================================
#  1. ETEH RETAIL  (from Sample - Superstore.csv)
# ============================================================================
def build_retail():
    df = pd.read_csv(os.path.join(SRC, "Sample - Superstore.csv"), encoding="cp1252")
    ren = {'Row ID':'row_id','Order ID':'order_id','Order Date':'order_date',
           'Ship Date':'ship_date','Ship Mode':'ship_mode','Customer ID':'customer_id',
           'Customer Name':'customer_name','Segment':'segment','Country':'country',
           'City':'city','State':'state','Postal Code':'postal_code','Region':'region',
           'Product ID':'product_id','Category':'category','Sub-Category':'sub_category',
           'Product Name':'product_name','Sales':'sales','Quantity':'quantity',
           'Discount':'discount','Profit':'profit'}
    df = df.rename(columns=ren)
    for c in ['order_date','ship_date']:
        df[c] = pd.to_datetime(df[c], format='%m/%d/%Y', errors='coerce')
    df['ship_days'] = (df['ship_date'] - df['order_date']).dt.days
    df = df.sort_values('order_date').reset_index(drop=True)
    # CLEAN
    clean = df.copy()
    for c in ['category','sub_category','segment','region','ship_mode','state','country']:
        clean[c] = clean[c].astype(str).str.strip().str.title()
    clean['order_date'] = clean['order_date'].dt.strftime('%Y-%m-%d')
    clean['ship_date']  = clean['ship_date'].dt.strftime('%Y-%m-%d')
    clean.to_csv(os.path.join(CLEAN, "eteh_retail.csv"), index=False)
    clean.to_excel(os.path.join(CLEAN, "eteh_retail.xlsx"), index=False)
    # DIRTY
    dirty = df.copy()
    dirty['category']    = jitter_case(ws(dirty['category'], 0.2), 0.15)
    dirty['sub_category']= jitter_case(ws(dirty['sub_category'], 0.2), 0.12)
    dirty['ship_mode']   = dirty['ship_mode'].replace({'Standard Class':'Stan dard Class','Same Day':'SameDay'})
    dirty['region']      = ws(dirty['region'], 0.15)
    dirty['state']       = dirty['state'].apply(lambda v: f"  {v} " if random.random()<0.1 else v)
    dirty['order_date']  = mixed_dates(dirty['order_date'])
    dirty['ship_date']   = mixed_dates(dirty['ship_date'])
    dirty = blanks(dirty, ['discount','profit'], 0.03)
    dirty = dups(dirty, 60)
    dirty.to_csv(os.path.join(DIRTY, "eteh_retail.csv"), index=False)
    dirty.to_excel(os.path.join(DIRTY, "eteh_retail.xlsx"), index=False)
    return len(clean)

# ============================================================================
#  2. ETEH LIFECARE HOSPITAL  (from healthcare_dataset.csv)
# ============================================================================
def build_lifecare():
    df = pd.read_csv(os.path.join(SRC, "healthcare_dataset.csv"))
    ren = {'Name':'patient_name','Age':'age','Gender':'gender','Blood Type':'blood_type',
           'Medical Condition':'medical_condition','Date of Admission':'admission_date',
           'Doctor':'doctor','Hospital':'hospital','Insurance Provider':'insurance_provider',
           'Billing Amount':'billing_amount','Room Number':'room_number','Admission Type':'admission_type',
           'Discharge Date':'discharge_date','Medication':'medication','Test Results':'test_results'}
    df = df.rename(columns=ren)
    for c in ['admission_date','discharge_date']:
        df[c] = pd.to_datetime(df[c], errors='coerce')
    df['length_of_stay'] = (df['discharge_date'] - df['admission_date']).dt.days
    cat_cols = ['medical_condition','admission_type','test_results','medication','insurance_provider','gender','blood_type']
    # CLEAN
    clean = df.copy()
    for c in cat_cols:
        clean[c] = clean[c].astype(str).str.strip().str.title()
    clean['patient_name'] = clean['patient_name'].apply(title_proper)
    clean['doctor'] = clean['doctor'].apply(title_proper)
    clean['hospital'] = clean['hospital'].apply(title_proper)
    clean['admission_date'] = clean['admission_date'].dt.strftime('%Y-%m-%d')
    clean['discharge_date'] = clean['discharge_date'].dt.strftime('%Y-%m-%d')
    clean = clean[['patient_name','age','gender','blood_type','medical_condition','admission_date',
                  'discharge_date','length_of_stay','doctor','hospital','insurance_provider',
                  'billing_amount','room_number','admission_type','medication','test_results']]
    clean.to_csv(os.path.join(CLEAN, "eteh_lifecare.csv"), index=False)
    clean.to_excel(os.path.join(CLEAN, "eteh_lifecare.xlsx"), index=False)
    # DIRTY
    dirty = df.copy()
    dirty['patient_name'] = dirty['patient_name'].apply(lambda v: v if random.random()<0.5 else v.swapcase())
    dirty['medical_condition'] = jitter_case(dirty['medical_condition'], 0.2)
    dirty['admission_type'] = jitter_case(ws(dirty['admission_type'],0.15),0.15)
    dirty['test_results'] = jitter_case(dirty['test_results'], 0.18)
    dirty['admission_date'] = mixed_dates(dirty['admission_date'], 0.55)
    dirty['discharge_date'] = mixed_dates(dirty['discharge_date'], 0.55)
    dirty['billing_amount'] = money_text(dirty['billing_amount'], 0.35)
    dirty = blanks(dirty, ['test_results','insurance_provider'], 0.04)
    dirty = dups(dirty, 50)
    dirty = dirty[['patient_name','age','gender','blood_type','medical_condition','admission_date',
                   'discharge_date','length_of_stay','doctor','hospital','insurance_provider',
                   'billing_amount','room_number','admission_type','medication','test_results']]
    dirty.to_csv(os.path.join(DIRTY, "eteh_lifecare.csv"), index=False)
    dirty.to_excel(os.path.join(DIRTY, "eteh_lifecare.xlsx"), index=False)
    return len(clean)

# ============================================================================
#  3. ETEH UNIVERSITY - STUDENT ANALYTICS  (3 related student files -> 1 workbook)
# ============================================================================
def build_university():
    # 3a Performance
    pf = pd.read_csv(os.path.join(SRC, "StudentPerformanceFactors.csv"))
    pf = pf.rename(columns={c:c.lower() for c in pf.columns})
    pf_clean = pf.copy()
    for c in ['parental_involvement','access_to_resources','motivation_level','family_income',
              'teacher_quality','school_type','peer_influence','learning_disabilities',
              'parental_education_level','distance_from_home','gender',
              'extracurricular_activities','internet_access']:
        if c in pf_clean: pf_clean[c] = pf_clean[c].astype(str).str.strip().str.title()
    # 3b Wellbeing / Depression
    dp = pd.read_csv(os.path.join(SRC, "Student Depression Dataset.csv"))
    dp = dp.rename(columns={c:c.lower().replace(' ','_').replace('?','').strip('_') for c in dp.columns})
    dp_clean = dp.copy()
    for c in ['gender','city','profession','dietary_habits','degree',
              'have_you_ever_had_suicidal_thoughts','family_history_of_mental_illness']:
        if c in dp_clean: dp_clean[c] = dp_clean[c].astype(str).str.strip().str.title()
    # 3c Social media
    sm = pd.read_csv(os.path.join(SRC, "Students Social Media Addiction.csv"))
    sm = sm.rename(columns={c:c.lower().replace(' ','_') for c in sm.columns})
    sm_clean = sm.copy()
    for c in ['gender','academic_level','country','most_used_platform',
              'affects_academic_performance','relationship_status']:
        if c in sm_clean: sm_clean[c] = sm_clean[c].astype(str).str.strip().str.title()

    # write clean workbook (3 sheets)
    with pd.ExcelWriter(os.path.join(CLEAN, "eteh_university.xlsx"), engine='openpyxl') as xw:
        pf_clean.to_excel(xw, sheet_name='Performance', index=False)
        dp_clean.to_excel(xw, sheet_name='Wellbeing', index=False)
        sm_clean.to_excel(xw, sheet_name='Social_Media', index=False)
    # also individual csvs
    pf_clean.to_csv(os.path.join(CLEAN, "eteh_uni_performance.csv"), index=False)
    dp_clean.to_csv(os.path.join(CLEAN, "eteh_uni_wellbeing.csv"), index=False)
    sm_clean.to_csv(os.path.join(CLEAN, "eteh_uni_social.csv"), index=False)

    # DIRTY workbook
    pf_d = pf.copy()
    pf_d['gender'] = jitter_case(ws(pf_d['gender'],0.1),0.1)
    pf_d['parental_involvement'] = jitter_case(pf_d['parental_involvement'],0.12)
    pf_d = blanks(pf_d, ['attendance','previous_scores'], 0.03)
    pf_d = dups(pf_d, 40)
    dp_d = dp.copy()
    dp_d['dietary_habits'] = jitter_case(dp_d['dietary_habits'],0.15)
    dp_d['sleep_duration'] = dp_d['sleep_duration'].replace({'5-6 hours':'5 – 6 hrs','7-8 hours':'7 to 8 hours'})
    dp_d = blanks(dp_d, ['financial_stress'], 0.03)
    sm_d = sm.copy()
    sm_d['most_used_platform'] = jitter_case(sm_d['most_used_platform'],0.12)
    sm_d = dups(sm_d, 25)
    with pd.ExcelWriter(os.path.join(DIRTY, "eteh_university.xlsx"), engine='openpyxl') as xw:
        pf_d.to_excel(xw, sheet_name='Performance', index=False)
        dp_d.to_excel(xw, sheet_name='Wellbeing', index=False)
        sm_d.to_excel(xw, sheet_name='Social_Media', index=False)
    pf_d.to_csv(os.path.join(DIRTY, "eteh_uni_performance.csv"), index=False)
    dp_d.to_csv(os.path.join(DIRTY, "eteh_uni_wellbeing.csv"), index=False)
    sm_d.to_csv(os.path.join(DIRTY, "eteh_uni_social.csv"), index=False)
    return len(pf_clean)

# ============================================================================
#  4. ETEH HOTELS  (from Hotel Reservations.csv)
# ============================================================================
def build_hotels():
    df = pd.read_csv(os.path.join(SRC, "Hotel Reservations.csv"))
    ren = {'Booking_ID':'booking_id','no_of_adults':'num_adults','no_of_children':'num_children',
           'no_of_weekend_nights':'weekend_nights','no_of_week_nights':'week_nights',
           'type_of_meal_plan':'meal_plan','required_car_parking_space':'car_parking',
           'room_type_reserved':'room_type','lead_time':'lead_time','arrival_year':'arrival_year',
           'arrival_month':'arrival_month','arrival_date':'arrival_day','market_segment_type':'market_segment',
           'repeated_guest':'repeated_guest','no_of_previous_cancellations':'prev_cancellations',
           'no_of_previous_bookings_not_canceled':'prev_not_canceled','avg_price_per_room':'avg_price_per_room',
           'no_of_special_requests':'special_requests','booking_status':'booking_status'}
    df = df.rename(columns=ren)
    ymd = (df['arrival_year'].astype(int).astype(str) + '-'
           + df['arrival_month'].astype(int).astype(str).str.zfill(2) + '-'
           + df['arrival_day'].astype(int).astype(str).str.zfill(2))
    df['arrival_date'] = pd.to_datetime(ymd, errors='coerce').dt.strftime('%Y-%m-%d')
    df = df.drop(columns=['arrival_year','arrival_month','arrival_day'])
    cat = ['meal_plan','room_type','market_segment','booking_status']
    clean = df.copy()
    for c in cat:
        clean[c] = clean[c].astype(str).str.strip().str.replace('_',' ').str.title()
    clean['booking_status'] = clean['booking_status'].str.replace('Not Canceled','Not Cancelled')
    clean.to_csv(os.path.join(CLEAN, "eteh_hotels.csv"), index=False)
    clean.to_excel(os.path.join(CLEAN, "eteh_hotels.xlsx"), index=False)
    # DIRTY
    dirty = df.copy()
    dirty['meal_plan'] = jitter_case(ws(dirty['meal_plan'],0.15),0.2).str.replace('Meal Plan','meal plan',regex=False)
    dirty['room_type'] = dirty['room_type'].str.replace('Room_Type','Room Type',regex=False)
    dirty['room_type'] = jitter_case(dirty['room_type'],0.15)
    dirty['booking_status'] = dirty['booking_status'].apply(lambda v: v.lower() if random.random()<0.15 else v)
    dirty['market_segment'] = jitter_case(dirty['market_segment'],0.12)
    dirty = blanks(dirty, ['avg_price_per_room','special_requests'], 0.03)
    dirty = dups(dirty, 45)
    dirty.to_csv(os.path.join(DIRTY, "eteh_hotels.csv"), index=False)
    dirty.to_excel(os.path.join(DIRTY, "eteh_hotels.xlsx"), index=False)
    return len(clean)

# ============================================================================
#  5. ETEH SALARY  (from Salary_Data.csv) - simple regression example
# ============================================================================
def build_salary():
    df = pd.read_csv(os.path.join(SRC, "Salary_Data.csv"))
    ren = {'Age':'age','Gender':'gender','Education Level':'education_level','Job Title':'job_title',
           'Years of Experience':'years_experience','Salary':'salary'}
    df = df.rename(columns=ren)
    edu_map = {"Bachelor's Degree":"Bachelor's","Bachelor's":"Bachelor's","Master's Degree":"Master's",
               "Master's":"Master's","phD":"PhD","PhD":"PhD","High School":"High School"}
    clean = df.copy()
    clean['education_level'] = clean['education_level'].map(edu_map).fillna(clean['education_level'])
    clean['gender'] = clean['gender'].str.strip().str.title()
    clean['job_title'] = clean['job_title'].str.strip().str.title()
    for c in ['age','years_experience','salary']:
        clean[c] = pd.to_numeric(clean[c], errors='coerce')
    clean = clean.dropna(subset=['salary']).reset_index(drop=True)
    clean.to_csv(os.path.join(CLEAN, "eteh_salary.csv"), index=False)
    clean.to_excel(os.path.join(CLEAN, "eteh_salary.xlsx"), index=False)
    # DIRTY = keep the original inconsistencies (the teaching point)
    dirty = df.copy()
    dirty['education_level'] = dirty['education_level'].fillna("Unknown")
    dirty = dups(dirty, 30)
    dirty.to_csv(os.path.join(DIRTY, "eteh_salary.csv"), index=False)
    dirty.to_excel(os.path.join(DIRTY, "eteh_salary.xlsx"), index=False)
    return len(clean)

# ============================================================================
#  6. ETEH HEALTH (DIABETES)  (from diabetes_prediction_dataset.csv) - 10k sample
# ============================================================================
def build_diabetes():
    df = pd.read_csv(os.path.join(SRC, "diabetes_prediction_dataset.csv"))
    ren = {'gender':'gender','age':'age','hypertension':'hypertension','heart_disease':'heart_disease',
           'smoking_history':'smoking_history','bmi':'bmi','HbA1c_level':'hba1c_level',
           'blood_glucose_level':'blood_glucose_level','diabetes':'diabetes'}
    df = df.rename(columns=ren)
    sample = df.sample(n=10000, random_state=7).reset_index(drop=True)
    clean = sample.copy()
    clean['gender'] = clean['gender'].str.strip().str.title()
    clean['smoking_history'] = clean['smoking_history'].str.strip().str.title().replace({'No Info':'No Info'})
    clean.to_csv(os.path.join(CLEAN, "eteh_health_diabetes.csv"), index=False)
    clean.to_excel(os.path.join(CLEAN, "eteh_health_diabetes.xlsx"), index=False)
    dirty = sample.copy()
    dirty['smoking_history'] = jitter_case(dirty['smoking_history'],0.15)
    dirty['gender'] = dirty['gender'].apply(lambda v: v.lower() if random.random()<0.1 else v)
    dirty = dups(dirty, 35)
    dirty.to_csv(os.path.join(DIRTY, "eteh_health_diabetes.csv"), index=False)
    dirty.to_excel(os.path.join(DIRTY, "eteh_health_diabetes.xlsx"), index=False)
    return len(clean)

if __name__ == "__main__":
    print("Retail rows      :", build_retail())
    print("Lifecare rows    :", build_lifecare())
    print("University rows  :", build_university())
    print("Hotels rows      :", build_hotels())
    print("Salary rows      :", build_salary())
    print("Diabetes sample  :", build_diabetes())
    print("DONE part 1")
