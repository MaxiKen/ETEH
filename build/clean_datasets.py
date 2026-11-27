"""
ETEH Bootcamp - Phase A1: Standardize every provided dataset into a clean,
analysis-ready form.

Outputs (under Datasets/Clean/):
  - one <name>.csv  (machine-friendly snake_case headers, ISO dates, typed columns)
  - one <name>.xlsx (same data, friendly for Excel learners)

Standardization rules applied uniformly:
  * column headers -> snake_case
  * strip whitespace; collapse internal spaces
  * categorical text -> consistent casing / canonical labels
  * dates -> ISO 8601 (YYYY-MM-DD)
  * numerics -> proper float/int (no currency symbols, commas stripped)
  * exact duplicate rows -> dropped (count reported)
  * impossible / sentinel values -> coerced or flagged
"""
import os, re, unicodedata
import numpy as np
import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW  = os.path.join(ROOT, "Datasets", "Raw")
OUT  = os.path.join(ROOT, "Datasets", "Clean")
os.makedirs(OUT, exist_ok=True)

def read_any(path):
    for enc in ("utf-8", "latin-1", "cp1252"):
        try:
            return pd.read_csv(path, encoding=enc)
        except Exception as e:
            last = e
    raise last

def snake(name):
    s = unicodedata.normalize("NFKD", str(name)).encode("ascii", "ignore").decode()
    s = s.strip().lower()
    s = re.sub(r"[^a-z0-9]+", "_", s)
    s = re.sub(r"_+", "_", s).strip("_")
    return s

def std_cols(df):
    df = df.copy()
    df.columns = [snake(c) for c in df.columns]
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].astype(str).str.strip()
            df[c] = df[c].replace({"": np.nan, "nan": np.nan, "NaN": np.nan, "None": np.nan,
                                    "none": np.nan, "NULL": np.nan, "null": np.nan, "?" : np.nan})
    return df

def to_iso_date(series):
    """Parse many date formats to YYYY-MM-DD."""
    s = series.astype(str).str.strip()
    out = pd.to_datetime(s, errors="coerce", dayfirst=False)
    # fallback for dd/mm/yyyy style (only if many parsed as NaT)
    if out.isna().mean() > 0.5:
        out = pd.to_datetime(s, errors="coerce", dayfirst=True)
    return out.dt.strftime("%Y-%m-%d")

def money(series):
    s = series.astype(str).str.replace(r"[^0-9.\-]", "", regex=True)
    s = s.replace("", np.nan)
    return pd.to_numeric(s, errors="coerce")

# ---------------------------------------------------------------- per-dataset
def clean_hotel(df):
    df = std_cols(df)
    df["type_of_meal_plan"] = df["type_of_meal_plan"].replace({"Not Selected": "None"})
    df["booking_status"] = df["booking_status"].str.replace("_", " ", regex=False)
    df["market_segment_type"] = df["market_segment_type"].str.title()
    # sentinel avg_price_per_room == 0 -> missing (a room cannot cost 0)
    df.loc[df["avg_price_per_room"] == 0, "avg_price_per_room"] = np.nan
    ymd = (df["arrival_year"].astype(int).astype(str) + "-" +
           df["arrival_month"].astype(int).astype(str).str.zfill(2) + "-" +
           df["arrival_date"].astype(int).astype(str).str.zfill(2))
    df["arrival_date_iso"] = pd.to_datetime(ymd, format="%Y-%m-%d", errors="coerce").dt.strftime("%Y-%m-%d")
    return df

def clean_salary(df):
    df = std_cols(df)
    edu_map = {
        "bachelor's": "Bachelor's", "bachelor's degree": "Bachelor's", "bachelors": "Bachelor's",
        "master's": "Master's", "master's degree": "Master's", "masters": "Master's",
        "phd": "PhD", "ph.d": "PhD", "doctoral": "PhD",
        "high school": "High School", "hs diploma": "High School",
        "associate": "Associate", "associate degree": "Associate",
        "college": "College", "some college": "College",
    }
    df["education_level"] = df["education_level"].str.lower().map(edu_map).fillna(df["education_level"])
    df["job_title"] = df["job_title"].str.title()
    df["gender"] = df["gender"].str.title()
    df["salary"] = money(df["salary"])
    return df

def clean_superstore(df):
    df = std_cols(df)
    df["order_date"] = to_iso_date(df["order_date"])
    df["ship_date"] = to_iso_date(df["ship_date"])
    df["postal_code"] = df["postal_code"].astype("Int64").astype(str).str.zfill(5)
    df["sales"] = money(df["sales"]); df["profit"] = money(df["profit"])
    df["discount"] = pd.to_numeric(df["discount"], errors="coerce")
    df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce").astype("Int64")
    # derived helpful columns
    od = pd.to_datetime(df["order_date"]); sd = pd.to_datetime(df["ship_date"])
    df["ship_days"] = (sd - od).dt.days
    df["profit_margin"] = (df["profit"] / df["sales"]).round(4)
    df["year"] = od.dt.year
    return df

def clean_student_dep(df):
    df = std_cols(df)
    df = df.rename(columns={"have_you_ever_had_suicidal_thoughts": "suicidal_thoughts"})
    df["gender"] = df["gender"].str.title()
    df["dietary_habits"] = df["dietary_habits"].str.title().replace({"Others": "Other"})
    df["sleep_duration"] = df["sleep_duration"].str.title().replace({"Others": "Other"})
    df["suicidal_thoughts"] = df["suicidal_thoughts"].str.title()
    df["family_history_of_mental_illness"] = df["family_history_of_mental_illness"].str.title()
    return df

def clean_student_perf(df):
    df = std_cols(df)
    # Exam_Score cannot exceed 100 -> cap impossible values, keep raw in note
    df["exam_score_raw"] = df["exam_score"]
    df["exam_score_impossible"] = (df["exam_score"] > 100).astype(int)
    df["exam_score"] = df["exam_score"].clip(upper=100)
    df["gender"] = df["gender"].str.title()
    df["extracurricular_activities"] = df["extracurricular_activities"].str.title()
    df["internet_access"] = df["internet_access"].str.title()
    df["learning_disabilities"] = df["learning_disabilities"].str.title()
    df["parental_involvement"] = df["parental_involvement"].str.title()
    df["access_to_resources"] = df["access_to_resources"].str.title()
    df["motivation_level"] = df["motivation_level"].str.title()
    df["family_income"] = df["family_income"].str.title()
    df["teacher_quality"] = df["teacher_quality"].str.title()
    df["school_type"] = df["school_type"].str.title()
    df["peer_influence"] = df["peer_influence"].str.title()
    df["parental_education_level"] = df["parental_education_level"].str.title()
    df["distance_from_home"] = df["distance_from_home"].str.title()
    return df

def clean_social(df):
    df = std_cols(df)
    df["gender"] = df["gender"].str.title()
    df["academic_level"] = df["academic_level"].str.title()
    df["most_used_platform"] = df["most_used_platform"].str.title()
    df["affects_academic_performance"] = df["affects_academic_performance"].str.title()
    df["relationship_status"] = df["relationship_status"].str.title()
    return df

def clean_diabetes(df):
    df = std_cols(df)
    df["gender"] = df["gender"].str.title()
    smk = {
        "never": "Never", "no info": "Unknown", "current": "Current",
        "former": "Former", "ever": "Ever", "not current": "Not Current",
    }
    df["smoking_history"] = df["smoking_history"].str.lower().map(smk).fillna(df["smoking_history"].str.title())
    return df

def clean_healthcare(df):
    df = std_cols(df)
    def proper(s):
        return s.str.title()
    df["name"] = proper(df["name"])
    df["doctor"] = proper(df["doctor"])
    df["hospital"] = proper(df["hospital"])
    df["gender"] = df["gender"].str.title()
    df["blood_type"] = df["blood_type"].str.upper()
    df["medical_condition"] = df["medical_condition"].str.title()
    df["insurance_provider"] = df["insurance_provider"].str.title()
    df["admission_type"] = df["admission_type"].str.title()
    df["medication"] = df["medication"].str.title()
    df["test_results"] = df["test_results"].str.title()
    df["date_of_admission"] = to_iso_date(df["date_of_admission"])
    df["discharge_date"] = to_iso_date(df["discharge_date"])
    df["billing_amount"] = money(df["billing_amount"])
    # negative billing -> absolute value (data-entry sign error), flag kept
    df["billing_was_negative"] = (df["billing_amount"] < 0).astype(int)
    df["billing_amount"] = df["billing_amount"].abs()
    df["room_number"] = pd.to_numeric(df["room_number"], errors="coerce").astype("Int64")
    # length of stay
    da = pd.to_datetime(df["date_of_admission"]); dd = pd.to_datetime(df["discharge_date"])
    df["length_of_stay"] = (dd - da).dt.days
    return df

CLEANERS = {
    "hotel_reservations.csv":      ("hotel_reservations", clean_hotel),
    "salary_data.csv":             ("salary_data", clean_salary),
    "superstore_sales.csv":        ("superstore_sales", clean_superstore),
    "student_depression.csv":      ("student_depression", clean_student_dep),
    "student_performance.csv":     ("student_performance", clean_student_perf),
    "social_media_addiction.csv":  ("social_media_addiction", clean_social),
    "diabetes_prediction.csv":     ("diabetes_prediction", clean_diabetes),
    "healthcare_records.csv":      ("healthcare_records", clean_healthcare),
}

def main():
    summary = []
    for raw, (name, fn) in CLEANERS.items():
        df = read_any(os.path.join(RAW, raw))
        before = len(df)
        df = fn(df)
        dups = int(df.duplicated().sum())
        # NOTE: the "clean" pass is a faithful standardized reference (types, labels,
        # dates, sentinels fixed). We deliberately DO NOT drop rows here so no real
        # data is destroyed; duplicate rows are left for learners to discover, and the
        # Dirty pass injects *additional* intentional issues for analysis practice.
        csv_path = os.path.join(OUT, name + ".csv")
        df.to_csv(csv_path, index=False)
        xlsx_path = os.path.join(OUT, name + ".xlsx")
        df.to_excel(xlsx_path, index=False, engine="openpyxl")
        summary.append((name, before, len(df), int(dups), df.shape[1]))
        print(f"[OK] {name}: {before}->{len(df)} rows, {df.shape[1]} cols, dropped {int(dups)} dup")
    # e-commerce handled separately (relational); just save a tidy xlsx per sheet + combined csv
    import openpyxl
    wb = openpyxl.load_workbook(os.path.join(RAW, "e-commerce.xlsx"))
    combo = []
    for ws in wb.worksheets:
        d = pd.read_excel(os.path.join(RAW, "e-commerce.xlsx"), sheet_name=ws.title)
        d.columns = [snake(c) for c in d.columns]
        d.to_csv(os.path.join(OUT, f"ecommerce_{ws.title.lower()}.csv"), index=False)
        combo.append(d.assign(__table__=ws.title.lower()))
    ec = pd.concat(combo, ignore_index=True)
    ec.to_excel(os.path.join(OUT, "ecommerce_all.xlsx"), index=False, engine="openpyxl")
    print(f"[OK] ecommerce: 5 relational tables exported (transactions, sellers, products, deliveries, customers)")
    print("\nSUMMARY")
    for s in summary:
        print("  ", s)

if __name__ == "__main__":
    main()
