"""
ETEH Bootcamp - Phase A2: Produce intentionally DIRTY copies of selected datasets
so learners can practice cleaning / QA / analysis.

Each dirty file is built from the corresponding Clean file and has a documented set
of issues (see Datasets/dirty_manifest.md). Seeded with random.seed(42) for
reproducibility. Issues are realistic: missing values, duplicate rows, mixed-case
labels, whitespace, currency-formatted numbers, mixed date formats, impossible
values and inconsistent category spellings.
"""
import os, re, random, string
import numpy as np
import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CLEAN = os.path.join(ROOT, "Datasets", "Clean")
DIRTY = os.path.join(ROOT, "Datasets", "Dirty")
os.makedirs(DIRTY, exist_ok=True)
random.seed(42)
np.random.seed(42)

def add_spaces(s):  # occasional leading/trailing whitespace + double spaces
    if random.random() < 0.06:
        s = "  " + s.strip() + " " + random.choice(string.ascii_lowercase) + " "
    return s

def mess_case(s):  # randomise categorical casing
    r = random.random()
    if r < 0.10: return s.lower()
    if r < 0.18: return s.upper()
    return s

def add_missing(df, cols, frac=0.05):
    for c in cols:
        idx = df.sample(frac=frac, random_state=random.randint(0, 1_000_000)).index
        df.loc[idx, c] = np.nan
    return df

def add_duplicates(df, frac=0.03):
    dups = df.sample(frac=frac, random_state=random.randint(0, 1_000_000))
    return pd.concat([df, dups], ignore_index=True)

def currency_format(series):
    s = pd.to_numeric(series, errors="coerce")
    out = s.astype("object")
    mask = s.notna() & (np.random.rand(len(s)) < 0.15)
    out[mask] = ["$" + f"{v:,.2f}" for v in s[mask]]
    return out

def as_obj(df, col):
    """Safely convert a column to object dtype so string values can be injected."""
    df[col] = df[col].astype(object)
    return df

def mixed_dates(series, formats):
    """Re-format a YYYY-MM-DD date column into random mixed formats."""
    dt = pd.to_datetime(series, errors="coerce")
    out = series.astype("object")
    for i in dt.index:
        d = dt.loc[i]
        if pd.isna(d):
            continue
        if np.random.rand() < 0.30:
            fmt = random.choice(formats)
            out.loc[i] = d.strftime(fmt)
    return out

MANIFEST = []

def record(name, issues):
    MANIFEST.append((name, issues))

def dirty_healthcare(df):
    issues = []
    df["name"] = df["name"].apply(lambda x: mess_case(str(x)) if random.random()<0.12 else x)
    issues.append("Mixed-case patient names (e.g. 'bObby jacksOn', 'LESILE TERRY').")
    df = add_missing(df, ["blood_type", "medical_condition", "insurance_provider", "doctor"], 0.04)
    issues.append("~4% missing values injected in blood_type, medical_condition, insurance_provider, doctor.")
    df = as_obj(df, "billing_amount")
    df["billing_amount"] = currency_format(df["billing_amount"])
    issues.append("~15% of billing_amount values formatted as currency text ('$12,345.67').")
    df["date_of_admission"] = mixed_dates(df["date_of_admission"], ["%d/%m/%Y", "%b %d, %Y", "%Y/%m/%d"])
    issues.append("~30% of date_of_admission rewritten in mixed formats (DD/MM/YYYY, 'Mon DD, YYYY', YYYY/MM/DD).")
    df["gender"] = df["gender"].apply(lambda x: random.choice(["male","Male","MALE","Female","female"]) if random.random()<0.08 else x)
    issues.append("Inconsistent gender labels (male/Male/MALE).")
    df = add_duplicates(df, 0.025)
    issues.append("~2.5% exact duplicate rows appended.")
    # reintroduce a few impossible negative bills
    neg = df.sample(frac=0.005, random_state=7).index
    df.loc[neg, "billing_amount"] = ["-" + str(v) for v in df.loc[neg, "billing_amount"]]
    issues.append("~0.5% billing_amount values turned negative (sign-error).")
    return df, issues

def dirty_salary(df):
    issues = []
    df["education_level"] = df["education_level"].apply(
        lambda x: random.choice(["bachelors","Bachelors","Bachelor's ","masters","Masters","PHD","Phd"]) if random.random()<0.15 else x)
    issues.append("Inconsistent education_level spellings/casing (bachelors, Bachelor's , PHD, masters...).")
    df = as_obj(df, "salary")
    df["salary"] = currency_format(df["salary"])
    issues.append("~15% of salary values formatted as currency text ('$95,000.00').")
    df["gender"] = df["gender"].apply(lambda x: random.choice(["M","F","Male","female","Female"]) if random.random()<0.10 else x)
    issues.append("Mixed gender labels (M, F, Male, female).")
    df["job_title"] = df["job_title"].apply(lambda x: add_spaces(str(x)))
    issues.append("Random leading/trailing whitespace in job_title.")
    df = add_missing(df, ["age", "years_of_experience"], 0.03)
    issues.append("~3% missing values in age and years_of_experience.")
    df = add_duplicates(df, 0.02)
    issues.append("~2% exact duplicate rows appended.")
    return df, issues

def dirty_superstore(df):
    issues = []
    df["ship_date"] = mixed_dates(df["ship_date"], ["%d-%b-%Y", "%m/%d/%Y", "%Y%m%d"])
    issues.append("~30% of ship_date rewritten in mixed formats (DD-Mon-YYYY, MM/DD/YYYY, YYYYMMDD).")
    df = add_missing(df, ["postal_code", "region"], 0.04)
    issues.append("~4% missing values in postal_code and region.")
    df["city"] = df["city"].apply(lambda x: add_spaces(str(x)))
    issues.append("Random whitespace added to city names.")
    df = as_obj(df, "sales")
    df["sales"] = currency_format(df["sales"])
    issues.append("~15% of sales formatted as currency text.")
    # reintroduce some negative profit (returns)
    df = as_obj(df, "profit")
    neg = df.sample(frac=0.01, random_state=3).index
    df.loc[neg, "profit"] = ["-" + str(v) for v in df.loc[neg, "profit"]]
    issues.append("~1% profit values turned negative (returned orders).")
    df = add_duplicates(df, 0.02)
    issues.append("~2% exact duplicate rows appended.")
    return df, issues

def dirty_student_perf(df):
    issues = []
    # reintroduce impossible exam scores > 100
    idx = df.sample(frac=0.01, random_state=9).index
    df.loc[idx, "exam_score"] = (df.loc[idx, "exam_score"].astype(float) + 5).round(1)
    issues.append("~1% exam_score values pushed above 100 (impossible).")
    df["gender"] = df["gender"].apply(lambda x: random.choice(["male","Male","M","Female","female"]) if random.random()<0.08 else x)
    issues.append("Inconsistent gender labels.")
    df["extracurricular_activities"] = df["extracurricular_activities"].apply(
        lambda x: random.choice(["yes","Yes","Y","no","No"]) if random.random()<0.08 else x)
    issues.append("Inconsistent Yes/No labels in extracurricular_activities.")
    df = add_missing(df, ["hours_studied", "sleep_hours", "previous_scores"], 0.04)
    issues.append("~4% missing in hours_studied, sleep_hours, previous_scores.")
    df = add_duplicates(df, 0.02)
    issues.append("~2% duplicate rows appended.")
    return df, issues

def dirty_diabetes(df):
    issues = []
    df["smoking_history"] = df["smoking_history"].apply(
        lambda x: random.choice(["No Info","no info","NOT CURRENT","Ever","never","CURRENT"]) if random.random()<0.15 else x)
    issues.append("Inconsistent smoking_history labels (No Info, no info, NOT CURRENT, Ever...).")
    df = add_missing(df, ["bmi", "hbA1c_level"], 0.04)
    issues.append("~4% missing in bmi and hbA1c_level.")
    df["gender"] = df["gender"].apply(lambda x: random.choice(["male","Male","Female","female"]) if random.random()<0.06 else x)
    issues.append("Mixed gender casing.")
    df = add_duplicates(df, 0.015)
    issues.append("~1.5% duplicate rows appended.")
    return df, issues

def dirty_hotel(df):
    issues = []
    df["booking_status"] = df["booking_status"].apply(
        lambda x: random.choice(["canceled","Cancelled","Canceled","not_canceled","Not_Canceled"]) if random.random()<0.08 else x)
    issues.append("Inconsistent booking_status labels (canceled/Cancelled/not_canceled...).")
    df = add_missing(df, ["type_of_meal_plan", "avg_price_per_room"], 0.04)
    issues.append("~4% missing in type_of_meal_plan and avg_price_per_room.")
    df["market_segment_type"] = df["market_segment_type"].apply(lambda x: add_spaces(str(x)))
    issues.append("Whitespace added to market_segment_type.")
    # a few negative lead times (data entry error)
    neg = df.sample(frac=0.005, random_state=5).index
    df.loc[neg, "lead_time"] = -df.loc[neg, "lead_time"].astype(float)
    issues.append("~0.5% negative lead_time values (impossible).")
    df = add_duplicates(df, 0.02)
    issues.append("~2% duplicate rows appended.")
    return df, issues

def dirty_student_dep(df):
    issues = []
    df["suicidal_thoughts"] = df["suicidal_thoughts"].apply(
        lambda x: random.choice(["yes","Yes","Y","no","No"]) if random.random()<0.07 else x)
    issues.append("Inconsistent Yes/No labels in suicidal_thoughts.")
    df["dietary_habits"] = df["dietary_habits"].apply(lambda x: mess_case(str(x)) if random.random()<0.10 else x)
    issues.append("Mixed-case dietary_habits.")
    df = add_missing(df, ["cgpa", "sleep_duration", "financial_stress"], 0.04)
    issues.append("~4% missing in cgpa, sleep_duration, financial_stress.")
    df = add_duplicates(df, 0.02)
    issues.append("~2% duplicate rows appended.")
    return df, issues

def dirty_social(df):
    issues = []
    df["gender"] = df["gender"].apply(lambda x: random.choice(["male","Male","Female","female"]) if random.random()<0.06 else x)
    issues.append("Mixed gender casing.")
    df = add_missing(df, ["avg_daily_usage_hours", "mental_health_score"], 0.04)
    issues.append("~4% missing in usage hours and mental health score.")
    df["most_used_platform"] = df["most_used_platform"].apply(lambda x: mess_case(str(x)) if random.random()<0.08 else x)
    issues.append("Mixed-case platform names.")
    df = add_duplicates(df, 0.02)
    issues.append("~2% duplicate rows appended.")
    return df, issues

CONFIG = {
    "healthcare_records": dirty_healthcare,
    "salary_data": dirty_salary,
    "superstore_sales": dirty_superstore,
    "student_performance": dirty_student_perf,
    "diabetes_prediction": dirty_diabetes,
    "hotel_reservations": dirty_hotel,
    "student_depression": dirty_student_dep,
    "social_media_addiction": dirty_social,
}

def main():
    for name, fn in CONFIG.items():
        df = pd.read_csv(os.path.join(CLEAN, name + ".csv"))
        ddf, issues = fn(df)
        out = os.path.join(DIRTY, name + "_dirty.csv")
        ddf.to_csv(out, index=False)
        record(name, issues)
        print(f"[OK] {name}_dirty.csv  rows={len(ddf)} cols={ddf.shape[1]}  issues={len(issues)}")
    # write manifest
    with open(os.path.join(ROOT, "Datasets", "dirty_manifest.md"), "w") as f:
        f.write("# Dirty Dataset Issue Manifest\n\n")
        f.write("Every file below is a deliberately corrupted copy of its Clean counterpart.\n")
        f.write("Use them for data-cleaning, QA and exploratory-analysis exercises.\n\n")
        for name, issues in MANIFEST:
            f.write(f"## {name}_dirty.csv\n")
            for i in issues:
                f.write(f"- {i}\n")
            f.write("\n")
    print("Wrote Datasets/dirty_manifest.md")

if __name__ == "__main__":
    main()
