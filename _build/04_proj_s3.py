"""Student Project 3 - Mercyland diabetes screening clinic: clean, measure prevalence, test a screening rule, handle PHI."""
from __future__ import annotations

import numpy as np
import pandas as pd

from asg_common import Asg
from common import SPR
from proj_common import SKEY, brief_pdf
from wk_common import DATASETS_TEXT, L

NAME = "Project3_Mercyland_Diabetes_Screening"
FOLDER = SPR / "Week_5_Capstone" / NAME
(FOLDER / "Data").mkdir(parents=True, exist_ok=True)
raw = L("dirty_diab").copy()
raw.to_csv(FOLDER / "Data" / "Mercyland_Screening_Clinic_Extract_RAW.csv", index=False)

c = raw.drop_duplicates().copy()
n_dedup = len(c)
n_code = int(raw.Diabetes.isin(["0", "1"]).sum())
n_pct = int(raw.HbA1c_Pct.map(lambda x: isinstance(x, str)).sum())
c["Diabetes"] = c.Diabetes.replace({"1": "Yes", "0": "No"})
c["HbA1c_Pct"] = c.HbA1c_Pct.map(lambda x: float(str(x).replace("%", "")) if isinstance(x, str) else float(x))
for col in ["Age", "BMI", "Glucose_mg_dL"]:
    c[col] = pd.to_numeric(c[col])
n_over70 = int((c.BMI > 70).sum())
n_slip = int((c.BMI >= 100).sum())
n_extreme = n_over70 - n_slip
c.loc[c.BMI >= 100, "BMI"] = (c.loc[c.BMI >= 100, "BMI"] / 10).round(2)
n_bmi_blank = int(c.BMI.isna().sum())
c["Dx"] = (c.Diabetes == "Yes").astype(int)
ref = L("lc_diab").set_index("Screen_ID").loc[c.Screen_ID]
assert (ref.Diabetes.values == c.Diabetes.values).all()
assert np.allclose(ref.HbA1c_Pct.values, c.HbA1c_Pct.values)
ok = c.BMI.notna()
assert np.allclose(ref.BMI.values[ok.values], c.BMI[ok].values, atol=0.01)

prev = c.Dx.mean()
by_g = c.groupby("Gender").Dx.agg(["mean", "size"])
band = pd.cut(c.Age, [-1, 17, 39, 59, 200], labels=["0-17", "18-39", "40-59", "60+"])
by_age = c.groupby(band, observed=True).Dx.mean()
by_ht = c.groupby("Hypertension").Dx.mean()
rr_ht = by_ht["Yes"] / by_ht["No"]
bmicat = pd.cut(c.BMI, [0, 18.5, 25, 30, 200], right=False, labels=["Underweight", "Normal", "Overweight", "Obese"])
by_bmi = c.groupby(bmicat, observed=True).Dx.mean()
by_smk = c.groupby("Smoking_Status").Dx.mean()
hba_d, hba_n = c[c.Dx == 1].HbA1c_Pct.mean(), c[c.Dx == 0].HbA1c_Pct.mean()
flag = (c.HbA1c_Pct >= 6.5) | (c.Glucose_mg_dL >= 200)
sens = flag[c.Dx == 1].mean()
fp = int((flag & (c.Dx == 0)).sum())
ppv = c.Dx[flag].mean()
r_hg = np.corrcoef(c.HbA1c_Pct, c.Glucose_mg_dL)[0, 1]
g200 = c[c.Glucose_mg_dL >= 200].Dx.mean()


def build(key):
    A = Asg("P3", "Student Project 3 - Mercyland Diabetes Screening Clinic", key,
            "Student Project 3 · cleaning · prevalence · screening-rule evaluation · protected health information",
            path=(SKEY / f"{NAME}_KEY.xlsx") if key else (FOLDER / f"{NAME}_Workbook.xlsx"), label="Student Project 3",
            submission="Submit this workbook, your analysis workbook, a dashboard and a 1-page clinical brief. Names: Surname_Firstname_P3_*.")
    s = A.task("M1 Clean", "Milestone 1 - Clean the clinic extract", "2.5-2.9",
               "Data/Mercyland_Screening_Clinic_Extract_RAW.csv is a 12,000-patient sample from the screening clinic, exported with defects. Clean it and log each fix.")
    rows = [("Rows in the raw extract.", len(raw), "int", 1, ""),
            ("Rows after removing duplicate screenings.", n_dedup, "int", 2, "80 duplicates (same Screen_ID, identical rows)."),
            ("Rows where Diabetes is coded 1/0 instead of Yes/No (raw).", n_code, "int", 1, "Recode 1→Yes, 0→No."),
            ("HbA1c values typed as text with a % sign (raw).", n_pct, "int", 1, "Strip '%' and convert; HbA1c is a % of haemoglobin - the unit belongs in the header."),
            ("BMI values above 70 after de-duplication (fail the 10-70 range check).", n_over70, "int", 2, ""),
            ("Of those, how many are 100 or more - decimal slips to divide by 10?", n_slip, "int", 3,
             f"Slips are all ≥ 100 (e.g. 273.2 → 27.32). The other {n_extreme} (72-96) are plausible severe-obesity readings: keep and flag for clinical review, do not 'fix'."),
            ("BMI values that are blank after cleaning (keep blank - exclude from BMI statistics).", n_bmi_blank, "int", 2, ""),
            ("How should Smoking_Status = 'Not Recorded' be treated?", "Keep as its own category", "gen", 3,
             "It is 36% of rows - deleting it biases every result; imputing invents data."),
            ("Mean BMI of the clean data (blanks excluded, 2 dp).", round(c.BMI.mean(), 2), "dec2", 3, "With the slips left in, the mean is visibly inflated."),
            ("Mean HbA1c of the clean data (2 dp).", round(c.HbA1c_Pct.mean(), 2), "dec2", 1, "")]
    for i, (t, v, k, m, gd) in enumerate(rows, 1):
        ch = ["Keep as its own category", "Delete the rows", "Replace with 'Never'", "Replace with the most common value"] if i == 8 else None
        A.q(s, f"1.{i}", t, v, None, k, m, gd, choices=ch, header=(i == 1))
    A.long(s, "1.11", "Cleaning log - defect, rows affected, fix, and why the fix is safe.",
           f"Duplicates 80 → removed; Diabetes 1/0 {n_code} → Yes/No; HbA1c '%' text {n_pct} → number; BMI ×10 slips {n_slip} (≥100) → ÷10; {n_extreme} BMI 70-100 → kept, flagged; BMI blank {n_bmi_blank} → kept "
           "blank, excluded from BMI stats; 'Not Recorded' smoking → kept as a category.", marks=6, rows=3)
    A.end_task(s)

    s = A.task("M2 Prevalence", "Milestone 2 - Who has diabetes?", "2.10, 4.1",
               "Prevalence = share of screened patients with Diabetes = Yes. Report each to 0.1 pt.")
    k = [("Overall prevalence.", prev, 2, ""),
         ("Prevalence - Male.", by_g.loc["Male", "mean"], 2, ""),
         ("Prevalence - Female.", by_g.loc["Female", "mean"], 2, ""),
         ("Prevalence - age 60+.", by_age["60+"], 2, ""),
         ("Prevalence - age 18-39.", by_age["18-39"], 2, ""),
         ("Prevalence - with hypertension.", by_ht["Yes"], 2, ""),
         ("Relative risk = prevalence WITH hypertension ÷ prevalence WITHOUT hypertension.", rr_ht, 4, "Shown as a ratio (e.g. 3.2x)."),
         ("Prevalence - Obese (BMI ≥ 30).", by_bmi["Obese"], 4, "Cut-offs: <18.5, 18.5-24.9, 25-29.9, ≥30; blanks excluded."),
         ("Prevalence - Smoking_Status 'Current'.", by_smk["Current"], 2, ""),
         ("Prevalence - Smoking_Status 'Not Recorded'.", by_smk["Not Recorded"], 3, "Much lower - many are children; check age before concluding.")]
    for i, (t, v, m, gd) in enumerate(k, 1):
        A.q(s, f"2.{i}", t, v, None, "dec2" if i == 7 else "pct", m, gd, header=(i == 1))
    A.end_task(s)

    s = A.task("M3 Screening Rule", "Milestone 3 - Clinical markers and a screening rule", "4.1-4.3",
               "The clinic proposes to flag a patient for diagnosis if HbA1c ≥ 6.5 OR blood glucose ≥ 200 mg/dL. Evaluate it against the recorded outcome. Flagged = the rule says yes. Sensitivity = flagged diabetics ÷ all diabetics; false positive = flagged but no diabetes; PPV = flagged diabetics ÷ all flagged (COUNTIFS does all three).")
    A.q(s, "3.1", "Mean HbA1c - patients WITH diabetes.", hba_d, None, "dec2", 2, header=True)
    A.q(s, "3.2", "Mean HbA1c - patients WITHOUT diabetes.", hba_n, None, "dec2", 2)
    A.q(s, "3.3", "Correlation between HbA1c and blood glucose.", r_hg, None, "dec3", 2, "Weak - the two tests capture different things (long-term vs moment).")
    A.q(s, "3.4", "Share of patients with glucose ≥ 200 who have diabetes.", g200, None, "pct", 2)
    A.q(s, "3.5", "Sensitivity: share of diabetic patients the rule flags.", sens, None, "pct", 4, "Flagged diabetics ÷ all diabetics.")
    A.q(s, "3.6", "False positives: how many non-diabetic patients does the rule flag (count)?", fp, None, "int", 4)
    A.q(s, "3.7", "Positive predictive value (PPV): flagged diabetics ÷ all flagged patients.", ppv, None, "pct", 4)
    A.long(s, "3.8", "Would you recommend the rule? Two sentences with numbers.",
           f"Yes as a screening step: it catches {sens:.1%} of diabetic patients, but only {ppv:.1%} of those it flags have diabetes ({fp:,} false positives), "
           "so every flag needs a confirmatory test before diagnosis. Thresholds are for analysis practice - clinical decisions belong to clinicians.",
           marks=5, rows=3)
    A.end_task(s)

    s = A.task("M4 PHI & Report", "Milestone 4 - Protect the data, then report", "4.7-4.9, 4.13",
               "Mercyland will publish your summary on its website and share row-level data with a university. Apply the Nigeria Data Protection Act 2023 principles from class.")
    A.q(s, "4.1", "How many patients have Gender = 'Other'?", int(by_g.loc["Other", "size"]) if "Other" in by_g.index else 0, None, "int", 2, header=True)
    A.q(s, "4.2", "What must you do with that group in the PUBLISHED summary?", "Suppress or merge the small cell", marks=3,
        choices=["Suppress or merge the small cell", "Publish it - it is accurate", "Delete those patients from the data", "Round to the nearest 10"],
        guide="Cells below 5 can identify individuals - suppress/merge (prevalence of 2 people is also statistically meaningless).")
    A.long(s, "4.3", "List three de-identification steps before row-level data goes to the university.",
           "• Replace Screen_ID with a new random study ID (keep the key file inside Mercyland).\n• Band Age (5- or 10-year bands; top-code 80+).\n"
           "• Remove or generalise rare combinations (small cells), share only the fields needed (data minimisation), under a data-sharing agreement.",
           marks=4, rows=4)
    A.q(s, "4.4", "Dashboard: prevalence by age, BMI category, hypertension and smoking; screening-rule results; small cells suppressed (instructor marks).",
        "Rubric" if key else "", marks=8, header=True)
    A.q(s, "4.5", "One-page clinical brief for the Head of the clinic (instructor marks).", "Rubric" if key else "", marks=8)
    A.end_task(s)
    return A.finish("Student Project 3 - Mercyland Diabetes Screening Clinic",
                    ["Client: Head of the Mercyland diabetes screening clinic. Questions: who is most at risk, how good is the proposed flagging rule, and how can results be shared safely?",
                     "Clean the extract (M1), measure prevalence (M2), evaluate the screening rule (M3) and handle protected health information in your report (M4). Duration about 12 hours."],
                    [("Mercyland Screening", DATASETS_TEXT["Mercyland Screening"] + " - you receive a 12,000-patient sample with deliberate defects.")],
                    files="Data folder (inside this project folder) › Mercyland_Screening_Clinic_Extract_RAW.csv",
                    rubric=[["Health data", "Treat every row as protected health information. Never paste row-level data into slides or AI tools."],
                            ["Tolerance", "Rates ±0.1 pt; means ±0.01."], ["Communication", "Plain English; no clinical advice beyond what the data shows."]])


def main():
    for k in (False, True):
        print("P3", "KEY" if k else "student", "->", build(k))
    brief_pdf(FOLDER / "Project3_Brief.pdf", "Student Project 3 - Mercyland Diabetes Screening Clinic",
              "Cleaning · prevalence analysis · evaluating a screening rule · protected health information", "Student Project 3 brief", [
                  ("h", "1. The client and the question"),
                  ("p", "Mercyland Specialist Hospital runs a walk-in diabetes screening clinic. The Head of the clinic asks: <b>Which patients are most at risk, does our proposed "
                        "flagging rule (HbA1c ≥ 6.5 or glucose ≥ 200 mg/dL) work, and how do we publish results without exposing patients?</b>"),
                  ("h", "2. The data"),
                  ("p", "Data/Mercyland_Screening_Clinic_Extract_RAW.csv - 12,080 rows, one per screening: Screen_ID, Gender, Age, Hypertension, Heart_Condition, Smoking_Status, "
                        "BMI, HbA1c_Pct, Glucose_mg_dL and the recorded Diabetes outcome. The extract has deliberate defects (duplicates, mixed coding, units inside values, "
                        "decimal slips, blanks)."),
                  ("note", "This is protected health information (PHI). Work only on your own machine, never upload rows to online tools, and suppress any published group smaller than 5."),
                  ("h", "3. Milestones"),
                  ("t", [["Milestone", "What you produce", "Marks"], ["M1 Clean", "Clean table, cleaning log, control figures", "25"],
                         ["M2 Prevalence", "Prevalence by sex, age band, hypertension, BMI category, smoking; relative risk", "25"],
                         ["M3 Screening rule", "Marker comparison, sensitivity, false positives, PPV, recommendation", "25"],
                         ["M4 PHI & report", "Small-cell suppression, de-identification plan, dashboard, clinical brief", "25"]], [35, 120, 25]),
                  ("h", "4. Assessment"),
                  ("p", "Accuracy, defensible cleaning decisions (never delete what you can repair), correct handling of PHI, and a brief the clinic could act on. "
                        "Pass 50%, distinction 70%."),
              ])


if __name__ == "__main__":
    main()
