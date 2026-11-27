"""Student Project 4 - ETEH Talent Salary Survey: pay benchmarking, pay equity and salary bands."""
from __future__ import annotations

import numpy as np
import pandas as pd

from asg_common import Asg
from common import SPR
from proj_common import SKEY, brief_pdf
from wk_common import DATASETS_TEXT, L

NAME = "Project4_Talent_Salary_Benchmark"
FOLDER = SPR / NAME
(FOLDER / "Data").mkdir(parents=True, exist_ok=True)
raw = L("sal_dirty").copy()
raw.to_csv(FOLDER / "Data" / "ETEH_Salary_Survey_Form_Export_RAW.csv", index=False)
s = L("sal_clean").copy()
SAL = "Annual_Salary_NGN"

n_dedup = len(raw.drop_duplicates())
n_noblank = len(raw.drop_duplicates().dropna())
fam_counts = s.Job_Family.value_counts()
med_fam = s.groupby("Job_Family")[SAL].median().sort_values(ascending=False)
da = s[s.Job_Family == "Data & Analytics"]
da_q = da[SAL].quantile([.25, .5, .75])
med_sen = s.groupby("Seniority")[SAL].median()
analyst = s[s.Job_Title == "Data Analyst"][SAL].median()
g = s[s.Gender.isin(["Male", "Female"])]
mf = g.groupby("Gender")[SAL].median()
gap_unadj = 1 - mf["Female"] / mf["Male"]
mean_gap = 1 - g[g.Gender == "Female"][SAL].mean() / g[g.Gender == "Male"][SAL].mean()
exp_m = g.groupby("Gender").Years_of_Experience.mean()
mid = g[g.Seniority == "Mid"].groupby("Gender")[SAL].median()
gap_mid = 1 - mid["Female"] / mid["Male"]
X = np.column_stack([g.Years_of_Experience, (g.Gender == "Female").astype(float), np.ones(len(g))])
beta, *_ = np.linalg.lstsq(X, g[SAL].values, rcond=None)
b_exp, b_fem = beta[0], beta[1]
adj_gap = -b_fem / g[g.Gender == "Male"][SAL].mean()
n_other = int((s.Gender == "Other").sum())
slope1 = np.polyfit(s.Years_of_Experience, s[SAL], 1)
pred5 = slope1[0] * 5 + slope1[1]
bands = da.groupby("Seniority")[SAL].quantile([.25, .5, .75]).unstack()
da_mid = bands.loc["Mid"]
OFFER = 9_000_000


def build(key):
    A = Asg("P4", "Student Project 4 - ETEH Talent Salary Survey: Benchmarks, Pay Equity and Bands", key,
            "Student Project 4 · cleaning · benchmarking · pay-gap analysis · regression · salary bands",
            path=(SKEY / f"{NAME}_KEY.xlsx") if key else (FOLDER / f"{NAME}_Workbook.xlsx"), label="Student Project 4",
            submission="Submit this workbook, your analysis workbook, a dashboard and a 1-page pay-equity note. Names: Surname_Firstname_P4_*.")
    sh = A.task("M1 Clean & Standardise", "Milestone 1 - From form export to analysable survey", "1.9, 2.5-2.9",
                "Data/ETEH_Salary_Survey_Form_Export_RAW.csv is the raw web-form export (respondents could submit many times). Clean it, then standardise job titles "
                "into job families and seniority using the rules in the brief. The key reconciles to 00_Core_Datasets/ETEH_Talent_Salary_Survey.csv.")
    rows = [("Rows in the raw export.", len(raw), "int", 1, ""), ("Rows after removing exact duplicates.", n_dedup, "int", 2, ""),
            ("Rows after also removing incomplete responses (any blank).", n_noblank, "int", 2, ""),
            ("Final clean respondents (after education-label standardisation, re-dedup, salary < ₦1m keying errors and experience > age − 16 removed).", len(s), "int", 4,
             "1,787 → 1,786 (label variants) → 1,782 (salaries < ₦1m) → 1,781 (impossible experience)."),
            ("Respondents in the Data & Analytics job family.", int(fam_counts["Data & Analytics"]), "int", 3, "Apply the ordered rules in the brief (Research Scientist → Research, not Data)."),
            ("Respondents classed Executive (Director/VP/Chief/CEO/CTO in title).", int((s.Seniority == "Executive").sum()), "int", 3, ""),
            ("Respondents classed Junior.", int((s.Seniority == "Junior").sum()), "int", 2, "")]
    for i, (t, v, k, m, gd) in enumerate(rows, 1):
        A.q(sh, f"1.{i}", t, v, None, k, m, gd, header=(i == 1))
    A.long(sh, "1.8", "Why must job titles be grouped before benchmarking? Give an example from this survey.",
           f"Titles are free text with {s.Job_Title.nunique()} distinct values - most have only a handful of respondents, so medians would be meaningless and unstable. "
           "Grouping 'Data Analyst', 'Senior Data Scientist', 'Director of Data Science' into Data & Analytics (and seniority separately) gives groups large enough to benchmark.",
           marks=3, rows=3)
    A.end_task(sh)

    sh = A.task("M2 Benchmarks", "Milestone 2 - Market pay benchmarks", "2.10, 4.1",
                "Use medians (robust to a few very high salaries). Quote ₦ to the nearest naira.")
    k = [("Job family with the highest median salary.", med_fam.index[0], "gen", 3, f"₦{med_fam.iloc[0]:,.0f}" + (f"; next {med_fam.index[1]} ₦{med_fam.iloc[1]:,.0f}" if len(med_fam) > 1 else "")),
         ("Median salary - Data & Analytics.", da_q[.5], "naira", 3, ""),
         ("25th percentile - Data & Analytics.", da_q[.25], "naira", 3, "QUARTILE.INC / PERCENTILE.INC."),
         ("75th percentile - Data & Analytics.", da_q[.75], "naira", 3, ""),
         ("Median salary - job title exactly 'Data Analyst'.", analyst, "naira", 3, ""),
         ("Median salary - Senior.", med_sen["Senior"], "naira", 2, ""), ("Median salary - Junior.", med_sen["Junior"], "naira", 2, ""),
         ("Lowest-paying job family (median).", med_fam.index[-1], "gen", 3, f"₦{med_fam.iloc[-1]:,.0f}"),
         ("Small-sample warning: how many respondents are in the Design family?", int(fam_counts["Design"]), "int", 3, "Fewer than ~30: show, but flag as indicative only.")]
    for i, (t, v, kk, m, gd) in enumerate(k, 1):
        A.q(sh, f"2.{i}", t, v, None, kk, m, gd, header=(i == 1))
    A.end_task(sh)

    sh = A.task("M3 Pay Equity", "Milestone 3 - Is there a gender pay gap, and what explains it?", "4.1, 4.3-4.4, 4.13",
                "Compare Male and Female respondents. Respondents with Gender = 'Other' are too few to analyse separately - report them only in totals.")
    A.q(sh, "3.1", "Unadjusted MEDIAN pay gap = 1 − Female median ÷ Male median.", gap_unadj, None, "pct", 3, f"Medians: F ₦{mf['Female']:,.0f}, M ₦{mf['Male']:,.0f}.", header=True)
    A.q(sh, "3.2", "Unadjusted MEAN pay gap.", mean_gap, None, "pct", 2)
    A.q(sh, "3.3", "Average years of experience - Female.", exp_m["Female"], None, "dec2", 1)
    A.q(sh, "3.4", "Average years of experience - Male.", exp_m["Male"], None, "dec2", 1)
    A.q(sh, "3.5", "Median gap WITHIN the Mid seniority level.", gap_mid, None, "pct", 3)
    A.q(sh, "3.6", "Regression Salary ~ Experience + Female (1/0): coefficient on Female (₦).", b_fem, None, "naira", 4,
        "LINEST on [Years_of_Experience, Female] (Male/Female only). Negative = lower pay at the same experience.")
    A.q(sh, "3.7", "Adjusted gap = −(Female coefficient) ÷ Male mean salary.", adj_gap, None, "pct", 3)
    A.q(sh, "3.8", "Respondents with Gender = 'Other' (report in totals only).", n_other, None, "int", 2)
    A.long(sh, "3.9", "Interpret 3.1 vs 3.7 for the HR Director in three sentences.",
           f"Women's median pay is {gap_unadj:.1%} below men's (mean gap {mean_gap:.1%}). About {1 - adj_gap / mean_gap:.0%} of the mean gap is explained by experience (women in the survey average "
           f"{exp_m['Female']:.1f} years vs {exp_m['Male']:.1f}), leaving an adjusted gap of about {adj_gap:.1%} at equal experience. "
           "A survey cannot prove discrimination - role, sector and negotiation are not controlled - but the residual gap justifies an internal pay audit.",
           marks=6, rows=4, guide="Raw vs adjusted 2; experience explanation with numbers 2; appropriate caution + action 2.")
    A.end_task(sh)

    sh = A.task("M4 Salary Bands", "Milestone 4 - Propose Data & Analytics salary bands", "4.4, 4.6-4.9",
                "ETEH Talent Partners advises clients on offers. Build P25-P50-P75 bands for Data & Analytics by seniority and test a live offer.")
    A.q(sh, "4.1", "Data & Analytics, Mid: band minimum (P25).", da_mid[.25], None, "naira", 3, header=True)
    A.q(sh, "4.2", "Data & Analytics, Mid: band midpoint (P50).", da_mid[.5], None, "naira", 3)
    A.q(sh, "4.3", "Data & Analytics, Mid: band maximum (P75).", da_mid[.75], None, "naira", 3)
    A.q(sh, "4.4", f"A client offers a Mid data analyst ₦{OFFER:,}. Where does it sit?",
        "Below band" if OFFER < da_mid[.25] else ("Within band" if OFFER <= da_mid[.75] else "Above band"), None, "gen", 3,
        choices=["Below band", "Within band", "Above band"])
    A.q(sh, "4.5", "Simple regression (all respondents): predicted salary at 5 years' experience (₦).", pred5, None, "naira", 3, "FORECAST.LINEAR(5, salary, experience).")
    A.q(sh, "4.6", "Dashboard: family benchmarks, seniority ladder, gender gap (raw vs adjusted), bands table, small-sample flags (instructor marks).",
        "Rubric" if key else "", marks=8, header=True)
    A.q(sh, "4.7", "One-page note to the client's HR Director with the offer recommendation (instructor marks).", "Rubric" if key else "", marks=7)
    A.end_task(sh)
    return A.finish("Student Project 4 - ETEH Talent Salary Survey",
                    ["Client: ETEH Talent Partners (HR analytics). Questions: what is the market rate by job family and seniority, is there a gender pay gap after accounting "
                     "for experience, and what bands should a client use for data roles?",
                     "Clean and standardise the raw survey (M1), benchmark (M2), analyse pay equity (M3), and propose bands (M4). Duration about 12 hours."],
                    [("Salary Survey", DATASETS_TEXT["Salary Survey"] + " - you receive the raw form export (6,704 rows).")],
                    files="04_Student_Projects/Project4_Talent_Salary_Benchmark/Data/ETEH_Salary_Survey_Form_Export_RAW.csv",
                    rubric=[["Statistics", "Medians and percentiles with PERCENTILE.INC/QUARTILE.INC; regression with LINEST or the ToolPak."],
                            ["Ethics", "Never single out individuals; suppress groups smaller than 5; describe gaps, do not allege discrimination."],
                            ["Tolerance", "Money ±₦1; percentages ±0.1 pt."]])


def main():
    for k in (False, True):
        print("P4", "KEY" if k else "student", "->", build(k))
    brief_pdf(FOLDER / "Project4_Brief.pdf", "Student Project 4 - ETEH Talent Salary Survey: Benchmarks, Pay Equity and Bands",
              "Cleaning and standardising · medians and percentiles · pay-gap analysis · regression · salary bands", "Student Project 4 brief", [
                  ("h", "1. The client and the question"),
                  ("p", "ETEH Talent Partners advises Nigerian employers on pay. It ran an online salary survey and needs three things: <b>market benchmarks by job family and "
                        "seniority, an honest view of the gender pay gap, and salary bands for data roles</b> that clients can use when making offers."),
                  ("h", "2. The data"),
                  ("p", "Data/ETEH_Salary_Survey_Form_Export_RAW.csv - 6,704 form submissions: Age, Gender, Education_Level, Job_Title, Years_of_Experience, Annual_Salary_NGN. "
                        "Many respondents submitted more than once; some fields are blank; education labels vary; a few salaries were keyed without zeros."),
                  ("h", "3. Standardisation rules"),
                  ("t", [["New field", "Rule (apply in this order; first match wins)"],
                         ["Job_Family", "Lower-case the title, then test in order: data → Data & Analytics; software, developer → Software Engineering; engineer → Engineering; "
                                        "marketing; sales; hr, human → Human Resources; financ, account → Finance; product; project → Project Management; operations; design; "
                                        "research; scientist → Data & Analytics; else Other Business"],
                         ["Seniority", "Director / VP / Chief / CEO / CTO → Executive; Senior / Manager / Lead / Principal → Senior; Junior → Junior; else Mid"],
                         ["Education_Level", "Bachelor's Degree → Bachelor's; Master's Degree → Master's; phD → PhD"]], [35, 145]),
                  ("note", "The order matters: 'Senior Data Engineer' is Data & Analytics because 'data' is tested before 'engineer'."),
                  ("h", "4. Milestones"),
                  ("t", [["Milestone", "What you produce", "Marks"], ["M1 Clean & standardise", "Clean survey, job families, seniority, control counts", "20"],
                         ["M2 Benchmarks", "Medians and P25/P75 by family and seniority, small-sample flags", "25"],
                         ["M3 Pay equity", "Raw and experience-adjusted gender gap, interpretation", "25"],
                         ["M4 Salary bands", "Data & Analytics bands, offer test, dashboard, one-page note", "30"]], [40, 115, 25]),
                  ("h", "5. Assessment"),
                  ("p", "Accuracy, correct statistical choices (medians, percentiles, regression), ethical reporting of sensitive comparisons, and a clear recommendation. "
                        "Pass 50%, distinction 70%."),
              ])


if __name__ == "__main__":
    main()
