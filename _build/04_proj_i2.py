"""Instructor Capstone 2 - Kingsway University: progression, drivers and the at-risk list.
Raw registry export -> full Excel-formula cleaning pipeline -> reconciled cohort -> drivers -> threshold-driven at-risk list -> intervention estimate -> senate paper."""
from __future__ import annotations

import numpy as np
import pandas as pd

from common import IPR
from proj_common import brief_pdf, kpi_cards
from wk_common import L
from xlhelp import Book, Sheet

OUT = IPR / "Capstone2_Kingsway_University"
(OUT / "Data").mkdir(parents=True, exist_ok=True)
XLSX = OUT / "Capstone2_Kingsway_SOLUTION.xlsx"
PDF = OUT / "Capstone2_Instructor_Brief.pdf"

cv = L("cv_perf").copy()
raw = L("dirty_cv").copy()
dlog = L("dirty_cv_log")
for c in raw.columns:
    conv = pd.to_numeric(raw[c], errors="coerce")
    if conv.notna().sum() == raw[c].notna().sum():
        raw[c] = conv.astype("Int64") if (conv.dropna() % 1 == 0).all() else conv
RAWCSV = OUT / "Data" / "Kingsway_Registry_Export_RAW.csv"
raw.to_csv(RAWCSV, index=False)

# ---------------------------------------------------------------- reference cleaning (asserted equal to the published clean cohort)
DEPTS = sorted(cv.Department.unique())
VARIANTS = {"Computer Sci.": "Computer Science", "Econs": "Economics", "Acct": "Accounting", "Micro-biology": "Microbiology",
            "Mass Comm": "Mass Communication", "Civil Engr": "Civil Engineering", "Elect. Engineering": "Electrical Engineering"}
dmap = pd.DataFrame([(d, d) for d in DEPTS] + list(VARIANTS.items()), columns=["Recorded_As", "Department"])
DM = {k.upper(): v for k, v in zip(dmap.Recorded_As, dmap.Department)}
NR = ["Lecturer_Rating", "Guardian_Education", "Commute_Band"]
c = raw.copy()
cl = pd.DataFrame({
    "ID_Clean": c.Matric_No.str.strip(),
    "Gender_Clean": c.Gender.replace({"M": "Male", "F": "Female"}),
    "Dept_Clean": c.Department.map(lambda x: DM[x.upper()]),
    "Attendance_Clean": c.Lecture_Attendance_Pct.map(lambda x: int(str(x).rstrip("%"))),
    "Internet_Clean": c.Home_Internet.map(lambda x: "Yes" if str(x).upper().startswith("Y") else "No"),
    "Teacher_Quality_Clean": c.Lecturer_Rating.fillna("Not Recorded"),
    "Parental_Education_Clean": c.Guardian_Education.fillna("Not Recorded"),
    "Distance_Clean": c.Commute_Band.fillna("Not Recorded"),
    "Exam_Clean": c.Final_Score.astype(int).clip(upper=100)})
n_pad = int((c.Matric_No != cl.ID_Clean).sum())
chk = c.assign(Matric_No=cl.ID_Clean, Gender=cl.Gender_Clean, Department=cl.Dept_Clean, Lecture_Attendance_Pct=cl.Attendance_Clean, Home_Internet=cl.Internet_Clean,
               Lecturer_Rating=cl.Teacher_Quality_Clean, Guardian_Education=cl.Parental_Education_Clean, Commute_Band=cl.Distance_Clean, Final_Score=cl.Exam_Clean)
n_exact = len(chk) - len(chk.drop_duplicates())
chk = chk.drop_duplicates("Matric_No")
ref = cv.drop(columns=["Grade", "Attendance_Band", "At_Risk"]).set_index("Matric_No").sort_index()
got = chk.set_index("Matric_No").sort_index()[ref.columns]
assert len(got) == len(ref) and (got.astype(str).values == ref.astype(str).values).all(), "reference clean does not reproduce cv_perf"

# ---------------------------------------------------------------- analysis (pandas reference values)
ST, AT = 62, 65
y = cv.Final_Score
risk = (cv.Final_Score < ST) | (cv.Lecture_Attendance_Pct < AT)
trig = np.where((cv.Final_Score < ST) & (cv.Lecture_Attendance_Pct < AT), "Score and attendance", np.where(cv.Final_Score < ST, "Low score", np.where(risk, "Low attendance", "")))
NUM = ["Study_Hours_Wk", "Lecture_Attendance_Pct", "Last_Sem_Score", "Sleep_Hrs_Night", "Tutorials_Attended", "Exercise_Hrs_Wk"]
corr = {k: float(np.corrcoef(cv[k], y)[0, 1]) for k in NUM}
XP = ["Lecture_Attendance_Pct", "Study_Hours_Wk", "Last_Sem_Score", "Tutorials_Attended"]
X = np.column_stack([cv[k] for k in XP] + [np.ones(len(cv))])
beta, *_ = np.linalg.lstsq(X, y, rcond=None)
r2 = 1 - ((y - X @ beta) ** 2).sum() / ((y - y.mean()) ** 2).sum()
B_ATT = beta[0]


def lifted(lift, coef=B_ATT):
    low = cv.Lecture_Attendance_Pct < AT
    att = cv.Lecture_Attendance_Pct + low * lift
    ex = cv.Final_Score + low * lift * coef
    return int((risk & (att >= AT) & (ex >= ST)).sum())


LIFT = 10
rescued = lifted(LIFT)
mech = int((risk & (cv.Lecture_Attendance_Pct < AT) & (cv.Lecture_Attendance_Pct + LIFT >= AT) & (cv.Final_Score >= ST)).sum())  # already pass the exam bar
cols = ["Matric_No", "Faculty", "Department", "Level", "Gender", "Final_Score", "Lecture_Attendance_Pct", "Study_Hours_Wk", "Last_Sem_Score", "Tutorials_Attended",
        "Sleep_Hrs_Night", "Exercise_Hrs_Wk", "Guardian_Support", "Library_Access", "Motivation", "Home_Internet", "Household_Income",
        "Lecturer_Rating", "Peer_Effect", "Learning_Support", "Guardian_Education", "Commute_Band", "Secondary_School", "Club_Member", "Grade", "Attendance_Band"]
D = cv[cols].copy()
BANDS = ["<70%", "70-79%", "80-89%", "90-100%"]
band_rate = {k: float(risk[cv.Attendance_Band == k].mean()) for k in BANDS}
band_n = {k: int((cv.Attendance_Band == k).sum()) for k in BANDS}
fac = cv.groupby("Faculty").size()
lst = D.loc[risk, ["Matric_No", "Faculty", "Department", "Level", "Gender", "Final_Score", "Lecture_Attendance_Pct"]].copy()
lst["Trigger"] = trig[risk.values]
lst = lst.sort_values("Final_Score", kind="stable")


def build():
    b = Book(str(XLSX), "Capstone 2 - Kingsway University (SOLUTION)", "Instructor Capstone 2 · Kingsway University · 2024/25 progression review")
    F = b.F
    s1 = b.sheet("1 Brief & Plan", "Capstone 2 - Brief, definitions and delivery plan", "Instructor demonstration project", "P.1, P.4.1")
    s2 = b.sheet("2 Clean & Reconcile", "Stage 2-3 - Cleaning the registry export with formulas, then proving the cohort", "", "P.2, P.4.2-3")
    s5 = b.sheet("3 Thresholds & At-Risk", "The progression rule, the at-risk list and the intervention estimate", "Yellow cells drive every sheet", "P.4.6-7")
    s3 = b.sheet("4 Dashboard", "Senate dashboard - is the problem in a faculty, or in student behaviour?", "", "P.4.4, P.4.8")
    s4 = b.sheet("5 Drivers", "What drives exam performance? Correlation, group means and multiple regression", "", "P.4.5, Vol 4")
    s6 = b.sheet("6 Senate Paper", "Senate paper, defence questions, ethics and marking guide", "", "P.4.9-10, P.7")

    # ---- thresholds first (the data table's calculated columns reference them)
    s = s5
    s.band("3.1", "Progression rule (agreed with the Registrar)")
    inp = s.inputs([("Minimum exam score to progress", ST, "int", "Below this → at risk"),
                    ("Minimum attendance % to progress", AT, "int", "Below this → at risk"),
                    ("Attendance lift from the programme (points)", LIFT, "int", "Assumption for the intervention estimate")])
    SC, ATc, LF = (f"'{s.name}'!{inp[k]}" for k in inp)
    SCl, ATl, LFl = (inp[k] for k in inp)

    # ---- data sheets
    def dsheet(name, df, tname, title, **kw):
        sh = Sheet(b, name, title, "")
        b.toc.append((name, title, f"Data sheet - {len(df):,} rows"))
        sh.r = 4
        t = sh.table(df, tname, style="Table Style Light 9", **kw)
        sh.ws.freeze_panes(5, 0)
        return t
    rawd = raw.copy()
    rcalc = {"ID_Clean": ("=TRIM([@Matric_No])", None, cl.ID_Clean.tolist()),
             "Gender_Clean": ('=IF([@Gender]="M","Male",IF([@Gender]="F","Female",[@Gender]))', None, cl.Gender_Clean.tolist()),
             "Dept_Clean": ("=VLOOKUP([@Department],tDeptMap,2,FALSE)", None, cl.Dept_Clean.tolist()),
             "Attendance_Clean": ('=IF(ISNUMBER([@Lecture_Attendance_Pct]),[@Lecture_Attendance_Pct],VALUE(SUBSTITUTE([@Lecture_Attendance_Pct],"%","")))', "gen", cl.Attendance_Clean.tolist()),
             "Internet_Clean": ('=IF(UPPER(LEFT([@Home_Internet],1))="Y","Yes","No")', None, cl.Internet_Clean.tolist()),
             "Teacher_Quality_Clean": ('=IF([@Lecturer_Rating]="","Not Recorded",[@Lecturer_Rating])', None, cl.Teacher_Quality_Clean.tolist()),
             "Parental_Education_Clean": ('=IF([@Guardian_Education]="","Not Recorded",[@Guardian_Education])', None, cl.Parental_Education_Clean.tolist()),
             "Distance_Clean": ('=IF([@Commute_Band]="","Not Recorded",[@Commute_Band])', None, cl.Distance_Clean.tolist()),
             "Exam_Clean": ("=MIN([@Final_Score],100)", "gen", cl.Exam_Clean.tolist())}
    tR = dsheet("D_Raw_Export", rawd, "tRaw", f"Registry export AS RECEIVED ({len(raw):,} rows) + cleaning formulas (columns ID_Clean → Exam_Clean)", calc=rcalc,
                fmts={"Level": "gen", "Final_Score": "gen"}, widths={"Matric_No": 19, "Department": 20, "Dept_Clean": 22})
    tM = dsheet("D_DeptMap", dmap, "tDeptMap", "Department master + recorded variants (VLOOKUP is case-insensitive, so upper-case variants need no row)",
                widths={"Recorded_As": 24, "Department": 24})
    risk_c = (f'=IF(OR([@Final_Score]<{SC},[@Lecture_Attendance_Pct]<{ATc}),"Yes","No")', None, np.where(risk, "Yes", "No").tolist())
    trig_c = (f'=IF([@At_Risk]="No","",IF(AND([@Final_Score]<{SC},[@Lecture_Attendance_Pct]<{ATc}),"Score and attendance",IF([@Final_Score]<{SC},"Low score","Low attendance")))',
              None, trig.tolist())
    tS = dsheet("D_Students", D, "tStudents", f"Clean cohort ({len(D):,} students) - At_Risk and Trigger recalculate from the thresholds on sheet 3",
                calc={"At_Risk": risk_c, "Trigger": trig_c}, fmts={"Level": "gen"}, widths={"Matric_No": 19, "Department": 22, "Trigger": 20})

    # ---- 1 Brief
    s = s1
    s.band("1.1", "The Senate's question")
    s.scenario(f"Kingsway's progression target is 90%. In 2024/25 {1 - risk.mean():.1%} of {len(cv):,} students met the progression rule (exam ≥ {ST} AND attendance ≥ {AT}%). "
               "The Senate asks: which faculties are failing, what drives performance, who needs support in Week 5, and what would a tutorial-plus-attendance programme achieve?")
    s.grid([["Definition", "Rule"], ["Cohort", "One row per registered student after de-duplication (Matric_No is the key)"],
            ["At risk", f"Final_Score < {ST} OR Lecture_Attendance_Pct < {AT}"], ["Progression %", "1 − at-risk share"],
            ["Grade", "A 70-100, B 60-69, C 50-59, D 45-49, F < 45 (Nigerian 5-point scale)"]], merge_last=True)
    s.band("1.2", "Delivery plan (about 5 contact hours)")
    s.grid([["Stage", "What you demonstrate", "Time"], ["Frame", "Rule, cohort definition, success criteria", "15 min"],
            ["Clean", "Formula pipeline on D_Raw_Export; remove duplicates on ID_Clean", "60 min"],
            ["Reconcile", f"6,652 rows → {len(cv):,} students; every check = 0", "20 min"],
            ["Analyse", "Faculty/level/attendance-band rates; drivers; regression", "75 min"],
            ["Act", "Threshold-driven at-risk list; intervention estimate with sensitivity", "45 min"],
            ["Communicate", "Dashboard, senate paper, ethics, defence", "45 min"]], merge_last=True)
    s.band("1.3", "Files")
    s.grid([["03_Instructor_Projects/Capstone2_Kingsway_University/Data/Kingsway_Registry_Export_RAW.csv", f"Raw export, {len(raw):,} rows, 8 defect types"],
            ["01_Section_Examples/00_Core_Datasets/Kingsway_Student_Results.csv", "Published clean cohort - the reconciliation target"]],
           header=["File", "Use"], merge_last=True)

    # ---- 2 Clean & Reconcile
    s = s2
    s.band("2.1", "Defect log")
    s.table(dlog, "tDefects", fmts={"Rows affected": "int"}, widths={dlog.columns[1]: 44, dlog.columns[-1]: 30})
    s.band("2.2", "The cleaning formulas (calculated columns on D_Raw_Export)")
    s.grid([[k, " " + v[0]] for k, v in rcalc.items()], header=["Clean column", "Formula (text)"], merge_last=True)
    s.warn("Attendance typed '84%': VALUE(\"84%\") returns 0.84, not 84 - strip the % sign first. Trim IDs BEFORE removing duplicates, or padded copies survive.")
    s.band("2.3", "Completeness checks - each should read 0 (or the stated value)")
    s.calc([("Rows received", "=ROWS(tRaw)", len(raw), "int", ""),
            ("Padded Student_IDs fixed by TRIM", "=SUMPRODUCT(--(tRaw[Matric_No]<>tRaw[ID_Clean]))", n_pad, "int", "Matches the defect log (25)"),
            ("Gender not Male/Female after cleaning", '=COUNTIFS(tRaw[Gender_Clean],"<>Male",tRaw[Gender_Clean],"<>Female")', 0, "int", ""),
            ("Attendance outside 0-100", '=COUNTIFS(tRaw[Attendance_Clean],">100")+COUNTIFS(tRaw[Attendance_Clean],"<0")', 0, "int", ""),
            ("Exam scores capped at 100", '=COUNTIFS(tRaw[Final_Score],">100")', 1, "int", "One genuine 101 in the source - query, then cap"),
            ("Distinct students (Remove Duplicates on ID_Clean)", "=COUNTA(UNIQUE(tRaw[ID_Clean]))", len(cv), "int", f"{len(raw) - len(cv)} double registrations removed"),
            ("Clean cohort rows", "=ROWS(tStudents)", len(cv), "int", "Must equal the line above"),
            ("Difference", f"=ROWS(tStudents)-{len(cv)}", 0, "int", "Registry headcount = 5,840")])
    s.insight(f"Every cleaned value in tStudents matches the formula pipeline (verified in the build): {n_exact} exact duplicate rows after trimming, cohort {len(cv):,}.")

    # ---- 3 Thresholds & At-Risk (continued)
    s = s5
    s.band("3.2", "Who is at risk, and why")
    s.calc([("At-risk students", '=COUNTIFS(tStudents[At_Risk],"Yes")', int(risk.sum()), "int", ""),
            ("Progression %", '=1-COUNTIFS(tStudents[At_Risk],"Yes")/ROWS(tStudents)', 1 - risk.mean(), "pct", "Target 90%"),
            ("…low score only", '=COUNTIFS(tStudents[Trigger],"Low score")', int((trig == "Low score").sum()), "int", "Academic tutoring"),
            ("…low attendance only", '=COUNTIFS(tStudents[Trigger],"Low attendance")', int((trig == "Low attendance").sum()), "int", "Pastoral call"),
            ("…both", '=COUNTIFS(tStudents[Trigger],"Score and attendance")', int((trig == "Score and attendance").sum()), "int", "Adviser meeting first")])
    s.band("3.3", "Intervention estimate - lift attendance of students below the threshold")
    o33 = s.calc([("Exam points per attendance point (multiple-regression coefficient)", "=INDEX(LINEST(tStudents[Final_Score],tStudents[[Lecture_Attendance_Pct]:[Tutorials_Attended]]),1,4)",
             B_ATT, "dec3", "Holds hours, prior score and tutoring constant (sheet 5)"),
            ("At-risk students who would progress after the lift",
             f"=SUMPRODUCT((tStudents[At_Risk]=\"Yes\")*((tStudents[Lecture_Attendance_Pct]+(tStudents[Lecture_Attendance_Pct]<{ATl})*{LFl})>={ATl})"
             f"*((tStudents[Final_Score]+(tStudents[Lecture_Attendance_Pct]<{ATl})*{LFl}*{B_ATT:.6f})>={SCl}))", rescued, "int", "Coefficient typed as a value so pycel/Excel agree"),
            ("Progression % after the programme (estimate)", f'=1-(COUNTIFS(tStudents[At_Risk],"Yes")-E{s.r + 3})/ROWS(tStudents)', 1 - (risk.sum() - rescued) / len(cv), "pct", "")])
    sens = [(lf, lifted(lf)) for lf in (5, 10, 15)]
    s.grid([[f"+{lf} points", n, 1 - (risk.sum() - n) / len(cv)] for lf, n in sens], header=["Attendance lift", "Students rescued", "Progression %"], fmts=[None, "int", "pct"])
    s.calc([("…of which already pass the exam bar (attendance-only gain)",
             f"=COUNTIFS(tStudents[At_Risk],\"Yes\",tStudents[Lecture_Attendance_Pct],\"<\"&{ATl},tStudents[Lecture_Attendance_Pct],\">=\"&({ATl}-{LFl}),tStudents[Final_Score],\">=\"&{SCl})",
             mech, "int", "Mechanical: attendance is part of the rule"),
            ("…who also need the predicted exam gain", f"={o33['At-risk students who would progress after the lift']}-E{s.r + 2}", rescued - mech, "int", "The genuinely uncertain part of the estimate")])
    s.warn(f"Be honest about the mechanics: {mech} of the {rescued} rescued students already pass the exam bar and progress simply because attendance is part of the rule. "
           f"Only {rescued - mech} depend on the predicted exam gain. Association, not proven cause - evaluate with a comparison group.")
    s.band("3.4", "At-risk list for advisers - live (restricted distribution)")
    s.para("Formula in B-cell below: =SORT(FILTER(HSTACK(tStudents[[Matric_No]:[Lecture_Attendance_Pct]],tStudents[Trigger]),tStudents[At_Risk]=\"Yes\"),6,1) - re-sorts and resizes when thresholds change.",
           "note")
    hdr = ["Matric_No", "Faculty", "Department", "Level", "Gender", "Final_Score", "Lecture_Attendance_Pct", "Trigger"]
    for j, h_ in enumerate(hdr):
        s.ws.write(s.r, 1 + j, h_, F["hdr"])
    s.r += 1
    s.spill(s.r, 1, '=SORT(FILTER(HSTACK(tStudents[[Matric_No]:[Lecture_Attendance_Pct]],tStudents[Trigger]),tStudents[At_Risk]="Yes"),6,1)',
            lst.astype(object).values.tolist())
    s.ws.set_column(1, 1, 19)
    s.ws.set_column(3, 3, 22)
    s.r += len(lst) + 1

    # ---- 4 Dashboard
    s = s3
    s.band("4.1", "KPI cards")
    kpi_cards(s, [("COHORT", "=ROWS(tStudents)", len(cv), "int", "registered students"),
                  ("PROGRESSION", '=1-COUNTIFS(tStudents[At_Risk],"Yes")/ROWS(tStudents)', 1 - risk.mean(), "pct", "target 90%"),
                  ("AT RISK", '=COUNTIFS(tStudents[At_Risk],"Yes")', int(risk.sum()), "int", "Week-5 support list"),
                  ("AVG ATTENDANCE", "=AVERAGE(tStudents[Lecture_Attendance_Pct])", cv.Lecture_Attendance_Pct.mean(), "dec1", "%"),
                  ("AVG EXAM", "=AVERAGE(tStudents[Final_Score])", y.mean(), "dec1", "out of 100")])
    s.band("4.2", "At-risk rate by faculty - flat")
    fd = pd.DataFrame({"Faculty": fac.index})
    tf = s.table(fd, "tFaculty", widths={"Faculty": 20},
                 calc={"Students": ("=COUNTIFS(tStudents[Faculty],[@Faculty])", "int", fac.tolist()),
                       "At_Risk": ('=COUNTIFS(tStudents[Faculty],[@Faculty],tStudents[At_Risk],"Yes")', "int", [int(risk[cv.Faculty == f].sum()) for f in fac.index]),
                       "Rate": ("=[@At_Risk]/[@Students]", "pct", [float(risk[cv.Faculty == f].mean()) for f in fac.index]),
                       "Avg_Exam": ("=AVERAGEIFS(tStudents[Final_Score],tStudents[Faculty],[@Faculty])", "dec1", [float(y[cv.Faculty == f].mean()) for f in fac.index])},
                 total={"Faculty": "Total", "Students": "sum", "At_Risk": "sum"})
    lv = sorted(cv.Level.unique())
    s.band("4.3", "At-risk rate by level - flat")
    tlv = s.table(pd.DataFrame({"Level": lv}), "tLevel", fmts={"Level": "gen"},
                  calc={"Students": ("=COUNTIFS(tStudents[Level],[@Level])", "int", [int((cv.Level == l).sum()) for l in lv]),
                        "Rate": ('=COUNTIFS(tStudents[Level],[@Level],tStudents[At_Risk],"Yes")/[@Students]', "pct", [float(risk[cv.Level == l].mean()) for l in lv])})
    s.band("4.4", "At-risk rate by attendance band - steep")
    ta = s.table(pd.DataFrame({"Attendance_Band": BANDS}), "tAttBand", widths={"Attendance_Band": 16},
                 calc={"Students": ('=COUNTIFS(tStudents[Attendance_Band],"="&[@Attendance_Band])', "int", [band_n[k] for k in BANDS]),
                       "Rate": ('=COUNTIFS(tStudents[Attendance_Band],"="&[@Attendance_Band],tStudents[At_Risk],"Yes")/[@Students]', "pct", [band_rate[k] for k in BANDS])})
    s.para('Criterion "="&[@Attendance_Band] is essential: a bare "<70%" is read by COUNTIFS as "less than 0.7" and counts nothing.', "note")
    s.chart("column", [{"name": "At-risk rate", "categories": ta.ref("Attendance_Band"), "values": ta.ref("Rate"), "color": "#B23A48"}],
            f"At-risk rate falls from {band_rate['<70%']:.0%} (<70% attendance) to {band_rate['90-100%']:.0%} (90%+)", y_fmt="0%", legend=False, rows=15, size=(520, 260))
    s.chart("bar", [{"name": "At-risk rate", "categories": tf.ref("Faculty"), "values": tf.ref("Rate")}],
            "Faculties differ by under 2 points - the problem is not where students study", y_fmt="0%", legend=False, rows=15, size=(520, 240))
    s.insight("Target students by behaviour (attendance, study hours), not faculties: a faculty-level programme would spend 86% of its effort on students already progressing.")

    # ---- 5 Drivers
    s = s4
    s.band("5.1", "Correlation with Final_Score")
    s.calc([(k, f"=CORREL(tStudents[{k}],tStudents[Final_Score])", corr[k], "dec3", "") for k in NUM])
    s.band("5.2", "Group means - categorical factors (exam points)")
    rows = []
    for fct, (hi, lo) in [("Guardian_Support", ("High", "Low")), ("Library_Access", ("High", "Low")), ("Motivation", ("High", "Low")),
                          ("Peer_Effect", ("Positive", "Negative")), ("Learning_Support", ("No", "Yes")), ("Home_Internet", ("Yes", "No"))]:
        gap = y[cv[fct] == hi].mean() - y[cv[fct] == lo].mean()
        rows.append((f"{fct}: {hi} minus {lo}", f'=AVERAGEIFS(tStudents[Final_Score],tStudents[{fct}],"{hi}")-AVERAGEIFS(tStudents[Final_Score],tStudents[{fct}],"{lo}")',
                     float(gap), "dec2", ""))
    s.calc(rows)
    s.band("5.3", "Multiple regression: Exam ~ Attendance + Hours + Previous score + Tutoring")
    lin = "LINEST(tStudents[Final_Score],tStudents[[Lecture_Attendance_Pct]:[Tutorials_Attended]],TRUE,TRUE)"
    s.calc([("R²", f"=INDEX({lin},3,1)", r2, "dec4", "Week 4 two-predictor model: 0.5415"),
            ("Attendance (per point)", f"=INDEX({lin},1,4)", beta[0], "dec3", "LINEST returns coefficients in REVERSE column order"),
            ("Hours studied (per hour/week)", f"=INDEX({lin},1,3)", beta[1], "dec3", ""),
            ("Previous score (per point)", f"=INDEX({lin},1,2)", beta[2], "dec3", ""),
            ("Tutoring (per session)", f"=INDEX({lin},1,1)", beta[3], "dec3", ""),
            ("Intercept", f"=INDEX({lin},1,5)", beta[4], "dec2", "")])
    s.insight(f"Attendance and study hours explain most of the variation (R² {r2:.3f}). Sleep and physical activity show no relationship (|r| < 0.03).")

    # ---- 6 Senate paper
    s = s6
    s.band("6.1", "Model senate paper (one page)")
    for m in [f"CONTEXT: Progression target 90%. FINDING: {1 - risk.mean():.1%} of {len(cv):,} students progressed ({(1 - risk.mean() - .9) * 100:+.1f} points vs target); {int(risk.sum())} are at risk.",
              f"EVIDENCE: At-risk rates are flat across faculties ({min(risk.groupby(cv.Faculty).mean()):.1%}-{max(risk.groupby(cv.Faculty).mean()):.1%}) and levels, but fall from "
              f"{band_rate['<70%']:.0%} for attendance below 70% to {band_rate['90-100%']:.0%} at 90%+. Attendance (r = {corr['Lecture_Attendance_Pct']:.2f}) and study hours "
              f"(r = {corr['Study_Hours_Wk']:.2f}) are the strongest drivers.",
              f"RECOMMENDATION: A Week-5 attendance-and-tutorial programme for the {int(risk.sum())} named students (owners: Deans; advisers hold the list). "
              f"A +{LIFT}-point attendance lift is estimated to move {rescued} students to progression ({mech} of them through attendance alone) ({1 - (risk.sum() - rescued) / len(cv):.1%}); range "
              f"{sens[0][1]}-{sens[2][1]} students for +5 to +15 points.",
              "RISKS & ETHICS: association not cause - evaluate with a comparison group; the named list is restricted to advisers; never publish groups under 10."]:
        s.para(m, "text")
    s.band("6.2", "Defence questions")
    s.grid([["Which faculty is worst?", "None materially - the spread is under 2 points; that is the finding."],
            ["Why not target 100-level students?", "Level rates are also flat (12.7%-14.5%); level is not the driver."],
            ["How reliable is the regression?", f"R² {r2:.2f} on {len(cv):,} students; coefficients are associations, the programme needs evaluation."],
            ["Could the thresholds be wrong?", "They are inputs on sheet 3 - the list, dashboard and estimate all recalculate."],
            ["Who sees names?", "Advisers only; the Senate sees aggregates."],
            ["What about the 101 score?", "Queried with the exams office and capped at 100; it does not affect any rate."]],
           header=["Question", "Model answer"], merge_last=True)
    s.band("6.3", "Marking guide")
    s.grid([["Cleaning & reconciliation", "25", "Formula pipeline; 5,840 students; all checks 0"], ["Analysis", "30", "Faculty/level/band rates, drivers, regression"],
            ["At-risk list & estimate", "20", "Threshold-driven list; estimate with range and caveat"], ["Dashboard", "10", "KPI cards, sorted bars, finding titles"],
            ["Senate paper & ethics", "15", "One page, numbers, restricted list, small-group rule"]], header=["Criterion", "Marks", "Full marks when…"], merge_last=True)

    b.cover("Capstone 2 - Kingsway University: progression, drivers and the at-risk list (SOLUTION)",
            ["Instructor solution workbook. The raw registry export is cleaned with visible formulas, reconciled to the 5,840-student cohort, then analysed.",
             "Change the thresholds on sheet 3: the At_Risk column, dashboard, list and estimate all recalculate."],
            [("Kingsway Student Performance", "Kingsway_Student_Results.csv (built from 'StudentPerformanceFactors.csv')"),
             ("Registry export (dirty)", f"Data/Kingsway_Registry_Export_RAW.csv - {len(raw):,} rows, 8 defect types")])
    b.close()
    return sens


def brief(sens):
    brief_pdf(PDF, "Instructor Capstone 2 - Kingsway University: progression and the at-risk list",
              "Formula cleaning pipeline · reconciliation · drivers & regression · threshold-driven lists · senate paper", "Instructor Capstone 2", [
                  ("h", "1. Purpose"),
                  ("p", "A people-centred capstone: the analysis ends in a named support list, so it also teaches restricted distribution and careful wording of causal claims."),
                  ("h", "2. Data"),
                  ("p", f"Data/Kingsway_Registry_Export_RAW.csv - {len(raw):,} rows: padded IDs, M/F codes, '84%' text, department abbreviations, Y/yes/N variants, blanks, "
                        f"45 double registrations and one 101 score. Target: the published {len(cv):,}-student cohort."),
                  ("h", "3. Answer key"),
                  ("t", [["Measure", "Value"], ["Clean cohort", f"{len(cv):,} (from {len(raw):,})"], ["At risk / progression", f"{int(risk.sum())} / {1 - risk.mean():.1%}"],
                         ["Faculty at-risk range", f"{min(risk.groupby(cv.Faculty).mean()):.1%} - {max(risk.groupby(cv.Faculty).mean()):.1%}"],
                         ["Attendance band <70% vs 90%+", f"{band_rate['<70%']:.1%} vs {band_rate['90-100%']:.1%}"],
                         ["r(attendance) / r(hours)", f"{corr['Lecture_Attendance_Pct']:.3f} / {corr['Study_Hours_Wk']:.3f}"],
                         ["Multiple R² (4 predictors)", f"{r2:.4f}"], ["Attendance coefficient", f"{B_ATT:.3f} exam points per point"],
                         ["Students rescued (+5/+10/+15)", " / ".join(str(n) for _, n in sens) + f" ({mech} of the +10 figure via attendance alone)"]], [80, 100]),
                  ("h", "4. Teaching points"),
                  ("b", ["A flat result is a finding: 'no faculty is the problem' redirects the budget.", "VLOOKUP is case-insensitive - upper-case variants match the master list.",
                         "VALUE(\"84%\") = 0.84: strip the sign first.", "LINEST returns coefficients in reverse order.",
                         "Make thresholds inputs: policy changes become a recalculation, not a rebuild."]),
                  ("note", "Solution: Capstone2_Kingsway_SOLUTION.xlsx. The at-risk list contains student IDs - treat the workbook as restricted."),
              ])


if __name__ == "__main__":
    sens = build()
    brief(sens)
    print(XLSX, "rescued", rescued, "r2", round(r2, 4), "sens", sens, "bands", {k: round(v, 3) for k, v in band_rate.items()})
