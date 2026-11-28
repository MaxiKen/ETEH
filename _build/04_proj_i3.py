"""Instructor Capstone 3 - Mercyland Hospital FY2023 board pack: capacity, length of stay, readmission safety, payer mix.
SQL extraction (mercyland_hospital.db) -> reconciled Excel model -> statistical safety flags -> board pack. No patient identifiers leave SQL."""
from __future__ import annotations

import sqlite3

import numpy as np
import pandas as pd

from common import IPR, W3D
from proj_common import brief_pdf, kpi_cards
from xlhelp import Book

OUT = IPR / "Capstone3_Mercyland_Hospital"
OUT.mkdir(parents=True, exist_ok=True)
XLSX = OUT / "Capstone3_Mercyland_SOLUTION.xlsx"
PDF = OUT / "Capstone3_Instructor_Brief.pdf"
SQLF = OUT / "Capstone3_Extraction_Queries.sql"
DB = W3D / "mercyland_hospital.db"
YEAR = 2023
BEDS_TARGET = 0.85

Q = {
    "Q1_control_totals": """-- Q1 Control totals by year (agree these with Health Records and Finance before any analysis)
SELECT substr(admit_date, 1, 4)              AS year,
       COUNT(*)                                   AS admissions,
       ROUND(SUM(bill_amount_ngn), 2)              AS billing_ngn,
       SUM(CASE WHEN readmitted_30d = 'Yes' THEN 1 ELSE 0 END) AS readmissions,
       ROUND(AVG(los_days), 2)              AS alos
FROM admissions
GROUP BY year
ORDER BY year;""",
    "Q2_admissions_2023": """-- Q2 FY2023 admissions extract for Excel - NO names or patient IDs (data minimisation)
SELECT a.episode_id,
       a.admit_date,
       substr(a.admit_date, 1, 7)             AS month,
       'Q' || ((CAST(substr(a.admit_date, 6, 2) AS INTEGER) + 2) / 3) AS quarter,
       w.ward,
       a.department,
       d.grade                                    AS doctor_grade,
       a.admission_route,
       a.age,
       CASE WHEN a.age < 18 THEN '0-17' WHEN a.age < 40 THEN '18-39'
            WHEN a.age < 60 THEN '40-59' WHEN a.age < 75 THEN '60-74' ELSE '75+' END AS age_band,
       a.diagnosis_group,
       a.los_days,
       a.hmo,
       a.bill_amount_ngn,
       a.readmitted_30d
FROM admissions a
JOIN wards   w ON w.ward_code   = a.ward_code
JOIN doctors d ON d.consultant_id = a.consultant_id
WHERE a.admit_date BETWEEN '2023-01-01' AND '2023-12-31'
ORDER BY a.admit_date, a.episode_id;""",
    "Q3_bed_census": """-- Q3 Daily bed census (whole history, for capacity trends)
SELECT census_date,
       substr(census_date, 1, 4) AS year,
       substr(census_date, 1, 7) AS month,
       inpatients,
       open_beds,
       occupancy_pct
FROM bed_census
ORDER BY census_date;""",
    "Q4_readmission_by_department": """-- Q4 Cross-check for the Excel safety table
SELECT department,
       COUNT(*) AS discharges,
       SUM(CASE WHEN readmitted_30d = 'Yes' THEN 1 ELSE 0 END) AS readmissions,
       ROUND(AVG(CASE WHEN readmitted_30d = 'Yes' THEN 1.0 ELSE 0 END), 4) AS rate
FROM admissions
WHERE admit_date LIKE '2023%'
GROUP BY department
ORDER BY rate DESC;""",
    "Q5_frequent_attenders": """-- Q5 Patients admitted 3+ times in FY2023 (counts only - the named list stays in the clinical system)
SELECT admissions_2023, COUNT(*) AS patients
FROM (SELECT patient_no, COUNT(*) AS admissions_2023
      FROM admissions
      WHERE admit_date LIKE '2023%'
      GROUP BY patient_no
      HAVING COUNT(*) >= 3)
GROUP BY admissions_2023
ORDER BY admissions_2023;""",
}

con = sqlite3.connect(DB)
R = {k: pd.read_sql(v, con) for k, v in Q.items()}
con.close()
ctl = R["Q1_control_totals"].set_index("year")
adm = R["Q2_admissions_2023"].copy()
adm["admit_date"] = pd.to_datetime(adm.admit_date)
cen = R["Q3_bed_census"].copy()
cen["census_date"] = pd.to_datetime(cen.census_date)
cen["year"] = cen.year.astype(int)
assert len(adm) == int(ctl.loc[str(YEAR), "admissions"])

SQLF.write_text("-- Mercyland Hospital - Capstone 3 extraction queries (SQLite; run against 01_Section_Examples/Week3_Data_Files/mercyland_hospital.db)\n"
                "-- Every query below was executed during the build; row counts are recorded on sheet 2 of the solution workbook.\n\n"
                + "\n\n".join(Q.values()) + "\n", encoding="utf-8")

# ---------------------------------------------------------------- reference values
c23 = cen[cen.year == YEAR]
occ = c23.inpatients.sum() / c23.open_beds.sum()
d95, d100 = int((c23.occupancy_pct > .95).sum()), int((c23.occupancy_pct > 1).sum())
avg_in = c23.inpatients.mean()
beds_needed = avg_in / BEDS_TARGET - c23.open_beds.mean()
alos = adm.los_days.mean()
rr = (adm.readmitted_30d == "Yes").mean()
bill = adm.bill_amount_ngn.sum()
months = sorted(adm.month.unique())
yrs = [2020, 2021, 2022, 2023]
WARDS = sorted(adm.ward.unique())
TYPES = ["Emergency", "Urgent", "Elective"]
DEPTS = sorted(adm.department.unique())
CONDS = sorted(adm.diagnosis_group.unique())
HMOS = sorted(adm.hmo.unique())
mocc = {m: cen[cen.month == m].inpatients.sum() / cen[cen.month == m].open_beds.sum() for m in months}
n_over = sum(v > BEDS_TARGET for v in mocc.values())
LOSB = [(1, 7, "1-7 days"), (8, 14, "8-14 days"), (15, 21, "15-21 days"), (22, 30, "22-30 days")]


def ci(k, n, p0):
    p = k / n
    se = np.sqrt(p * (1 - p) / n)
    lo, hi = p - 1.96 * se, p + 1.96 * se
    return p, lo, hi, "Above hospital rate" if lo > p0 else ("Below hospital rate" if hi < p0 else "In line")


dept_ci = {d: ci(int((adm[adm.department == d].readmitted_30d == "Yes").sum()), int((adm.department == d).sum()), rr) for d in DEPTS}
cond_ci = {c: ci(int((adm[adm.diagnosis_group == c].readmitted_30d == "Yes").sum()), int((adm.diagnosis_group == c).sum()), rr) for c in CONDS}
above = [d for d, v in dept_ci.items() if v[3].startswith("Above")]
below = [d for d, v in dept_ci.items() if v[3].startswith("Below")]


def build():
    b = Book(str(XLSX), "Capstone 3 - Mercyland Hospital (SOLUTION)", f"Instructor Capstone 3 · Mercyland Hospital · FY{YEAR} board pack")
    F = b.F
    s1 = b.sheet("1 Brief & Plan", "Capstone 3 - Brief, data protection and delivery plan", "", "P.1, P.5.1")
    s2 = b.sheet("2 SQL & Reconcile", "Extraction with SQL, then reconciliation to the database control totals", "Queries: Capstone3_Extraction_Queries.sql", "Vol 3, P.5.2-3")
    s3 = b.sheet("3 Dashboard", f"Board dashboard - FY{YEAR}", "", "P.5.8")
    s4 = b.sheet("4 Capacity", "Bed occupancy - how full is the hospital, and how many beds would 85% need?", "", "P.5.4")
    s5 = b.sheet("5 Length of Stay", "Length of stay by ward, admission type and band", "", "P.5.5")
    s6 = b.sheet("6 Readmission Safety", "30-day readmission - which differences are real? (95% confidence intervals)", "", "P.5.6, Vol 4")
    s7 = b.sheet("7 Payer Mix", "Payer mix and revenue per bed-day", "", "P.5")
    s8 = b.sheet("8 Board Pack", "Board summary, recommendations, defence and marking guide", "", "P.5.9-10, P.7")
    ta = b.data_sheet("D_Adm2023", adm, "tAdm", fmts={"admit_date": "date", "bill_amount_ngn": "naira", "age": "gen", "los_days": "gen"},
                      title=f"Q2 result - FY{YEAR} admissions ({len(adm):,} rows, no patient identifiers)", widths={"department": 24, "hmo": 28, "diagnosis_group": 14})
    tc = b.data_sheet("D_Census", cen, "tCensus", fmts={"census_date": "date", "occupancy_pct": "pct", "year": "gen"}, title=f"Q3 result - daily bed census ({len(cen):,} days)")

    # ---- 1 Brief
    s = s1
    s.band("1.1", "The board's questions")
    s.scenario(f"Mercyland (4 wards × 130 staffed beds) must decide on FY2024 capacity and safety priorities. The board asks: how full are we, which patients stay longest, "
               "are any departments' readmission rates genuinely worse, and does our payer mix carry revenue risk?")
    s.grid([["Rule", "Applied as"], ["Data minimisation", "Only Q2's non-identifying columns leave the database; frequent attenders are counted, not named"],
            ["Occupancy", "Σ patients in house ÷ Σ staffed beds (bed-days), never an average of daily percentages"],
            ["Safety claims", "A department is flagged only when its 95% interval excludes the hospital rate"],
            ["Sign-off", "Clinical lead validates readmission findings before the board sees them"]], merge_last=True)
    s.band("1.2", "Delivery plan (about 5 contact hours)")
    s.grid([["Stage", "Content", "Time"], ["Frame", "Questions, rules, data protection", "15 min"], ["Extract", "Five SQL queries (joins, CASE, HAVING)", "60 min"],
            ["Reconcile", "Control totals SQL = Excel", "15 min"], ["Analyse", "Capacity, LOS, readmission CIs, payer mix", "120 min"],
            ["Communicate", "Dashboard, board pack, defence", "45 min"]], merge_last=True)

    # ---- 2 SQL & Reconcile
    s = s2
    s.band("2.1", "Queries run (full text in the .sql file)")
    s.grid([[k, Q[k].splitlines()[0].lstrip("- "), len(R[k])] for k in Q], header=["Query", "Purpose", "Rows returned"], fmts=[None, None, "int"])
    s.band("2.2", "Q1 control totals by year")
    s.table(ctl.reset_index(), "tControl", fmts={"admissions": "int", "billing_ngn": "naira", "readmissions": "int", "alos": "dec2"})
    s.band("2.3", f"Reconciliation - FY{YEAR} extract (live) vs the database control totals (typed from Q1)")
    k = ctl.loc[str(YEAR)]
    rows = [("Admissions", "=ROWS(tAdm)", len(adm), float(k.admissions), "int"),
            ("Billing (₦)", "=ROUND(SUM(tAdm[bill_amount_ngn]),2)", round(bill, 2), float(k.billing_ngn), "naira"),
            ("Readmissions", '=COUNTIFS(tAdm[readmitted_30d],"Yes")', int((adm.readmitted_30d == "Yes").sum()), float(k.readmissions), "int"),
            ("Census days in year", f"=COUNTIFS(tCensus[year],{YEAR})", len(c23), 365.0, "int")]
    s.grid([], header=["Control", "Database (Q1)", "Excel (live)", "Difference", "Status"])
    for lab, fx, val, exp, kind in rows:
        r = s.r
        s.ws.write(r, 1, lab, F["lab"])
        s.ws.write_number(r, 2, exp, b.nf(kind))
        s.ws.write_formula(r, 3, fx, b.nf(kind, style="result"), float(val))
        s.ws.write_formula(r, 4, f"=ROUND(D{r + 1}-C{r + 1},2)", b.nf(kind), round(float(val) - exp, 2))
        s.ws.write_formula(r, 5, f'=IF(ROUND(D{r + 1}-C{r + 1},2)=0,"OK","CHECK")', F["cell"], "OK" if round(float(val) - exp, 2) == 0 else "CHECK")
        s.r += 1
    s.gap()
    s.band("2.4", "Q5 frequent attenders (counts only)")
    s.table(R["Q5_frequent_attenders"], "tFrequent", fmts={"patients": "int", "admissions_2023": "gen"})
    s.para("The named list stays inside the clinical system; the board sees counts. This is the data-minimisation principle from Vol 3 in practice.", "note")

    # ---- 3 Dashboard
    s = s3
    s.band("3.1", "KPI cards")
    kpi_cards(s, [("ADMISSIONS", "=ROWS(tAdm)", len(adm), "int", f"FY{YEAR}"),
                  ("BED OCCUPANCY", f"=SUMIFS(tCensus[inpatients],tCensus[year],{YEAR})/SUMIFS(tCensus[open_beds],tCensus[year],{YEAR})", occ, "pct", "target 85%"),
                  ("AVERAGE LOS", "=AVERAGE(tAdm[los_days])", alos, "dec1", "days"),
                  ("30-DAY READMISSION", '=COUNTIFS(tAdm[readmitted_30d],"Yes")/ROWS(tAdm)', rr, "pct", "all departments"),
                  ("BILLING", "=SUM(tAdm[bill_amount_ngn])", bill, "nairam", "FY billed")])
    s.band("3.2", "Month by month")
    md = pd.DataFrame({"month": months})
    tm = s.table(md, "tMonthly", calc={"Admissions": ("=COUNTIFS(tAdm[month],[@month])", "int", [int((adm.month == m).sum()) for m in months]),
                                       "Occupancy": ("=SUMIFS(tCensus[inpatients],tCensus[month],[@month])/SUMIFS(tCensus[open_beds],tCensus[month],[@month])", "pct",
                                                     [cen[cen.month == m].inpatients.sum() / cen[cen.month == m].open_beds.sum() for m in months]),
                                       "Readmission": ('=COUNTIFS(tAdm[month],[@month],tAdm[readmitted_30d],"Yes")/[@Admissions]', "pct",
                                                       [(adm[adm.month == m].readmitted_30d == "Yes").mean() for m in months])})
    s.chart("column", [{"name": "Admissions", "categories": tm.ref("month"), "values": tm.ref("Admissions"), "color": "#9DB4C0"}],
            f"Occupancy above the 85% target in {n_over} of 12 months (FY {occ:.1%})", y_title="Admissions",
            secondary={"type": "line", "name": "Occupancy", "categories": tm.ref("month"), "values": tm.ref("Occupancy"), "y2_fmt": "0%", "y2_title": "Occupancy"},
            rows=17, size=(780, 300))

    # ---- 4 Capacity
    s = s4
    s.band("4.1", "Occupancy by year (bed-day method)")
    yd = pd.DataFrame({"year": yrs})
    ty = s.table(yd, "tOccYear", fmts={"year": "gen"},
                 calc={"Days": ("=COUNTIFS(tCensus[year],[@year])", "int", [int((cen.year == y).sum()) for y in yrs]),
                       "Occupancy": ("=SUMIFS(tCensus[inpatients],tCensus[year],[@year])/SUMIFS(tCensus[open_beds],tCensus[year],[@year])", "pct",
                                     [cen[cen.year == y].inpatients.sum() / cen[cen.year == y].open_beds.sum() for y in yrs]),
                       "Days_over_95pct": ('=COUNTIFS(tCensus[year],[@year],tCensus[occupancy_pct],">0.95")', "int", [int((cen[cen.year == y].occupancy_pct > .95).sum()) for y in yrs]),
                       "Days_over_100pct": ('=COUNTIFS(tCensus[year],[@year],tCensus[occupancy_pct],">1")', "int", [int((cen[cen.year == y].occupancy_pct > 1).sum()) for y in yrs])})
    s.band("4.2", "How many beds would bring average occupancy to the target?")
    inp = s.inputs([("Target average occupancy", BEDS_TARGET, "pct0", "Common safe-staffing planning level")])
    T = inp["Target average occupancy"]
    s.calc([("Average patients in house", f"=AVERAGEIFS(tCensus[inpatients],tCensus[year],{YEAR})", avg_in, "dec1", ""),
            ("Staffed beds", f"=AVERAGEIFS(tCensus[open_beds],tCensus[year],{YEAR})", c23.open_beds.mean(), "int", ""),
            ("Additional staffed beds needed", f"=ROUNDUP(AVERAGEIFS(tCensus[inpatients],tCensus[year],{YEAR})/{T}-AVERAGEIFS(tCensus[open_beds],tCensus[year],{YEAR}),0)",
             int(np.ceil(beds_needed)), "int", "Average demand only - peak days need more"),
            (f"Days above 95% in FY{YEAR}", f'=COUNTIFS(tCensus[year],{YEAR},tCensus[occupancy_pct],">0.95")', d95, "int", "Escalation days"),
            (f"Days above 100% in FY{YEAR}", f'=COUNTIFS(tCensus[year],{YEAR},tCensus[occupancy_pct],">1")', d100, "int", "Patients on trolleys / outlying")])
    s.insight(f"Occupancy averaged {occ:.1%} in FY{YEAR} with {d95} escalation days. About {int(np.ceil(beds_needed))} extra staffed beds (≈ one bay per ward) "
              "would bring the average to 85%; alternatively, reducing average LOS by one day frees a similar capacity (see sheet 5).")

    # ---- 5 LOS
    s = s5
    s.band("5.1", "By ward")
    wd = pd.DataFrame({"ward": WARDS})
    s.table(wd, "tLOSWard", widths={"ward": 16},
            calc={"Admissions": ("=COUNTIFS(tAdm[ward],[@ward])", "int", [int((adm.ward == w).sum()) for w in WARDS]),
                  "Mean_LOS": ("=AVERAGEIFS(tAdm[los_days],tAdm[ward],[@ward])", "dec2", [adm[adm.ward == w].los_days.mean() for w in WARDS]),
                  "P90_LOS": ("=AGGREGATE(16,6,tAdm[los_days]/(tAdm[ward]=[@ward]),0.9)", "dec1", [adm[adm.ward == w].los_days.quantile(.9) for w in WARDS]),
                  "Bed_days": ("=SUMIFS(tAdm[los_days],tAdm[ward],[@ward])", "int", [int(adm[adm.ward == w].los_days.sum()) for w in WARDS])})
    s.band("5.2", "By admission type")
    s.table(pd.DataFrame({"admission_route": TYPES}), "tLOSType", widths={"admission_route": 16},
            calc={"Admissions": ("=COUNTIFS(tAdm[admission_route],[@admission_route])", "int", [int((adm.admission_route == t).sum()) for t in TYPES]),
                  "Mean_LOS": ("=AVERAGEIFS(tAdm[los_days],tAdm[admission_route],[@admission_route])", "dec2", [adm[adm.admission_route == t].los_days.mean() for t in TYPES])})
    s.band("5.3", "Distribution of stays")
    lb = pd.DataFrame({"Band": [x[2] for x in LOSB], "From": [x[0] for x in LOSB], "To": [x[1] for x in LOSB]})
    tl = s.table(lb, "tLOSBand", fmts={"From": "gen", "To": "gen"},
                 calc={"Admissions": ('=COUNTIFS(tAdm[los_days],">="&[@From],tAdm[los_days],"<="&[@To])', "int",
                                      [int(adm.los_days.between(a, z).sum()) for a, z, _ in LOSB]),
                       "Share": ("=[@Admissions]/ROWS(tAdm)", "pct", [adm.los_days.between(a, z).mean() for a, z, _ in LOSB])})
    s.chart("column", [{"name": "Admissions", "categories": tl.ref("Band"), "values": tl.ref("Admissions")}], "Stays are spread evenly from 1 to 30 days",
            legend=False, rows=14, size=(520, 240))
    s.calc([("Bed-days saved if every stay of 22+ days were 1 day shorter", '=COUNTIFS(tAdm[los_days],">=22")', int((adm.los_days >= 22).sum()), "int",
             "≈ beds freed = bed-days ÷ 365"),
            ("Equivalent beds freed", '=COUNTIFS(tAdm[los_days],">=22")/365', (adm.los_days >= 22).sum() / 365, "dec1", "")])
    s.para("Mean ≈ median here because stays are uniform; in real hospitals LOS is right-skewed and the median and P90 matter more than the mean.", "note")

    # ---- 6 Readmission
    s = s6
    s.band("6.1", "Departments - rate with a 95% confidence interval")
    inp6 = s.inputs([("z for the interval", 1.96, "dec2", "1.96 = 95% confidence")])
    Z = inp6["z for the interval"]
    hosp = 'COUNTIFS(tAdm[readmitted_30d],"Yes")/ROWS(tAdm)'
    dd = pd.DataFrame({"department": DEPTS})

    def ci_cols(col, keys, ref):
        return {"Discharges": (f"=COUNTIFS(tAdm[{col}],[@{col}])", "int", [int((adm[col] == x).sum()) for x in keys]),
                "Readmitted": (f'=COUNTIFS(tAdm[{col}],[@{col}],tAdm[readmitted_30d],"Yes")', "int", [int((adm[adm[col] == x].readmitted_30d == "Yes").sum()) for x in keys]),
                "Rate": ("=[@Readmitted]/[@Discharges]", "pct", [ref[x][0] for x in keys]),
                "CI_low": (f"=[@Rate]-{Z}*SQRT([@Rate]*(1-[@Rate])/[@Discharges])", "pct", [ref[x][1] for x in keys]),
                "CI_high": (f"=[@Rate]+{Z}*SQRT([@Rate]*(1-[@Rate])/[@Discharges])", "pct", [ref[x][2] for x in keys]),
                "Signal": (f'=IF([@CI_low]>{hosp},"Above hospital rate",IF([@CI_high]<{hosp},"Below hospital rate","In line"))', None, [ref[x][3] for x in keys])}
    td = s.table(dd, "tReadmitDept", widths={"department": 26, "Signal": 20}, calc=ci_cols("department", DEPTS, dept_ci))
    s.chart("bar", [{"name": "Readmission rate", "categories": td.ref("department"), "values": td.ref("Rate"),
                     "points": [{"fill": {"color": "#B23A48" if dept_ci[d][3].startswith("Above") else ("#13807A" if dept_ci[d][3].startswith("Below") else "#9DB4C0")}} for d in DEPTS]}],
            f"{', '.join(above) or 'No department'} genuinely above the {rr:.1%} hospital rate", y_fmt="0%", legend=False, rows=15, size=(620, 260))
    s.band("6.2", "Conditions")
    s.table(pd.DataFrame({"diagnosis_group": CONDS}), "tReadmitCond", widths={"diagnosis_group": 18, "Signal": 20}, calc=ci_cols("diagnosis_group", CONDS, cond_ci))
    s.warn("A flag is a question for the clinical lead, not a verdict: case mix (cancer patients return for planned treatment) may explain Surgery's rate.")

    # ---- 7 Payer mix
    s = s7
    s.band("7.1", "Admissions, billing and revenue per bed-day by payer")
    hd = pd.DataFrame({"hmo": HMOS})
    th = s.table(hd, "tPayer", widths={"hmo": 32},
                 calc={"Admissions": ("=COUNTIFS(tAdm[hmo],[@hmo])", "int", [int((adm.hmo == h_).sum()) for h_ in HMOS]),
                       "Billing": ("=SUMIFS(tAdm[bill_amount_ngn],tAdm[hmo],[@hmo])", "nairam", [adm[adm.hmo == h_].bill_amount_ngn.sum() for h_ in HMOS]),
                       "Share": ("=[@Billing]/SUM(tAdm[bill_amount_ngn])", "pct", [adm[adm.hmo == h_].bill_amount_ngn.sum() / bill for h_ in HMOS]),
                       "Avg_bill": ("=[@Billing]/[@Admissions]", "naira", [adm[adm.hmo == h_].bill_amount_ngn.mean() for h_ in HMOS]),
                       "Per_bed_day": ("=[@Billing]/SUMIFS(tAdm[los_days],tAdm[hmo],[@hmo])", "naira",
                                       [adm[adm.hmo == h_].bill_amount_ngn.sum() / adm[adm.hmo == h_].los_days.sum() for h_ in HMOS])},
                 total={"hmo": "Total", "Admissions": "sum", "Billing": "sum"})
    shares = [adm[adm.hmo == h_].bill_amount_ngn.sum() / bill for h_ in HMOS]
    s.chart("pie", [{"name": "Billing share", "categories": th.ref("hmo"), "values": th.ref("Share"), "labels": True, "label_fmt": "0%"}],
            f"No payer above {max(shares):.0%} of billing - low concentration risk", rows=16, size=(560, 280))

    # ---- 8 Board pack
    s = s8
    s.band("8.1", "Board summary")
    for m in [f"CAPACITY: Occupancy averaged {occ:.1%} in FY{YEAR} (target 85%), with {d95} days above 95% and {d100} above 100%. "
              f"About {int(np.ceil(beds_needed))} extra staffed beds - or shorter long stays - are needed to reach the target.",
              f"FLOW: {len(adm):,} admissions, average stay {alos:.1f} days, evenly spread 1-30 days; {int((adm.los_days >= 22).sum()):,} stays of 22+ days drive bed demand.",
              f"SAFETY: 30-day readmission {rr:.1%}. {', '.join(above) or 'None'} sit genuinely above the hospital rate (95% CI); {', '.join(below) or 'none'} below. "
              "Findings validated with clinical leads before publication.",
              f"FINANCE: ₦{bill / 1e9:,.2f}bn billed; five payers each hold {min(shares):.0%}-{max(shares):.0%} - no concentration risk.",
              f"RECOMMENDATIONS: (1) fund {int(np.ceil(beds_needed))} staffed beds or a discharge-planning team targeting 22+ day stays (owner: COO); (2) Surgery readmission review (owner: Medical Director); "
              "(3) monthly board dashboard from the same SQL queries."]:
        s.para(m, "text")
    s.band("8.2", "Defence questions")
    s.grid([["Why not average the daily occupancy %?", "Bed-day ratios weight each day by its beds; with constant beds they agree, but the method must survive ward changes."],
            ["Is Surgery really worse?", f"Its interval ({dept_ci['Surgery'][1]:.1%}-{dept_ci['Surgery'][2]:.1%}) excludes {rr:.1%}; clinical review decides whether case mix explains it."],
            ["Where are the patient names?", "Never extracted. Q2 selects no identifiers; Q5 returns counts only."],
            ["Do the numbers match Finance?", "Billing reconciles to the database to the kobo (sheet 2)."],
            ["Why 85%?", "It is an input on sheet 4 - change it and the bed requirement recalculates."]],
           header=["Question", "Model answer"], merge_last=True)
    s.band("8.3", "Marking guide")
    s.grid([["SQL extraction", "20", "Joins, CASE, HAVING; no identifiers; row counts recorded"], ["Reconciliation", "15", "All controls OK"],
            ["Capacity & LOS", "25", "Bed-day method; target what-if; distribution"], ["Readmission safety", "20", "CIs and correct interpretation"],
            ["Board pack", "20", "One page, recommendations with owners, confidentiality"]], header=["Criterion", "Marks", "Full marks when…"], merge_last=True)

    b.cover(f"Capstone 3 - Mercyland Hospital FY{YEAR} board pack (SOLUTION)",
            ["Instructor solution workbook. Data arrives through five SQL queries (Capstone3_Extraction_Queries.sql) and is reconciled to database control totals.",
             "No patient names or IDs are present - only admission IDs and aggregate counts."],
            [("Mercyland Hospital", "mercyland_hospital.db (built from 'healthcare_dataset.csv') - admissions, wards, doctors, patients, bed census")])
    b.close()


def brief():
    brief_pdf(PDF, f"Instructor Capstone 3 - Mercyland Hospital FY{YEAR} board pack", "SQL extraction · reconciliation · capacity · confidence intervals · board reporting",
              "Instructor Capstone 3", [
                  ("h", "1. Purpose"),
                  ("p", "Shows the full SQL-to-Excel pipeline on sensitive health data: extract only what is needed, reconcile, analyse with statistical care, report to a board."),
                  ("h", "2. Data"),
                  ("p", "01_Section_Examples/Week3_Data_Files/mercyland_hospital.db and Capstone3_Extraction_Queries.sql (five queries, all tested)."),
                  ("h", "3. Answer key"),
                  ("t", [["Measure", f"FY{YEAR}"], ["Admissions", f"{len(adm):,}"], ["Occupancy (bed-days)", f"{occ:.1%}"], ["Days > 95% / > 100%", f"{d95} / {d100}"],
                         ["Extra beds for 85%", f"{int(np.ceil(beds_needed))}"], ["Average LOS", f"{alos:.2f} days"], ["Readmission", f"{rr:.2%}"],
                         ["Above hospital rate (95% CI)", ", ".join(above) or "none"], ["Below hospital rate", ", ".join(below) or "none"],
                         ["Billing", f"₦{bill:,.2f}"]], [80, 100]),
                  ("h", "4. Teaching points"),
                  ("b", ["SELECT only non-identifying columns; count, don't name.", "Occupancy = Σ patients ÷ Σ beds.",
                         "Rates need intervals: n ≈ 1,800 gives about ±1.5 points.", "Flags start a clinical conversation; they are not verdicts."]),
                  ("note", "Solution: Capstone3_Mercyland_SOLUTION.xlsx in this folder."),
              ])


if __name__ == "__main__":
    build()
    brief()
    print(XLSX, {"occ": round(occ, 4), "d95": d95, "d100": d100, "beds": beds_needed, "alos": round(alos, 2), "rr": round(rr, 4), "above": above, "below": below})
