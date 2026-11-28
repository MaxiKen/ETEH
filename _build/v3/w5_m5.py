"""Capstone edition-3 slide copy: Module P.5, Capstone 3 (Mercyland Hospital operations). Figures from w5_fx, i.e. from
the reference analysis behind Capstone3_Mercyland_SOLUTION.xlsx (SQL extraction from mercyland_hospital.db)."""
from slidekit import GREEN, NAVY, ORANGE, RED
from v3kit import T
from w5_fx import (ABOVE, ADM, ALOS, BEDS, BEDS_NEED, BELOW, BILL, BD_LONG, C3, CEN_DAYS, CTL, D95, D100, DEPT_CI, EP_ALL, ESC,
                   HMO, LONG, LOSB, LOSMAX, MEDLOS, MH_TABLES, MOCC, N_LONG, N_MOVER, NCONS, NQ4, NWARDS, OCC, PK_MON, RR,
                   SH_LONG, TYPE_LOS, WARD_LOS, WARD_P90, WARDS_CSV, YEAR, Nb, I, P)

MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
_mk = sorted(MOCC)
_mv = [float(MOCC[k]) for k in _mk]
_dept = sorted(DEPT_CI, key=lambda d: -DEPT_CI[d][0])
_dcol = [RED if DEPT_CI[d][3].startswith("Above") else (GREEN if DEPT_CI[d][3].startswith("Below") else NAVY) for d in _dept]
_short = {"Obstetrics & Gynaecology": "Obs & Gynae", "Internal Medicine": "Int. Medicine"}
_ctl = CTL.loc[str(YEAR)]
FREQ = C3.R["Q5_frequent_attenders"]
CH_OCC = dict(kind="line", cats=MON, series=[("Occupancy", _mv), ("Target 85%", [0.85] * 12)], fmt="0%", colors=[NAVY, ORANGE],
              labels=[_mk.index(PK_MON)], label_fmt="0.0%", y_min=0.6, y_max=1.05, title=f"Monthly occupancy {YEAR}")
CH_LOS = dict(kind="histogram", cats=[b for b, _ in LOSB], series=[("Admissions", [n for _, n in LOSB])], fmt="#,##0", labels=True,
              legend=False, title="Length of stay")
CH_RR = dict(kind="column", cats=[_short.get(d, d) for d in _dept], series=[("Readmission", [DEPT_CI[d][0] for d in _dept])], fmt="0%",
             legend=False, point_colors=_dcol, title="30-day readmission by department")
DASH = dict(kind="dashboard", full=True, h=4.95, banner=f"Mercyland FY{YEAR}: capacity, flow and safety",
            kpis=[("Admissions", I(len(ADM)), f"FY{YEAR}"), ("Occupancy", P(OCC), "target 85%", RED),
                  ("Days over 95%", I(D95), f"{D100} over 100%", RED), ("Mean stay", f"{ALOS:.1f} days", f"median {MEDLOS:.0f}"),
                  ("Readmission", P(RR), f"{len(ABOVE)} departments above", ORANGE)],
            charts=[CH_OCC, CH_LOS, CH_RR], widths=[1.15, 0.95, 1])

M_P5 = dict(
    id="P.5", title="Capstone 3: Mercyland Hospital Operations",
    subtitle="A complete healthcare build from mercyland_hospital.db: SQL extraction, occupancy, length of stay, readmission "
             "safety, one escalation rule, a board dashboard and a board pack, with no patient identifiers.",
    objectives=["Frame the board's questions under data-protection rules.", "Extract with SQL and reconcile to control totals.",
                "Measure occupancy with bed-days and analyse length of stay.",
                "Judge readmission differences with confidence intervals.", "Brief the board and verify the pack."],
    topics=[
        T("P.5.1", "Capstone 3 Brief and Success Criteria",
          defn=f"Mercyland ({NWARDS} wards, {BEDS} staffed beds) must set FY{YEAR + 1} capacity and safety priorities. How full "
               "are we, who stays longest, which readmission rates are genuinely worse, and is the payer mix a risk?",
          grid=dict(label="Rules fixed in the brief", widths=[1.6, 4.3], size=12, rows=[
              ["Rule", "Applied as"],
              ["Data minimisation", "Only non-identifying columns leave the database"],
              ["Occupancy", "Σ inpatients ÷ Σ staffed beds (bed-days), never an average of daily percentages"],
              ["Safety claims", "A department is flagged only when its 95% interval excludes the hospital rate"],
              ["Sign-off", "The clinical lead validates readmission findings before the board sees them"]])),

        T("P.5.2", "Mercyland Data Inventory",
          defn=f"The source is mercyland_hospital.db: {len(MH_TABLES)} tables ({', '.join(MH_TABLES)}).",
          grid=dict(label="Tables and grain", widths=[1.4, 1.0, 3.2], size=12, rows=[
              ["Table", "Rows", "Grain"], ["admissions", I(EP_ALL), "One inpatient episode, 2019-2024"],
              ["wards", I(len(WARDS_CSV)), f"One ward; {' + '.join(map(str, WARDS_CSV.Open_Beds))} = {WARDS_CSV.Open_Beds.sum()} beds"],
              ["doctors", I(NCONS), "One consultant"], ["bed_census", I(CEN_DAYS), "One day"],
              ["patients", "", "One patient: identifiers, never extracted"]]),
          watch="Patient names and numbers stay in the database. The analysis never needs them."),

        T("P.5.3", "Extracting and Reconciling Clinical Data",
          defn="Extract with five SQL queries, then reconcile the Excel extract to the database control totals before any "
               "analysis.",
          code=["SELECT substr(admit_date, 1, 4) AS year,", "       COUNT(*) AS admissions,",
                "       ROUND(SUM(bill_amount_ngn), 2) AS billing_ngn,",
                "       SUM(CASE WHEN readmitted_30d = 'Yes'", "                THEN 1 ELSE 0 END) AS readmissions",
                "FROM admissions", "GROUP BY year;"],
          code_label="Q1  CONTROL TOTALS BY YEAR (Capstone3_Extraction_Queries.sql)", code_size=11.5,
          grid=dict(label=f"FY{YEAR} reconciliation", widths=[1.5, 1.5, 0.8], size=12, rows=[
              ["Control", "Database (Q1)", "Excel"], ["Admissions", I(_ctl.admissions), I(len(ADM))],
              ["Billing", Nb(_ctl.billing_ngn), Nb(BILL)], ["Readmissions", I(_ctl.readmissions), I((ADM.readmitted_30d == "Yes").sum())],
              ["Census days", "365", I(len(C3.c23))]]),
          order=["code", "grid"], split=0.58),

        T("P.5.4", "Measuring Occupancy Properly",
          defn="Occupancy = occupied bed-days ÷ available bed-days. Sum both over the period, then divide.",
          simple=f"FY{YEAR}: {P(OCC)} against an 85% target; {N_MOVER} of 12 months were above it.",
          chart=CH_OCC | dict(h=2.6, title=f"Peak {MON[_mk.index(PK_MON)]} ({P(MOCC[PK_MON])}); {D95} days above 95%"),
          grid=dict(label="Capacity", widths=[2.6, 1.0], size=12.5, rows=[
              ["Measure", "Value"], ["Occupancy (bed-days)", P(OCC)], ["Days above 95%", I(D95)], ["Days above 100%", I(D100)],
              ["Extra beds for 85%", I(BEDS_NEED)]]),
          order=["chart", "grid"], split=0.6,
          watch="An average of daily percentages gives the same answer only while beds never change. Use bed-days."),

        T("P.5.5", "Length of Stay Analysis",
          defn="Length of stay (LOS) is usually right-skewed: most patients go home quickly and a few stay a long time.",
          chart=CH_LOS | dict(h=2.5, title=f"Mean {ALOS:.1f} days, median {MEDLOS:.0f}, longest {LOSMAX}"),
          grid=dict(label="By ward", widths=[1.5, 0.8, 0.8, 0.7], size=12, rows=[
              ["Ward", "Mean", "Median", "P90"]] + [[w.replace(" Ward", ""), f"{r['mean']:.1f}", f"{r['median']:.0f}", f"{WARD_P90[w]:.0f}"]
                                                     for w, r in WARD_LOS.iterrows()]),
          order=["chart", "grid"], split=0.55,
          watch=f"Stays of {LONG}+ days are {P(SH_LONG, 0)} of admissions but {P(BD_LONG, 0)} of bed-days: the lever for capacity."),

        T("P.5.6", "Readmission Analysis and Safety",
          defn="Compare each department's 30-day readmission rate with the hospital rate using a 95% confidence interval; "
               "flag it only when the interval excludes the hospital rate.",
          chart=CH_RR | dict(h=2.5, title=f"Hospital rate {P(RR)}: red above, green below"),
          grid=dict(label="Rate and 95% interval", widths=[1.6, 0.8, 1.4], size=11.5, rows=[
              ["Department", "Rate", "95% interval"]] + [[_short.get(d, d), P(DEPT_CI[d][0]), f"{P(DEPT_CI[d][1])}-{P(DEPT_CI[d][2])}"]
                                                          for d in _dept]),
          order=["chart", "grid"], split=0.52,
          watch="A flag is a question for the clinical lead, not a verdict: case mix may explain it."),

        T("P.5.7", "Escalation Logic: One Rule for Every Report",
          defn="An escalation rule flags cases with one documented order of tests, applied identically in every report.",
          code=["= Table.AddColumn(Source, \"Escalation_reason\", each",
                "    if [Readmitted_30d] = \"Yes\" then \"Readmitted within 30 days\"",
                "    else if [LOS_Days] >= 28 then \"Stay of 28+ days\"",
                "    else if [Age] >= 75 then \"Age 75+\"", "    else \"\", type text)"],
          code_label="POWER QUERY CONDITIONAL COLUMN", code_size=11,
          grid=dict(label=f"Q4 {YEAR}: {I(NQ4)} admissions", widths=[1.8, 1.0], size=12.5, rows=[
              ["Reason (first match)", "Cases"]] + [[k, I(v)] for k, v in ESC.items()] + [["Total escalations", I(ESC.sum())]]),
          order=["code", "grid"], split=0.6,
          watch="Order matters: a readmitted 80-year-old counts once, as a readmission."),

        T("P.5.8", "Building the Hospital Dashboard",
          defn="The hospital dashboard answers the board's three questions on one page: capacity, flow and safety.",
          grid=dict(label="Payer mix: share of FY billing", widths=[2.8, 1.0], size=12.5, rows=[
              ["Payer", "Share"]] + [[h, P(v)] for h, v in HMO.items()]),
          watch=f"No payer holds more than {P(HMO.max(), 0)} of billing: low concentration risk.",
          more=[dict(title="The Capstone 3 Dashboard", pill="DASHBOARD", chart=DASH,
                     notes="Five KPI cards and three charts: occupancy against the 85% target, the length-of-stay distribution "
                           "and readmission by department coloured by the confidence-interval verdict.")]),

        T("P.5.9", "The Clinical Board Pack",
          defn="The board pack gives each question one paragraph, one chart and one recommendation with an owner.",
          grid=dict(label="Board summary", widths=[1.2, 4.7], size=11.5, rows=[
              ["Area", "Finding"],
              ["Capacity", f"Occupancy {P(OCC)} (target 85%); {D95} days above 95%; about {BEDS_NEED} extra staffed beds needed"],
              ["Flow", f"Mean stay {ALOS:.1f} days (median {MEDLOS:.0f}); {LONG}+ day stays use {P(BD_LONG, 0)} of bed-days"],
              ["Safety", f"Readmission {P(RR)}; above the rate: {', '.join(ABOVE)}; below: {', '.join(BELOW)}"],
              ["Action", f"Fund {BEDS_NEED} beds or a discharge team for {LONG}+ day stays (COO); readmission review (Medical Director)"]]),
          watch=f"Frequent attenders are counted, not named: {I(FREQ.patients.sum())} patients had 3+ admissions in FY{YEAR}."),

        T("P.5.10", "Mercyland Verification and Handover",
          defn="Verify the pack: extract = control totals, occupancy recomputed from the census, every safety flag signed off, "
               "and no identifiers in any output.",
          grid=dict(label="Verification checks", widths=[2.6, 2.2, 0.8], size=12.5, rows=[
              ["Check", "Evidence", "Status"], ["Admissions = database", f"{I(len(ADM))} = {I(_ctl.admissions)}", "Pass"],
              ["Billing = database", Nb(BILL), "Pass"], ["LOS bands add up", f"{I(sum(n for _, n in LOSB))} = {I(len(ADM))}", "Pass"],
              ["Identifiers in outputs", "None (Q2 selects none)", "Pass"], ["Clinical sign-off", "Readmission flags reviewed", "Pass"]])),
    ],
    takeaways=["Fix the data-protection rules in the brief.", "Reconcile the SQL extract to control totals first.",
               "Occupancy is bed-days ÷ bed-days.", "LOS is right-skewed: report median and P90; long stays fill beds.",
               "Flag safety differences only when the confidence interval says so."],
)

MODULES = [M_P5]
