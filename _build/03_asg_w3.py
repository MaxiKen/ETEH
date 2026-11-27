"""Week 3 assignment - Lifecare hospital database: SQL, joins, data model, DAX and dashboard (modules 3.1-3.10)."""
from __future__ import annotations

import shutil
import sqlite3

import pandas as pd

from asg_common import KEYDIR, Asg
from common import ASG, W3D
from wk_common import DATASETS_TEXT

DB = W3D / "lifecare_hospital.db"
DATA_DIR = ASG / "Week3_Assignment_Data"
DATA_DIR.mkdir(parents=True, exist_ok=True)
(DATA_DIR / "lifecare_hospital.db").unlink(missing_ok=True)   # use the shared copy in 01_Section_Examples/Week3_Data_Files
con = sqlite3.connect(DB)


def q(sql):
    return pd.read_sql_query(sql, con)


def one(sql):
    return q(sql).iloc[0, 0]


Y23 = "strftime('%Y', admission_date) = '2023'"
adm = q("SELECT admission_id, patient_id, age, medical_condition, department, admission_type, admission_date, discharge_date, length_of_stay, "
        "ward_id, doctor_id, hmo, billing_amount, readmit_30d FROM admissions WHERE strftime('%Y', admission_date) IN ('2022','2023') ORDER BY admission_date, admission_id")
adm["admission_date"] = pd.to_datetime(adm.admission_date)
adm["discharge_date"] = pd.to_datetime(adm.discharge_date)
adm.insert(7, "adm_year", adm.admission_date.dt.year)
wards = q("SELECT * FROM wards")
docs = q("SELECT * FROM doctors")
cen = q("SELECT * FROM bed_census WHERE strftime('%Y', census_date) IN ('2022','2023') ORDER BY census_date")
cen["census_date"] = pd.to_datetime(cen.census_date)
cen["year"] = cen.census_date.dt.year
a23, a22 = adm[adm.adm_year == 2023], adm[adm.adm_year == 2022]
wmap = wards.set_index("ward_id").ward
dgrade = docs.set_index("doctor_id").grade
adm_ward = adm.ward_id.map(wmap)
adm_grade = adm.doctor_id.map(dgrade)
THR = 400

SQLQ = [
    ("How many admissions were there in 2023?",
     f"SELECT COUNT(*) AS admissions\nFROM admissions\nWHERE {Y23};", "int"),
    ("What was the average length of stay in 2023 (2 dp)?",
     f"SELECT ROUND(AVG(length_of_stay), 2) AS alos\nFROM admissions\nWHERE {Y23};", "dec2"),
    ("Which department had the most admissions in 2023?",
     f"SELECT department, COUNT(*) AS n\nFROM admissions\nWHERE {Y23}\nGROUP BY department\nORDER BY n DESC\nLIMIT 1;", "gen"),
    ("Total 2023 billing for Emergency admissions (₦, nearest naira).",
     f"SELECT ROUND(SUM(billing_amount), 0) AS emergency_billing\nFROM admissions\nWHERE {Y23}\n  AND admission_type = 'Emergency';", "naira"),
    ("JOIN: which ward (by NAME) had the highest 30-day readmission rate in 2023?",
     "SELECT w.ward, ROUND(AVG(a.readmit_30d = 'Yes'), 4) AS readmit_rate\nFROM admissions a\nJOIN wards w ON w.ward_id = a.ward_id\n"
     f"WHERE strftime('%Y', a.admission_date) = '2023'\nGROUP BY w.ward\nORDER BY readmit_rate DESC\nLIMIT 1;", "gen"),
    ("JOIN: which doctor (by NAME) handled the most admissions in 2023?",
     "SELECT d.doctor_name, COUNT(*) AS n\nFROM admissions a\nJOIN doctors d ON d.doctor_id = a.doctor_id\n"
     "WHERE strftime('%Y', a.admission_date) = '2023'\nGROUP BY d.doctor_name\nORDER BY n DESC\nLIMIT 1;", "gen"),
    (f"HAVING: how many doctors handled MORE than {THR} admissions in 2023?",
     f"SELECT COUNT(*) AS busy_doctors FROM (\n  SELECT doctor_id\n  FROM admissions\n  WHERE {Y23}\n  GROUP BY doctor_id\n  HAVING COUNT(*) > {THR}\n);", "int"),
    ("How many patients have been admitted three or more times (all years)?",
     "SELECT COUNT(*) AS frequent_patients FROM (\n  SELECT patient_id\n  FROM admissions\n  GROUP BY patient_id\n  HAVING COUNT(*) >= 3\n);", "int"),
    ("Average daily bed occupancy in 2023 from bed_census (as a %, 1 dp).",
     "SELECT ROUND(AVG(occupancy_pct) * 100, 1) AS avg_occupancy_pct\nFROM bed_census\nWHERE strftime('%Y', census_date) = '2023';", "dec1"),
    ("SUBQUERY: how many 2023 admissions were billed ABOVE the 2023 average bill?",
     f"SELECT COUNT(*) AS above_avg\nFROM admissions\nWHERE {Y23}\n  AND billing_amount > (SELECT AVG(billing_amount)\n                        FROM admissions\n                        WHERE {Y23});", "int"),
]
SQLA = [one(sqlx) for _, sqlx, _ in SQLQ]
WRONG = int(one("SELECT COUNT(*) FROM admissions a JOIN doctors d ON a.department = d.department WHERE strftime('%Y', a.admission_date) = '2023'"))
RIGHT = int(SQLA[0])

# student SQL template + key SQL file
tpl = ["-- ETEH Week 3 Assignment - SQL answers (lifecare_hospital.db)", "-- Write each query under its question. Run it in DB Browser for SQLite and copy the result into the workbook.", ""]
keyl = ["-- ETEH Week 3 Assignment - ANSWER KEY SQL (instructor only)", ""]
for i, ((t, sqlx, _), a) in enumerate(zip(SQLQ, SQLA), 1):
    tpl += [f"-- Q2.{i}  {t}", "", "", ""]
    keyl += [f"-- Q2.{i}  {t}   -> expected: {a}", sqlx, ""]
(DATA_DIR / "ETEH_Week3_Assignment_Queries.sql").write_text("\n".join(tpl), encoding="utf-8")
(KEYDIR / "ETEH_Week3_Assignment_Queries_KEY.sql").write_text("\n".join(keyl), encoding="utf-8")


def build(key: bool):
    A = Asg(3, "ETEH Week 3 Assignment - Lifecare Hospital Database: SQL, Data Model & Dashboard", key,
            "Week 3 · SQL, joins, data model, DAX and dashboards · lifecare_hospital.db")
    b = A.b

    # ---------------------------------------------------------------- function guide (self-contained help)
    # Every worksheet function the tasks need is taught in the Week 1-3 examples workbooks; the guide shows the exact patterns on this
    # workbook's own tables, each answering a DIFFERENT question from the tasks, so it teaches without giving answers away.
    g = b.sheet("Function Guide", "Function guide: the worksheet patterns this assignment uses, on its own tables", A.subtitle, "Reference")
    g.scenario("Every formula below is live and works on this workbook's tables (tAdm, tWards, tDoctors, tCensus on the D_ sheets). Each one answers "
               "a different question from the tasks: learn the pattern, then apply it. SQL and DAX patterns are on the Week 3 slides and in "
               "01_Section_Examples/Week3_Data_Files/ETEH_Week3_SQL_Examples.sql.")
    g.h2("Bring a column across from a lookup table: XLOOKUP([@key], lookup[key], lookup[column])   [Task 3]")
    beds = wards.set_index("ward_id").staffed_beds
    g.calc([("Staffed beds of the first admission's ward", "=XLOOKUP(INDEX(tAdm[ward_id],1),tWards[ward_id],tWards[staffed_beds])", int(beds[adm.ward_id.iloc[0]]),
             "int", f"Ward {adm.ward_id.iloc[0]}. Inside the table, write it once as a calculated column with [@ward_id] and Excel fills every row.")])
    g.h2("Count or average one year from a two-year table: add the year as a criterion   [Tasks 3-4]")
    g.calc([("2022 admissions", "=COUNTIFS(tAdm[adm_year],2022)", len(a22), "int", "The adm_year column makes every 2023 measure twin a COUNTIFS / SUMIFS / AVERAGEIFS."),
            ("2022 average bill", "=AVERAGEIFS(tAdm[billing_amount],tAdm[adm_year],2022)", float(a22.billing_amount.mean()), "naira",
             "Average range first, then criteria pairs."),
            ("2022 Elective share", '=COUNTIFS(tAdm[admission_type],"Elective",tAdm[adm_year],2022)/COUNTIFS(tAdm[adm_year],2022)',
             float((a22.admission_type == "Elective").mean()), "pct", "A rate = count with the condition ÷ count without it (DAX: DIVIDE(CALCULATE(...), ...)).")])
    g.h2("Count rows in one month: COUNTIFS(date, \">=\"&first day, date, \"<=\"&EOMONTH(first day, 0))   [Task 5]")
    g.calc([("Admissions in March 2022", '=COUNTIFS(tAdm[admission_date],">="&DATE(2022,3,1),tAdm[admission_date],"<="&EOMONTH(DATE(2022,3,1),0))',
             int(((adm.admission_date >= "2022-03-01") & (adm.admission_date <= "2022-03-31")).sum()), "int",
             "EOMONTH(date, 0) is the last day of that month, so the pair of criteria brackets the whole month.")])
    g.h2("Census rates: AVERAGEIFS / COUNTIFS on the daily table   [Task 5]")
    c22 = cen[cen.year == 2022]
    g.calc([("2022 days above 95% full", '=COUNTIFS(tCensus[occupancy_pct],">0.95",tCensus[year],2022)', int((c22.occupancy_pct > 0.95).sum()), "int",
             "One census row per day, so a count of rows is a count of days."),
            ("Highest 2022 occupancy", "=MAXIFS(tCensus[occupancy_pct],tCensus[year],2022)", float(c22.occupancy_pct.max()), "pct",
             "Above 100% means more patients than staffed beds that day.")])

    # ---------------------------------------------------------------- T1 database basics
    s = A.task("T1 Database Basics", "Task 1 - Read the schema", "3.1",
               "Open 01_Section_Examples/Week3_Data_Files/lifecare_hospital.db in DB Browser for SQLite (Database Structure tab). Tables: admissions, patients, doctors, wards, bed_census.")
    A.q(s, "1.1", "Primary key of the admissions table?", "admission_id", marks=1, choices=["admission_id", "patient_id", "doctor_id", "admission_date"], header=True)
    A.q(s, "1.2", "Which admissions column is a foreign key to the doctors table?", "doctor_id", marks=1, choices=["doctor_id", "department", "admission_id", "ward_id"])
    A.q(s, "1.3", "Relationship from patients to admissions?", "One-to-many", marks=2, choices=["One-to-one", "One-to-many", "Many-to-many"],
        guide="5,766 patients have 2+ admissions.")
    A.q(s, "1.4", "What is the grain (one row = ?) of bed_census?", "One calendar day", marks=2, choices=["One calendar day", "One admission", "One ward per day", "One bed"])
    A.long(s, "1.5", "Patient name, gender and blood type live in patients, not on every admission row. Give TWO reasons.",
           "• Avoid repetition: a patient with 4 admissions would have their details stored 4 times (larger file, slower).\n"
           "• Avoid update anomalies: a name correction is made once, not on every admission row - no conflicting versions.\n"
           "• Separates who the patient is (dimension) from what happened (event/fact).", marks=4, guide="2 marks per valid reason.", rows=3)
    A.end_task(s)

    # ---------------------------------------------------------------- T2 SQL
    s = A.task("T2 SQL Queries", "Task 2 - Answer the Medical Director's questions with SQL", "3.2",
               "Write each query in Week3_Assignment_Data/ETEH_Week3_Assignment_Queries.sql, run it, and enter the result in the yellow cell. "
               "Paste your SQL into the box under each question. Result 2 marks, SQL 1 mark.")
    for i, ((t, sqlx, kind), a) in enumerate(zip(SQLQ, SQLA), 1):
        gd = {6: "Dr Hadiza Uzor 515 vs Dr Mustapha Ike 514 - a one-admission margin: report both and say so.",
              3: "Endocrinology has 8 doctors and covers diabetes AND obesity - roughly double any other department."}.get(i, "")
        A.q(s, f"2.{i}", t, a.item() if hasattr(a, "item") else a, None, kind, 2, gd, header=True)
        A.long(s, "SQL", "Your SQL:", sqlx, marks=1, rows=max(3, sqlx.count("\n") + 1), guide="Any query that returns the same result is correct.")
    A.end_task(s)

    # ---------------------------------------------------------------- T3 connect / join / merge
    s = A.task("T3 Joins, Merge & Append", "Task 3 - Joining tables safely (SQL and Excel)", "3.3-3.4",
               "Sheet D_Admissions holds the 2022 and 2023 admissions (table tAdm), already APPENDED into one table. Two lookup columns have been added "
               "at its right edge (yellow): fill them with XLOOKUP calculated-column formulas, then answer.")
    A.q(s, "3.1", "Which component lets Excel's Get Data > From ODBC query a SQLite file?", "SQLite3 ODBC driver + DSN", marks=2,
        choices=["SQLite3 ODBC driver + DSN", "From Web connector", "Power Pivot only", "VBA macro"], header=True)
    A.q(s, "3.2", "An analyst joined admissions to doctors ON a.department = d.department for 2023. How many rows came back?", WRONG, None, "int", 2,
        f"{WRONG:,} rows from {RIGHT:,} admissions - each admission matched every doctor in its department.")
    A.long(s, "3.3", "Explain why that join is wrong and give the correct ON clause.",
           f"department is not a key - it repeats in both tables, so the join is many-to-many and every admission is multiplied by the number of doctors "
           f"in its department ({WRONG:,} rows instead of {RIGHT:,}; billing would be inflated about {WRONG / RIGHT:.1f}x). Correct: ON d.doctor_id = a.doctor_id "
           "(primary key = foreign key). Always check the row count after a join equals the fact-table count.", marks=3, rows=4,
           guide="Many-to-many / not a key 1; inflation consequence 1; correct ON clause 1.")
    s.grid([["ward", "'=XLOOKUP([@ward_id],tWards[ward_id],tWards[ward])" if key else "?", "Ward name from tWards"],
            ["doctor_grade", "'=XLOOKUP([@doctor_id],tDoctors[doctor_id],tDoctors[grade])" if key else "?", "Consultant / Registrar ... from tDoctors"]],
           header=["New column in tAdm", "Formula", "Meaning"], merge_last=True)
    A.q(s, "3.4", "Both lookup columns completed as calculated columns (instructor inspects D_Admissions).", "2 columns" if key else "", marks=2, header=True)
    A.q(s, "3.5", "2023 admissions to Topaz Ward (formula using your new ward column).", '=COUNTIFS(tAdm[ward],"Topaz Ward",tAdm[adm_year],2023)',
        int(((adm_ward == "Topaz Ward") & (adm.adm_year == 2023)).sum()), "int", 2)
    top_grade = adm_grade[adm.adm_year == 2023].value_counts()
    A.q(s, "3.6", "2023 admissions handled by Consultants.", '=COUNTIFS(tAdm[doctor_grade],"Consultant",tAdm[adm_year],2023)',
        int(top_grade.get("Consultant", 0)), "int", 2)
    A.q(s, "3.7", "Append check: rows in tAdm minus (2022 rows + 2023 rows) from SQL. Expected?", f'=ROWS(tAdm)-({len(a22)}+{len(a23)})', 0, "int", 2,
        f"2022 = {len(a22):,}, 2023 = {len(a23):,}. Must be 0 - an append adds rows, never changes them.")
    A.end_task(s)

    # ---------------------------------------------------------------- T4 data model & DAX
    s = A.task("T4 Data Model & DAX", "Task 4 - Star schema and core measures", "3.5-3.6",
               "Add tAdm, tWards, tDoctors and a DimDate table (CALENDAR(DATE(2022,1,1),DATE(2023,12,31))) to the Data Model, build the relationships, "
               "and write the measures. Enter the 2023 value each measure returns (PivotTable filtered to 2023). The key re-computes each value with worksheet formulas.")
    for i, (tb, ans) in enumerate([("tAdm (admissions)", "Fact"), ("tWards", "Dimension"), ("tDoctors", "Dimension"), ("DimDate", "Dimension"),
                                   ("tCensus (one row per day)", "Fact")], 1):
        A.q(s, f"4.{i}", f"Classify: {tb}", ans, marks=1, choices=["Fact", "Dimension"], header=(i == 1),
            guide="Census is a periodic-snapshot fact (measures per day) related to DimDate." if i == 5 else "")
    A.q(s, "4.6", "Relationship tAdm[ward_id] → tWards[ward_id]: cardinality and filter direction?", "Many-to-one, single", marks=2,
        choices=["Many-to-one, single", "One-to-one, both", "Many-to-many, both", "One-to-many, single"])
    r23 = (a23.readmit_30d == "Yes").mean()
    rev23, rev22 = a23.billing_amount.sum(), a22.billing_amount.sum()
    em23 = (a23.admission_type == "Emergency").mean()
    meas = [
        ("Total Admissions", "Total Admissions := COUNTROWS ( tAdm )", "=COUNTIFS(tAdm[adm_year],2023)", len(a23), "int"),
        ("ALOS", "ALOS := AVERAGE ( tAdm[length_of_stay] )", "=AVERAGEIFS(tAdm[length_of_stay],tAdm[adm_year],2023)", a23.length_of_stay.mean(), "dec2"),
        ("Readmission Rate", 'Readmission Rate := DIVIDE ( CALCULATE ( [Total Admissions], tAdm[readmit_30d] = "Yes" ), [Total Admissions] )',
         '=COUNTIFS(tAdm[readmit_30d],"Yes",tAdm[adm_year],2023)/COUNTIFS(tAdm[adm_year],2023)', r23, "pct2"),
        ("Revenue", "Revenue := SUM ( tAdm[billing_amount] )", "=SUMIFS(tAdm[billing_amount],tAdm[adm_year],2023)", rev23, "naira"),
        ("Revenue YoY %", "Revenue YoY % := VAR PY = CALCULATE ( [Revenue], SAMEPERIODLASTYEAR ( DimDate[Date] ) )\n                 RETURN DIVIDE ( [Revenue] - PY, PY )",
         "=SUMIFS(tAdm[billing_amount],tAdm[adm_year],2023)/SUMIFS(tAdm[billing_amount],tAdm[adm_year],2022)-1", rev23 / rev22 - 1, "pct2"),
        ("Emergency Share", 'Emergency Share := DIVIDE ( CALCULATE ( [Total Admissions], tAdm[admission_type] = "Emergency" ), [Total Admissions] )',
         '=COUNTIFS(tAdm[admission_type],"Emergency",tAdm[adm_year],2023)/COUNTIFS(tAdm[adm_year],2023)', em23, "pct"),
    ]
    s.band("DAX", "Measures - DAX 1 mark, 2023 value 2 marks")
    for i, (nm, dax, fx, v, k) in enumerate(meas, 7):
        A.q(s, f"4.{i}", f"Measure '{nm}' - 2023 value", fx, v, k, 2, "", header=(i == 7))
        A.long(s, "DAX", f"Your DAX for {nm}:", dax, marks=1, rows=2)
    A.end_task(s)

    # ---------------------------------------------------------------- T5 charts & dashboard
    s = A.task("T5 Charts & Dashboard", "Task 5 - One-page operations dashboard", "3.7-3.9",
               "Choose the right chart for each question, then build a one-page dashboard on a new sheet My_Dashboard (KPI cards + 2 charts + a Ward slicer).")
    ch = ["Line", "Clustered bar (sorted)", "100% stacked column", "Scatter", "Pie with 12 slices"]
    for i, (t, a) in enumerate([("Monthly admissions, 2022 vs 2023", "Line"), ("Revenue by department, largest first", "Clustered bar (sorted)"),
                                ("Admission-type mix (Emergency/Urgent/Elective) for each ward", "100% stacked column"),
                                ("Does age relate to length of stay?", "Scatter")], 1):
        A.q(s, f"5.{i}", t, a, marks=1, choices=ch, header=(i == 1))
    occ23 = cen[cen.year == 2023].occupancy_pct.mean()
    occ22 = cen[cen.year == 2022].occupancy_pct.mean()
    A.q(s, "5.5", "KPI card: 2023 average occupancy (formula on tCensus).", '=AVERAGEIFS(tCensus[occupancy_pct],tCensus[year],2023)', occ23, "pct", 2, header=True)
    A.q(s, "5.6", "KPI card: change in average occupancy vs 2022 (percentage points, as a %).",
        '=AVERAGEIFS(tCensus[occupancy_pct],tCensus[year],2023)-AVERAGEIFS(tCensus[occupancy_pct],tCensus[year],2022)', occ23 - occ22, "pct2", 2)
    A.q(s, "5.7", "Dashboard built: 4 KPI cards, 2 correct charts, Ward slicer connected to both (instructor inspects My_Dashboard).",
        "Rubric" if key else "", marks=4, guide="KPIs 1, charts 1 (titles are findings, no 3-D, sorted bars), slicer drives both 1, layout/alignment 1.")
    A.long(s, "5.8", "Write the dashboard's headline and one recommendation for the Medical Director.",
           f"e.g. 'Beds ran at {occ23:.1%} average occupancy in 2023 - above the 85% safe-occupancy guide - while admissions held at {len(a23):,}.' "
           "Recommendation: review discharge planning / open surge capacity in peak months; track readmission alongside LOS so faster discharge does not raise readmissions.",
           marks=3, rows=3, guide="Headline states a finding with a number 2; actionable recommendation 1.")
    if key:
        s.band("KEY", "Reference dashboard elements")
        months = pd.date_range("2022-01-01", periods=24, freq="MS")
        md = pd.DataFrame({"Month": months.strftime("%b").tolist()[:12]})
        cnt = {y: [int(((adm.adm_year == y) & (adm.admission_date.dt.month == m)).sum()) for m in range(1, 13)] for y in (2022, 2023)}
        top = s.r + 1
        s.frame(pd.DataFrame({"Month": md.Month, "2022": cnt[2022], "2023": cnt[2023]}),
                formulas={(i, c): (f'=COUNTIFS(tAdm[admission_date],">="&DATE({c},{i + 1},1),tAdm[admission_date],"<="&EOMONTH(DATE({c},{i + 1},1),0))', cnt[int(c)][i]) for i in range(12) for c in ("2022", "2023")},
                fmts={"2022": "int", "2023": "int"})
        s.chart("line", [{"name": "2022", "categories": [s.name, top, 1, top + 11, 1], "values": [s.name, top, 2, top + 11, 2]},
                         {"name": "2023", "categories": [s.name, top, 1, top + 11, 1], "values": [s.name, top, 3, top + 11, 3]}],
                "Admissions are steady at ~900 a month in both years", y_title="Admissions")
    A.end_task(s)

    # ---------------------------------------------------------------- T6 files to SQL
    s = A.task("T6 Files to SQL", "Task 6 - From CSV files to a queryable database", "3.10",
               "The 2025 admissions will arrive as monthly CSV files. Explain how you would load them into lifecare_hospital.db and keep the data model working.")
    A.q(s, "6.1", "In DB Browser for SQLite, which menu imports a CSV into a new table?", "File > Import > Table from CSV file", marks=1,
        choices=["File > Import > Table from CSV file", "Execute SQL > Run", "Edit > Preferences", "Tools > Compact Database"], header=True)
    A.q(s, "6.2", "Which SQL statement adds the imported rows into the existing admissions table?", "INSERT INTO admissions SELECT ... FROM staging", marks=2,
        choices=["INSERT INTO admissions SELECT ... FROM staging", "CREATE VIEW admissions AS ...", "UPDATE admissions SET ...", "DROP TABLE admissions"])
    A.long(s, "6.3", "List two checks you would run after loading a month's file.",
           "• Row count in the CSV = rows inserted (SELECT COUNT(*) ... WHERE admission_date in the month).\n"
           "• No orphan keys: LEFT JOIN doctors/wards and count rows where the dimension key IS NULL.\n"
           "• No duplicate admission_id (PRIMARY KEY violation) and control total of billing matches Finance.", marks=2, rows=3, guide="1 mark per valid check.")
    A.end_task(s)

    # ---------------------------------------------------------------- data sheets
    ds = b.sheet("D_Admissions", "Lifecare admissions 2022-2023 (table tAdm) - appended from two yearly extracts",
                 f"{len(adm):,} rows · one row per admission · billing in ₦", "Data sheet")
    ds.r = 4
    A.build_table(ds, adm, "tAdm", {"ward": ("=XLOOKUP([@ward_id],tWards[ward_id],tWards[ward])", "gen", adm_ward.tolist()),
                                    "doctor_grade": ("=XLOOKUP([@doctor_id],tDoctors[doctor_id],tDoctors[grade])", "gen", adm_grade.tolist())},
                  fmts={"admission_date": "date", "discharge_date": "date", "billing_amount": "naira2", "adm_year": "0"}, style="Table Style Light 9",
                  widths={"medical_condition": 14, "department": 24, "hmo": 26})
    ds.ws.freeze_panes(5, 0)
    b.data_sheet("D_Wards", wards, "tWards", title="Wards (table tWards)")
    b.data_sheet("D_Doctors", docs, "tDoctors", title="Doctors (table tDoctors)", widths={"doctor_name": 24, "department": 26})
    b.data_sheet("D_Census", cen, "tCensus", fmts={"census_date": "date", "occupancy_pct": "pct", "year": "0"}, title="Daily bed census 2022-2023 (table tCensus)")

    return A.finish("Week 3 Assignment - Lifecare Hospital Database",
                    ["Scenario: Lifecare General Hospital has moved its records into a SQLite database. The Medical Director wants answers straight from the database "
                     "(Task 2), a safe way to combine tables (Task 3), a Power Pivot model with standard measures (Task 4) and a one-page dashboard (Task 5).",
                     "You need DB Browser for SQLite (free) and Excel with Power Pivot. Time guide: 7-8 hours."],
                    [("Lifecare", DATASETS_TEXT["Lifecare"])],
                    files="Function Guide sheet: the worksheet patterns the tasks use. Database: 01_Section_Examples/Week3_Data_Files/lifecare_hospital.db (shared with the Section Examples). Query template: 02_Assignments/Week3_Assignment_Data/ETEH_Week3_Assignment_Queries.sql.",
                    rubric=[["SQL", "Any query returning the correct result earns the SQL mark; style (aliases, indentation) is commented on, not penalised."],
                            ["DAX", "Equivalent DAX accepted (e.g. COUNTA, FILTER inside CALCULATE)."],
                            ["Tolerance", "Counts exact; money ±₦1; rates ±0.01 pt."]])


if __name__ == "__main__":
    for k in (False, True):
        print("Week 3", "KEY" if k else "student", "->", build(k))
