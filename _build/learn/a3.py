"""Week 3 assignment (Edition 3.1): Mercyland Hospital database. Week 1-3 skills only.

    Task 1  database design: keys, relationships, normal forms (3.1)
    Task 2  SQL on mercyland_hospital.db in DB Browser for SQLite (3.2)
    Task 3  merge and the Data Model (3.4 - 3.5)
    Task 4  DAX measures, each checked with a Week 2 formula (3.6A - 3.6B)
    Task 5  charts and dashboard design (3.7 - 3.9)
"""
from __future__ import annotations

import sqlite3

import pandas as pd

from common import W3D

DB = "mercyland_hospital.db"


def build(b):
    con = sqlite3.connect(W3D / DB)

    def v(sql):
        return con.execute(sql).fetchone()[0]

    adm = pd.read_sql("SELECT episode_id, patient_no, admit_date, department, admission_route, ward_code, consultant_id, "
                      "los_days, bill_amount_ngn, readmitted_30d FROM admissions WHERE admit_date LIKE '2023%' "
                      "ORDER BY admit_date, episode_id", con)
    adm["admit_date"] = pd.to_datetime(adm.admit_date)
    wards = pd.read_sql("SELECT ward_code, ward, open_beds FROM wards", con)
    docs = pd.read_sql("SELECT consultant_id, consultant_name, department, grade FROM doctors", con)
    dd = pd.DataFrame({"Date": pd.date_range("2023-01-01", "2023-12-31")})
    dd["Month_No"] = dd.Date.dt.month
    dd["Month"] = dd.Date.dt.strftime("%b")
    dd["Quarter"] = "Q" + dd.Date.dt.quarter.astype(str)

    t1 = b.module("T1", "T1 Database Design", "Database design", "Module 3.1: tables, primary and foreign keys, "
                  "relationships, normal forms, views and the star schema.")
    t2 = b.module("T2", "T2 SQL", "SQL queries", "Module 3.2: SELECT, WHERE, ORDER BY, LIMIT, GROUP BY, HAVING, "
                  "COUNT / SUM / AVG, DISTINCT, JOIN, CASE and date functions.")
    t3 = b.module("T3", "T3 Merge & Model", "Merge and the Data Model", "Modules 3.4 - 3.5: Power Query merge and join "
                  "kinds, row-count checks, the Data Model, relationships and the date table.")
    t4 = b.module("T4", "T4 DAX", "DAX measures", "Modules 3.6A - 3.6B: measures, CALCULATE, DIVIDE, DISTINCTCOUNT, "
                  "RELATED and time intelligence.")
    t5 = b.module("T5", "T5 Dashboard", "Charts and dashboard design", "Modules 3.7 - 3.9: choosing charts, dashboard "
                  "layout, KPI cards, slicers and the QA checklist.")
    f = t1.link_files([W3D / "mercyland_hospital_schema.sql", W3D / DB], folder="Assignment_Files")
    t2.link_files([W3D / "mercyland_hospital_schema.sql", W3D / DB], folder="Assignment_Files")
    r_a = b.data("Adm23", adm, "One Mercyland admission in 2023 (from the admissions table of mercyland_hospital.db)",
                 fmts={"admit_date": "date", "bill_amount_ngn": "naira2"}, used_in="T3, T4")
    r_w = b.data("Wards", wards, "One ward with its open beds", used_in="T3, T4")
    r_d = b.data("Doctors", docs, "One consultant", used_in="T3")
    r_dd = b.data("Date23", dd, "One calendar day of 2023 (the date table)", fmts={"Date": "date"}, used_in="T3, T4")
    t3.set_refs([r_a, r_w, r_d, r_dd])
    t4.set_refs([r_a, r_w, r_dd])

    # ================================================================== T1
    m = t1
    m.info("Scenario", "Mercyland's Medical Director wants a 2023 activity pack built from the hospital database: how many "
                       "patients, which wards and departments are busiest, how long patients stay and how many come back "
                       "within 30 days.")
    m.part("Read the schema")
    m.do(f"Open {f} › mercyland_hospital_schema.sql in Notepad. What is the primary key of the admissions table?",
         "episode_id", "episode_id TEXT PRIMARY KEY.")
    m.mc("In admissions, ward_code REFERENCES wards(ward_code). What is ward_code in admissions?",
         ["A primary key", "A foreign key", "A view", "An index"], "B", "It points at the wards table.")
    m.pick("What is the relationship between wards and admissions?", ["One-to-many", "One-to-one", "Many-to-many"],
           "One-to-many", "One ward, many admissions.")
    m.pick("admissions (one row per hospital stay, with bill and length of stay)", ["Fact table", "Dimension table"],
           "Fact table", "Events with numbers.")
    m.pick("doctors (one row per consultant)", ["Fact table", "Dimension table"], "Dimension table", "Describes who.")
    m.mc("Why is patient_name stored in patients and not repeated on every admission?",
         ["To save colours", "Normalisation: store each fact once, so a name change is made in one place",
          "SQLite cannot store names twice", "No reason"], "B", "Avoids update errors.")
    m.yn("Is v_admissions a view (a saved query joining the tables)?", "Yes", "CREATE VIEW v_admissions AS SELECT ...")

    # ================================================================== T2
    m = t2

    def sql(text, q, fmt="int"):
        val = v(q)
        if isinstance(val, str):
            m.short(text + "\n" + q, val, q)
        else:
            m.num(text + "\n" + q, None, val, "Query: " + q, fmt=fmt, do=True)

    m.info("How", f"Open DB Browser for SQLite › Open Database › {f} › {DB} › Execute SQL. Each question gives the query to "
                  "run (some you finish yourself). Type the number the query returns.")
    sql("How many admissions are in the database?", "SELECT COUNT(*) FROM admissions;")
    sql("How many admissions were in 2023?", "SELECT COUNT(*) FROM admissions WHERE strftime('%Y', admit_date) = '2023';")
    sql("How many 2023 admissions were in Cardiology?",
        "SELECT COUNT(*) FROM admissions WHERE admit_date LIKE '2023%' AND department = 'Cardiology';")
    sql("How many different patients were admitted in 2023?",
        "SELECT COUNT(DISTINCT patient_no) FROM admissions WHERE admit_date LIKE '2023%';")
    sql("Which department had the most 2023 admissions? (GROUP BY, ORDER BY, LIMIT 1)",
        "SELECT department FROM admissions WHERE admit_date LIKE '2023%' GROUP BY department ORDER BY COUNT(*) DESC LIMIT 1;")
    sql("How many open beds does the hospital have in total?", "SELECT SUM(open_beds) FROM wards;")
    sql("JOIN: how many 2023 admissions were in Iroko Ward?",
        "SELECT COUNT(*) FROM admissions a JOIN wards w ON w.ward_code = a.ward_code "
        "WHERE a.admit_date LIKE '2023%' AND w.ward = 'Iroko Ward';")
    sql("How many consultants have more than 600 admissions in 2023? (GROUP BY ... HAVING)",
        "SELECT COUNT(*) FROM (SELECT consultant_id FROM admissions WHERE admit_date LIKE '2023%' "
        "GROUP BY consultant_id HAVING COUNT(*) > 600);")
    sql("How many 2023 stays were long (los_days of 14 or more)?",
        "SELECT COUNT(*) FROM admissions WHERE admit_date LIKE '2023%' AND los_days >= 14;")

    # ================================================================== T3
    m = t3
    m.part("Part A · Merge in Power Query")
    m.num("Merge tAdm23 with tWards on ward_code (Left Outer) and expand ward. How many rows does the result have?", None,
          len(adm), "Each ward_code appears once in tWards, so the row count does not change.", do=True)
    m.yn("If the row count went UP after a merge, would that be a warning sign?", "Yes",
         "Duplicate keys on the right side multiply rows.")
    m.pick("Find doctors with no 2023 admissions: merge tDoctors with tAdm23 using ...",
           ["Inner", "Left Outer", "Left Anti", "Full Outer"], "Left Anti", "Left Anti keeps unmatched left rows.")
    m.part("Part B · The Data Model")
    m.mc("In the Data Model, which column links tAdm23 to tWards?", ["ward_code", "episode_id", "los_days", "admit_date"],
         "A", "Relationship tWards[ward_code] (one) → tAdm23[ward_code] (many).")
    m.mc("Which column links tAdm23 to tDate23?", ["admit_date → Date", "episode_id → Date", "ward_code → Month",
                                                  "los_days → Quarter"], "A", "Mark tDate23 as the date table.")
    m.num("How many days are in tDate23?", "=ROWS(tDate23)", len(dd), "One row per day of 2023.")

    # ================================================================== T4
    m = t4
    A = "tAdm23"
    n = len(adm)
    rr = float((adm.readmitted_30d == "Yes").mean())
    m.part("Write each measure; check it with the Week 2 formula shown in the Solutions")
    m.num("Admissions := COUNTROWS(tAdm23). Value?", f"=ROWS({A})", n, "Excel check: =ROWS(tAdm23).")
    m.num('Readmission % := DIVIDE(CALCULATE([Admissions], tAdm23[readmitted_30d] = "Yes"), [Admissions]). Value?',
          f'=COUNTIFS({A}[readmitted_30d],"Yes")/ROWS({A})', rr, "Readmitted ÷ admissions.", fmt="pct")
    m.num("Avg LOS := AVERAGE(tAdm23[los_days]). Value? (1 decimal place)", f"=AVERAGE({A}[los_days])",
          float(adm.los_days.mean()), "AVERAGE.", fmt="dec1")
    m.num('Total Bills := SUM(tAdm23[bill_amount_ngn]). Value for Cardiology? (nearest naira)',
          f'=ROUND(SUMIFS({A}[bill_amount_ngn],{A}[department],"Cardiology"),0)',
          round(float(adm.loc[adm.department == "Cardiology", "bill_amount_ngn"].sum())), "SUMIFS by department.",
          fmt="naira")
    m.num("Patients := DISTINCTCOUNT(tAdm23[patient_no]). Value?", f"=COUNTA(UNIQUE({A}[patient_no]))",
          int(adm.patient_no.nunique()), "Excel check: =COUNTA(UNIQUE(...)).", dynamic=True)
    q1 = int((adm.admit_date < "2023-04-01").sum())
    m.num("Admissions YTD := TOTALYTD([Admissions], tDate23[Date]). Value at 31 March 2023?",
          f'=COUNTIFS({A}[admit_date],"<"&DATE(2023,4,1))', q1, "January to March.")
    m.mc("Which DAX function brings ward from tWards into a calculated column of tAdm23?", ["RELATED", "ALL", "SUMX",
                                                                                           "DATEADD"], "A",
         "RELATED(tWards[ward]).")

    # ================================================================== T5
    m = t5
    C = ["Line", "Clustered bar", "Pie", "KPI card"]
    m.part("Choose and design")
    m.pick("Admissions per month in 2023.", C, "Line", "A trend over time.")
    m.pick("Admissions by department, largest first.", C, "Clustered bar", "Compare categories, sorted.")
    m.pick("The single readmission % for 2023.", C, "KPI card", "One big number.")
    m.mc("Where should the most important KPI go on the dashboard?", ["Top left", "Bottom right", "Hidden sheet",
                                                                      "In the footer"], "A", "People read top-left first.")
    m.yn("Should one Department slicer filter every visual on the dashboard?", "Yes", "Report Connections.")
    m.tf("Before sharing, check the dashboard totals against the SQL results.", True, "Reconcile.")
    con.close()
