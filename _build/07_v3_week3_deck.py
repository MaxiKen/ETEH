"""Edition 3: build the Week 3 deck.

    python3 _build/07_v3_week3_deck.py [--preview]

Engine: _build/v3deck.py (template _build/templates/ETEH_deck_template.pptx). Copy lives in _build/v3/w3_*.py;
example tables and every SQL result come from _build/w3_data.py (the SQL printed on a slide is the SQL that was run on
01_Section_Examples/Week_3/Ex3.2_Files: supermart_online.db or mercyland_hospital.db); other figures are computed in _build/v3/w3_fx.py.
"""
from __future__ import annotations

import importlib
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE / "v3"))
import v3deck  # noqa: E402
import w3_data  # noqa: E402
from w3_extra import GLOSSARY, SUMMARY  # noqa: E402

PARTS = ["w3_m1_2", "w3_m3_5", "w3_m6a", "w3_m6b", "w3_m7", "w3_m8_10"]
MODULES = []
for p in PARTS:
    try:
        MODULES += importlib.import_module(p).MODULES
    except ModuleNotFoundError as e:
        if e.name != p:
            raise
        print(f"(skipping {p}: not written yet)")

SHEETS = {"3.1": "3.1 Database Basics", "3.2": "3.2 SQL with SQLite", "3.3": "3.3 ODBC Connections",
          "3.4": "3.4 Merge & Append", "3.5": "3.5 Data Model & Star", "3.6A": "3.6A Data Model Pivots",
          "3.6B": "3.6B KPIs & Time", "3.7": "3.7 Chart Types", "3.8": "3.8 Dashboard Build",
          "3.9": "3.9 Blueprints", "3.10": "3.10 Files to Dashboards"}

CFG = v3deck.Config(
    week=3,
    week_title="Week 3: Databases, Power Query, Data Modelling & Dashboards",
    title_sub="Explore databases, combine tables in Power Query, model them in a star, analyse with PivotTables and build dashboards that reconcile",
    out=v3deck.ROOT / "ETEH_Bootcamp_Pack" / "ETEH_Vol3_Week3_Databases_PowerQuery_Dashboards_Slides.pptx",
    workbook="ETEH_Week3_Section_Examples.xlsx",
    modules=MODULES,
    sheets=SHEETS,
    data=w3_data,
    summary=SUMMARY,
    glossary=GLOSSARY,
    roadmap_label="TEN MODULES THIS WEEK (MODULE 3.6 IN TWO PARTS)",
    roadmap_notes="Walk through the modules. 3.1-3.2 show how a database is built and let students try simple queries in DB Browser; 3.3 brings the data into Excel; 3.4-3.5 combine tables in Power Query and "
                  "model them as a star; 3.6 analyses the model with PivotTables (no DAX); 3.7-3.9 are charts and dashboards; 3.10 puts it all together.",
    title_notes="Welcome students. Weeks 1 and 2 worked on files; this week the data lives in databases. We look at it "
                "with a little SQL, combine it in Power Query, model it in a star schema, analyse it with PivotTables and present "
                "it on dashboards that reconcile.",
    assignment="ETEH_Week3_Assignment.xlsx",
    volume="Volume 3 of 5",
    modules_text="10 modules (Module 3.6 in two parts)",
    example_table="tTransactions",
)

def write_sql_examples():
    """ETEH_Week3_SQL_Examples.sql = exactly the SQL panels on the slides (copied to Ex3.2_Files by the Week 3 workbook)."""
    out = ["-- ETEH Week 3 - SQL examples (SQLite), edition 3.3: every query shown on the Week 3 slides.",
           "-- Run each against the database named above it (both databases are in this Ex3.2_Files folder) in DB Browser for "
           "SQLite (Execute SQL tab).",
           "-- The result printed on the slide is what the query returns on these files.", ""]
    for m in MODULES:
        for t in m["topics"] + [x for t in m["topics"] for x in t.get("more", [])]:
            if t.get("code") and str(t.get("code_label", "")).startswith("SQL"):
                db = w3_data.DBNAME.get(t.get("code_db"), "the staging table stg_admissions_2023 (load it as in 3.10.2)")
                out += [f"-- {t.get('tag', m['id'])} {t['title']}", f"-- database: {db}", t["code"].rstrip().rstrip(";") + ";", ""]
    path = w3_data.W3D / "ETEH_Week3_SQL_Examples.sql"
    path.write_text("\n".join(out), encoding="utf-8")
    print(f"{sum(1 for x in out if x.startswith('-- database'))} SQL panels -> {path.name}")


if __name__ == "__main__":
    v3deck.build(CFG)
    write_sql_examples()
