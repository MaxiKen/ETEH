"""Edition 3: build the Week 3 deck.

    python3 _build/07_v3_week3_deck.py [--preview]

Engine: _build/v3deck.py (template _build/templates/ETEH_deck_template.pptx). Copy lives in _build/v3/w3_*.py;
example tables and every SQL result come from _build/w3_data.py (the SQL printed on a slide is the SQL that was run on
Week3_Data_Files/supermart_online.db or mercyland_hospital.db); other figures are computed in _build/v3/w3_fx.py.
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
          "3.4": "3.4 Merge & Append", "3.5": "3.5 Data Model & Star", "3.6A": "3.6A DAX Foundations",
          "3.6B": "3.6B DAX Time & KPIs", "3.7": "3.7 Chart Types", "3.8": "3.8 Dashboard Build",
          "3.9": "3.9 Blueprints", "3.10": "3.10 Files to SQL & DAX"}

CFG = v3deck.Config(
    week=3,
    week_title="Week 3: SQL, Data Modelling, DAX & Dashboard Development",
    title_sub="Query databases, model data in a star, write DAX measures and build dashboards that reconcile",
    out=v3deck.ROOT / "ETEH_Bootcamp_Pack" / "ETEH_Vol3_Week3_SQL_DAX_Dashboards_Slides.pptx",
    workbook="ETEH_Week3_Section_Examples.xlsx",
    modules=MODULES,
    sheets=SHEETS,
    data=w3_data,
    summary=SUMMARY,
    glossary=GLOSSARY,
    roadmap_label="TEN MODULES THIS WEEK (MODULE 3.6 IN TWO PARTS)",
    roadmap_notes="Walk through the modules. 3.1-3.3 are databases, SQL and connecting Excel; 3.4-3.5 combine tables and "
                  "model them as a star; 3.6 is DAX; 3.7-3.9 are charts and dashboards; 3.10 puts it all together.",
    title_notes="Welcome students. Weeks 1 and 2 worked on files; this week the data lives in databases. We query it with "
                "SQL, model it in a star schema, measure it with DAX and present it on dashboards that reconcile.",
    assignment="ETEH_Week3_Assignment.xlsx",
    volume="Volume 3 of 5",
    modules_text="10 modules (Module 3.6 in two parts)",
    example_table="tTransactions",
)

if __name__ == "__main__":
    v3deck.build(CFG)
