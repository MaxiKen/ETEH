"""Edition 3: build the Week 2 deck.

    python3 _build/07_v3_week2_deck.py [--preview]

Engine: _build/v3deck.py (template _build/templates/ETEH_deck_template.pptx). Copy lives in _build/v3/w2_*.py;
example tables come from _build/w2_data.py and every figure is computed in _build/v3/w2_fx.py.
"""
from __future__ import annotations

import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE / "v3"))
import v3deck  # noqa: E402
import w2_data  # noqa: E402
from w2_extra import GLOSSARY, SUMMARY  # noqa: E402
from w2_m1_4 import MODULES as MA  # noqa: E402
from w2_m5_7 import MODULES as MB  # noqa: E402
from w2_m8 import MODULES as MC  # noqa: E402
from w2_m9_11 import MODULES as MD  # noqa: E402

import w2_charts  # noqa: E402
print(f"charts on {w2_charts.attach(MD)} topics")

CFG = v3deck.Config(
    week=2,
    week_title="Week 2: Data Import, Power Query Cleaning & PivotTable Mastery",
    title_sub="Connect, transform, calculate with 65 functions, summarise and visualise",
    out=v3deck.ROOT / "ETEH_Bootcamp_Pack" / "ETEH_Vol2_Week2_PowerQuery_Pivots_Slides.pptx",
    workbook="ETEH_Week2_Section_Examples.xlsx",
    modules=MA + MB + MC + MD,
    sheets={"2.1": "2.1 External Formats", "2.2": "2.2 Power Query Intro", "2.3": "2.3 Connect & Combine",
            "2.4": "2.4 Web & PDF Sources", "2.5": "2.5 Excel Cleaning", "2.6": "2.6 PQ Transforms", "2.7": "2.7 Excel vs PQ",
            "2.8A": "2.8A Functions 1-33", "2.8B": "2.8B Functions 34-65", "2.9": "2.9 M Language",
            "2.10": "2.10 PivotTables", "2.11": "2.11 PivotCharts"},
    data=w2_data,
    summary=SUMMARY,
    glossary=GLOSSARY,
    roadmap_label="ELEVEN MODULES THIS WEEK (MODULE 2.8 IN TWO PARTS)",
    roadmap_notes="Walk through the modules. 2.1-2.4 bring data in; 2.5-2.7 clean it; 2.8 is the function toolkit; "
                  "2.9 is the M language behind Power Query; 2.10-2.11 summarise and chart with pivots.",
    title_notes="Welcome students. Week 1 built tidy tables by hand; this week we bring data in from outside, clean it "
                "repeatably with Power Query, and summarise it with functions and PivotTables.",
    assignment="ETEH_Week2_Assignment.xlsx",
    volume="Volume 2 of 5",
    modules_text="11 modules (Module 2.8 in two parts)",
    example_table="tSales25",
)

if __name__ == "__main__":
    v3deck.build(CFG)
