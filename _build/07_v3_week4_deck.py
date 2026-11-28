"""Edition 3: build the Week 4 deck.

    python3 _build/07_v3_week4_deck.py [--preview]

Engine: _build/v3deck.py (template _build/templates/ETEH_deck_template.pptx). Copy lives in _build/v3/w4_*.py; example tables come
from _build/w4_data.py and every figure is computed in _build/v3/w4_fx.py or in the copy files from those tables.
"""
from __future__ import annotations

import importlib
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE / "v3"))
import v3deck  # noqa: E402
import w4_data  # noqa: E402
from w4_extra import GLOSSARY, SUMMARY  # noqa: E402

PARTS = ["w4_m1_2", "w4_m3_4", "w4_m5_6", "w4_m7_10", "w4_m11_14"]
MODULES = []
for p in PARTS:
    try:
        MODULES += importlib.import_module(p).MODULES
    except ModuleNotFoundError as e:
        if e.name != p:
            raise
        print(f"(skipping {p}: not written yet)")

SHEETS = {"4.1": "4.1 Descriptive Stats", "4.2": "4.2 Distributions & Z", "4.3": "4.3 Correlation", "4.4": "4.4 Regression",
          "4.5": "4.5 Forecasting", "4.6": "4.6 What-If & Solver", "4.7": "4.7 Storytelling", "4.8": "4.8 Report Writing",
          "4.9": "4.9 Presentations", "4.10": "4.10 Audit & Protect", "4.11": "4.11 AI Co-pilots", "4.12": "4.12 Verify AI",
          "4.13": "4.13 Ethics PII PHI", "4.14": "4.14 Benchmarks"}

import w4_charts  # noqa: E402
print(f"charts on {w4_charts.attach(MODULES)} topics")

CFG = v3deck.Config(
    week=4,
    week_title="Week 4: Statistics, Forecasting, What-If, Communication, AI & Ethics",
    title_sub="Describe and model data honestly, forecast and test decisions, communicate the result, and work safely with AI and personal data",
    out=v3deck.ROOT / "ETEH_Bootcamp_Pack" / "ETEH_Vol4_Week4_Stats_Forecast_AI_Slides.pptx",
    workbook="ETEH_Week4_Section_Examples.xlsx",
    modules=MODULES,
    sheets=SHEETS,
    data=w4_data,
    summary=SUMMARY,
    glossary=GLOSSARY,
    roadmap_label="FOURTEEN MODULES THIS WEEK",
    roadmap_notes="Walk through the modules. 4.1-4.5 are statistics and forecasting; 4.6 is what-if analysis; 4.7-4.9 are "
                  "communication; 4.10 is auditing and automation; 4.11-4.13 are AI and ethics; 4.14 puts it together.",
    title_notes="Welcome students. Weeks 1-3 built clean data, models and dashboards. This week we describe and model data with "
                "statistics, forecast and test decisions, communicate the findings, and work responsibly with AI and personal data.",
    assignment="ETEH_Week4_Assignment.xlsx",
    volume="Volume 4 of 5",
    modules_text="14 modules",
    example_table="tSal",
)

if __name__ == "__main__":
    v3deck.build(CFG)
