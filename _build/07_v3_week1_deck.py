"""Edition 3: build the Week 1 deck.

    python3 _build/07_v3_week1_deck.py [--preview]

Engine: _build/v3deck.py (template _build/templates/ETEH_deck_template.pptx). Copy lives in _build/v3/w1_*.py;
example tables come from _build/w1_data.py. Previews go to _build/_stage/v3_preview/NNN.png.
"""
from __future__ import annotations

import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE / "v3"))
import v3deck  # noqa: E402
import w1_data  # noqa: E402
from w1_extra import GLOSSARY, SUMMARY  # noqa: E402
from w1_m1_4 import MODULES as MA  # noqa: E402
from w1_m5_6 import MODULES as MB  # noqa: E402
from w1_m7_9 import MODULES as MC  # noqa: E402

CFG = v3deck.Config(
    week=1,
    week_title="Week 1: Data Analytics Foundations & Core Microsoft Excel Mastery",
    title_sub="Analytics thinking, the Excel workspace, formulas, Excel tables and tidy data",
    out=v3deck.ROOT / "ETEH_Bootcamp_Pack" / "ETEH_Vol1_Week1_Foundations_Slides.pptx",
    workbook="ETEH_Week1_Section_Examples.xlsx",
    modules=MA + MB + MC,
    sheets={"1.1": "1.1 Intro to Analytics", "1.2": "1.2 Types of Analytics", "1.3": "1.3 Analytics Process",
            "1.4": "1.4 Industry Applications", "1.5": "1.5 Excel Interface", "1.6": "1.6 Working with Data",
            "1.7": "1.7 Formulas & Functions", "1.8": "1.8 Excel Tables", "1.9": "1.9 Tidy Data Laws"},
    data=w1_data,
    summary=SUMMARY,
    glossary=GLOSSARY,
    roadmap_label="NINE MODULES THIS WEEK",
    roadmap_notes="Walk through the nine modules. Modules 1.1-1.4 are thinking skills; 1.5-1.9 are hands-on Excel.",
    title_notes="Welcome students. This week builds the foundations every later week depends on.",
    assignment="ETEH_Week1_Assignment.xlsx",
)

if __name__ == "__main__":
    v3deck.build(CFG)
