"""Edition 3: build the Capstone deck (Volume 5).

    python3 _build/07_v3_capstone_deck.py [--preview]

Engine: _build/v3deck.py. Copy lives in _build/v3/w5_*.py; every figure is computed in _build/v3/w5_fx.py from the
instructor capstone builders (04_proj_i1/i2/i3) and the pack datasets. Charts are attached in the copy files.
"""
from __future__ import annotations

import importlib
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE / "v3"))
import v3deck  # noqa: E402
import w5_data  # noqa: E402
from w5_extra import GLOSSARY, SUMMARY  # noqa: E402

PARTS = ["w5_m1_2", "w5_m3", "w5_m4", "w5_m5", "w5_m6_7", "w5_m8_9"]
MODULES = []
for p in PARTS:
    try:
        MODULES += importlib.import_module(p).MODULES
    except ModuleNotFoundError as e:
        if e.name != p:
            raise
        print(f"(skipping {p}: not written yet)")

SHEETS = {"P.1": "P.1 Project Lifecycle", "P.2": "P.2 Source Register", "P.3": "P.3 Capstone1 SuperMart",
          "P.4": "P.4 Capstone2 Kingsway", "P.5": "P.5 Capstone3 Mercyland", "P.6": "P.6 Workbook QA",
          "P.7": "P.7 Executive Defence", "P.8": "P.8 Portfolio & CV", "P.9": "P.9 Interviews & Roadmap"}

CFG = v3deck.Config(
    week=5,
    unit="Capstone",
    week_title="Capstone: Three End-to-End Builds, Executive Defence & Career Launch",
    title_sub="Run a full analytics project from brief to handover, prove every number, defend it to a panel and turn it "
              "into a portfolio",
    out=v3deck.ROOT / "ETEH_Bootcamp_Pack" / "ETEH_Vol5_Capstone_Career_Slides.pptx",
    workbook="ETEH_Capstone_Section_Examples.xlsx",
    modules=MODULES,
    sheets=SHEETS,
    data=w5_data,
    summary=SUMMARY,
    glossary=GLOSSARY,
    roadmap_label="NINE MODULES IN THE CAPSTONE",
    roadmap_notes="Walk through the modules. P.1-P.2 are the method; P.3-P.5 are the three capstones (retail, education, "
                  "health); P.6 is QA; P.7 is the defence; P.8-P.9 are portfolio, CV, interviews and what to learn next.",
    title_notes="Welcome to the capstone. Weeks 1-4 taught the tools. Now we run three complete projects on the bootcamp data, "
                "prove every number, defend the findings to a panel and turn the work into a portfolio.",
    assignment="",
    assessment="Assessment: the capstone builds, the four projects (04_Student_Projects › Week_5_Capstone), "
               "a ten-minute defence and a write-up.",
    volume="Volume 5 of 5",
    modules_text="9 modules",
    example_table="tSM",
)

if __name__ == "__main__":
    v3deck.build(CFG)
