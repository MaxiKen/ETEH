# -*- coding: utf-8 -*-
"""
ETEH Bootcamp - Presentation cleanup (Phase E).

For each of the 5 volumes:
  1. DELETE the last two "lab / non-learning" modules (delete from the start
     slide of the second-to-last module through the end of the deck).
  2. REPLACE the recurring external "example" dataset brand names with the real
     curated datasets shipped in this bootcamp, so inline examples stay intact
     but are now grounded in data the learner actually has:
        SuperMart Retail      -> Superstore Retail          (Sample - Superstore)
        Crestview University -> Student Performance Dataset  (Student Performance Factors)
        Lifecare Hospital     -> Lifecare Health Records     (healthcare_records)
  3. Modernise the welcome "Datasets:" line to list every curated dataset.

Outputs go to Presentations_Clean/ (originals are untouched).
"""
import os, re
from pptx import Presentation
from pptx.oxml.ns import qn

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC  = ROOT
OUT  = os.path.join(ROOT, "Presentations_Clean")
os.makedirs(OUT, exist_ok=True)

FILES = {
 1: ("ETEH_Vol1_Week1_Foundations_Slides-2.pptx", 178),
 2: ("ETEH_Vol2_Week2_PowerQuery_Pivots_Slides-1.pptx", 229),
 3: ("ETEH_Vol3_Week3_SQL_DAX_Dashboards_Slides-1.pptx", 223),
 4: ("ETEH_Vol4_Week4_Stats_Forecast_AI_Slides-1.pptx", 203),
 5: ("ETEH_Vol5_Capstone_Career_Slides-1.pptx", 130),
}

# ordered longest-first to avoid nested double substitution
REPL = [
 ("SuperMart Retail Analytics", "Superstore Retail Analytics"),
 ("SuperMart Retail Nigeria",  "Superstore Retail Nigeria"),
 ("SuperMart Retail",          "Superstore Retail"),
 ("SuperMart",                 "Superstore"),
 ("Crestview University Progression", "Student Performance Progression"),
 ("Crestview University",      "Student Performance Dataset"),
 ("Crestview Univ",           "Student Performance"),
 ("Crestview",                "Student Performance"),
 ("Lifecare General Hospital","Lifecare Health Records"),
 ("Lifecare Hospital Operations","Lifecare Health Operations"),
 ("Lifecare Hospital",        "Lifecare Health Records"),
 ("Lifecare",                 "Lifecare Health Records"),
]
DATASETS_LINE = ("Datasets: Superstore Sales, Healthcare Records, E-Commerce Marketplace (SQL), "
                 "Student Performance, Salary Survey, Hotel Reservations, Diabetes Prediction, "
                 "Social-Media Wellbeing, Student Depression.")

def apply_repl(text):
    for a, b in REPL:
        text = text.replace(a, b)
    return text

def process_text_frame(tf, fix_datasets):
    for para in tf.paragraphs:
        full = para.text.strip()
        if fix_datasets and full.lower().startswith("datasets"):
            # replace whole paragraph with the canonical curated list
            if para.runs:
                para.runs[0].text = DATASETS_LINE
                for r in para.runs[1:]:
                    r.text = ""
            continue
        for run in para.runs:
            if run.text:
                run.text = apply_repl(run.text)

def process_table(table):
    for row in table.rows:
        for cell in row.cells:
            if cell.text_frame:
                process_text_frame(cell.text_frame, False)

def delete_slides_from(prs, start_1based):
    xml = prs.slides._sldIdLst
    ids = list(xml)
    # keep indices 0 .. start_1based-1 (0-based)
    for i in range(len(ids) - 1, start_1based - 1 - 1, -1):
        xml.remove(ids[i])

def clean(path, del_from):
    prs = Presentation(path)
    before = len(prs.slides)
    delete_slides_from(prs, del_from)
    after = len(prs.slides)
    # text replacement on remaining slides
    n_changed = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                process_text_frame(shape.text_frame, True)
            if shape.has_table:
                process_table(shape.table)
        if slide.has_notes_slide:
            ns = slide.notes_slide
            if ns.notes_text_frame:
                process_text_frame(ns.notes_text_frame, False)
    out = os.path.join(OUT, os.path.basename(path))
    prs.save(out)
    return before, after

if __name__ == "__main__":
    for vol, (fname, del_from) in FILES.items():
        b, a = clean(os.path.join(SRC, fname), del_from)
        print(f"Vol{vol}: {fname}  slides {b} -> {a}  (removed {b-a}, last-2 modules deleted)")
