#!/usr/bin/env python3
"""
ETEH Bootcamp - Presentation cleanup (edits the 5 volumes)
 * Removes the non-learning "lab" modules: 1.10, 2.12, 3.11, 4.15.
 * Strips each removed module from its Programme Roadmap.
 * Retargets the three external dataset names to our standardised datasets,
   keeping inline worked-example tables (per the agreed 'retarget' decision).
Saves cleaned decks into ETEH_Bootcamp/presentations/. Originals are untouched.
"""
import os, re
from pptx import Presentation
from pptx.oxml.ns import qn

ROOT="/home/user/ETEH"
SRC=["ETEH_Vol1_Week1_Foundations_Slides-2.pptx",
     "ETEH_Vol2_Week2_PowerQuery_Pivots_Slides-1.pptx",
     "ETEH_Vol3_Week3_SQL_DAX_Dashboards_Slides-1.pptx",
     "ETEH_Vol4_Week4_Stats_Forecast_AI_Slides-1.pptx",
     "ETEH_Vol5_Capstone_Career_Slides-1.pptx"]
OUT=os.path.join(ROOT,"ETEH_Bootcamp","presentations")
os.makedirs(OUT,exist_ok=True)

# module code -> (title prefix used on roadmap, nothing else) for the labs we drop
REMOVE={
 "ETEH_Vol1_Week1_Foundations_Slides-2.pptx": {"1.10":"Practical Labs, Data Sins Audit and Week 1 Capston"},
 "ETEH_Vol2_Week2_PowerQuery_Pivots_Slides-1.pptx": {"2.12":"Week 2 Practical Exercises and Assignment"},
 "ETEH_Vol3_Week3_SQL_DAX_Dashboards_Slides-1.pptx": {"3.11":"Labs, Capstone and Desk Reference"},
 "ETEH_Vol4_Week4_Stats_Forecast_AI_Slides-1.pptx": {"4.15":"Labs, Capstone and Desk Reference"},
 "ETEH_Vol5_Capstone_Career_Slides-1.pptx": {},
}

# Retargeting rules (longest / most specific first)
RULES=[
 (r"SuperMart Retail", "ETEH Retail"),
 (r"Supermart Retail", "ETEH Retail"),
 (r"SuperMart", "ETEH Retail"),
 (r"Supermart", "ETEH Retail"),
 (r"Superstore", "ETEH Retail"),
 (r"Crestview University", "ETEH University"),
 (r"Crestview", "ETEH University"),
 # Lifecare rules: negative look-arounds stop the generic rule re-matching
 # inside its own output ("ETEH Lifecare Hospital").
 (r"(?<!ETEH )Lifecare General Hospital", "ETEH Lifecare Hospital"),
 (r"(?<!ETEH )Lifecare Hospital", "ETEH Lifecare Hospital"),
 (r"(?<!ETEH )Lifecare(?! Hospital)", "ETEH Lifecare Hospital"),
 (r"e-commerce", "ETEH Commerce"),
 (r"E-Commerce", "ETEH Commerce"),
 # retarget internal folder references created by the teaching modules to the real structure
 (r"assignment_data", "datasets/sql"),
 (r"assignments_and_projects", "assignments"),
]
COMPILED=[(re.compile(p, re.IGNORECASE), r) for p,r in RULES]

def repl(text):
    if not isinstance(text,str): return text
    for rx,rep in COMPILED:
        text=rx.sub(rep, text)
    return text

def shape_text(sh):
    if sh.has_text_frame: return sh.text_frame.text
    if sh.has_table:
        return "\n".join(" | ".join(c.text for c in r.cells) for r in sh.table.rows)
    return ""

def edit_shape(sh):
    changed=False
    if sh.has_text_frame:
        for para in sh.text_frame.paragraphs:
            for run in para.runs:
                n=repl(run.text)
                if n!=run.text: run.text=n; changed=True
    if sh.has_table:
        for row in sh.table.rows:
            for cell in row.cells:
                for para in cell.text_frame.paragraphs:
                    for run in para.runs:
                        n=repl(run.text)
                        if n!=run.text: run.text=n; changed=True
    return changed

def module_code(text):
    m=re.search(r"module\s+([0-9]+\.[0-9]+|[A-Z]\.[0-9]+)", text, re.I)
    return m.group(1) if m else None

def slide_text(slide):
    parts=[]
    for sh in slide.shapes:
        t=shape_text(sh)
        if t: parts.append(t)
    return "\n".join(parts)

def delete_slide(prs, idx):
    sldIdLst=prs.slides._sldIdLst
    sldIds=list(sldIdLst)
    sldIdLst.remove(sldIds[idx])

for fname in SRC:
    prs=Presentation(os.path.join(ROOT,fname))
    remove_map=REMOVE[fname]
    n0=len(prs.slides)
    # 1) mark lab slides for deletion
    drop=[]
    for i,slide in enumerate(prs.slides):
        code=module_code(slide_text(slide))
        if code in remove_map:
            drop.append(i)
    for i in sorted(drop, reverse=True):
        delete_slide(prs, i)
    # 2) clean roadmap entries for removed modules
    for slide in prs.slides:
        txt=slide_text(slide)
        if "Roadmap" in txt and any(c in txt for c in remove_map):
            shapes=list(slide.shapes)
            for k,sh in enumerate(shapes):
                t=sh.text_frame.text if sh.has_text_frame else ""
                for code,prefix in remove_map.items():
                    if t.startswith(prefix):
                        # remove title (k), code (k-1), count (k+1)
                        for j in (k,k-1,k+1):
                            if 0<=j<len(shapes):
                                sp=shapes[j]
                                # only remove if it looks like the code or a short count
                                st=sp.text_frame.text if sp.has_text_frame else ""
                                if j==k-1 and st.strip()!=code: continue
                                sp._element.getparent().remove(sp._element)
                        break
    # 3) retarget dataset names everywhere
    for slide in prs.slides:
        for sh in slide.shapes:
            edit_shape(sh)
    n1=len(prs.slides)
    out=os.path.join(OUT, fname.replace(".pptx","_CLEAN.pptx"))
    prs.save(out)
    print(f"{fname}: slides {n0} -> {n1} (removed {len(drop)} lab slides)")
print("DONE presentations")
