#!/usr/bin/env python3
"""Clean the 5 ETEH volumes:
- Delete last 2 modules per volume (labs/capstone + extension) + external-data modules
- Renumber sequentially, fix duplicates
- Replace fictional datasets (SuperMart/Crestview/Lifecare) with real attached datasets
- Remove Wikipedia/PDF-bulletin/external-benchmark mentions (keep legitimate PDF-export content)
- Update Welcome + Roadmap slides
Outputs to ETEH_Bootcamp/Presentations/
"""
import re
from pptx import Presentation
from pathlib import Path

ROOT = Path("/home/user/ETEH")
OUT = ROOT/"ETEH_Bootcamp/Presentations"
OUT.mkdir(parents=True, exist_ok=True)

# Deletions: (file, [(start,end) inclusive], new_filename, module_map, new_counts_text, new_dataset_text)
# module_map: old module label -> new module label, applied to "Module X.Y" and "MODULE X.Y" and topic codes
# For topic codes we map prefix: e.g. "2.8." -> "2.9." etc. Handled via module_map on first two components.

CONFIGS = [
    {
        "src": "ETEH_Vol1_Week1_Foundations_Slides-2.pptx",
        "dst": "ETEH_Vol1_Week1_Foundations_CLEANED.pptx",
        "delete": [(178,202)],
        "module_map": {},  # 1.1-1.9 unchanged
        "welcome_old": "121 topics across 11 modules.",
        "welcome_new": "111 topics across 9 modules.",
        "roadmap_remove": ["1.10","1.11"],
    },
    {
        "src": "ETEH_Vol2_Week2_PowerQuery_Pivots_Slides-1.pptx",
        "dst": "ETEH_Vol2_Week2_PowerQuery_Pivots_CLEANED.pptx",
        "delete": [(37,47),(229,252)],
        # old -> new (sequential 2.1-2.11). Note duplicates: two old 2.4s, two old 2.7s.
        # We distinguish by slide range below via divider order, not just label.
        "module_map": {"2.1":"2.1","2.2":"2.2","2.3":"2.3","2.5":"2.5","2.6":"2.6"},
        "welcome_old": "184 topics across 14 modules.",
        "welcome_new": "170 topics across 11 modules.",
        "roadmap_remove": ["External Web and PDF","Week 2 Practical","Repeatable Power Query"],
    },
    {
        "src": "ETEH_Vol3_Week3_SQL_DAX_Dashboards_Slides-1.pptx",
        "dst": "ETEH_Vol3_Week3_SQL_DAX_Dashboards_CLEANED.pptx",
        "delete": [(212,222),(223,242)],
        "module_map": {"3.1":"3.1","3.2":"3.2","3.3":"3.3","3.4":"3.4","3.5":"3.5"},
        "welcome_old": "167 topics across 13 modules.",
        "welcome_new": "153 topics across 10 modules.",
        "roadmap_remove": ["External Data in SQL","Labs, Capstone","Database Design"],
    },
    {
        "src": "ETEH_Vol4_Week4_Stats_Forecast_AI_Slides-1.pptx",
        "dst": "ETEH_Vol4_Week4_Stats_Forecast_AI_CLEANED.pptx",
        "delete": [(192,202),(203,225)],
        "module_map": {},  # 4.1-4.13 unchanged
        "welcome_old": "148 topics across 16 modules.",
        "welcome_new": "134 topics across 13 modules.",
        "roadmap_remove": ["External Benchmarks","Labs, Capstone","Statistical Reasoning"],
    },
    {
        "src": "ETEH_Vol5_Capstone_Career_Slides-1.pptx",
        "dst": "ETEH_Vol5_Capstone_Career_CLEANED.pptx",
        "delete": [(130,152)],
        "module_map": {},  # P.1-P.10 unchanged
        "welcome_old": "92 topics across 12 modules.",
        "welcome_new": "83 topics across 10 modules.",
        "roadmap_remove": ["Diagnostic Examination","Governed Delivery"],
    },
]

# Dataset name replacements (applied to all remaining text, case-sensitive variants)
DATASET_REPLACEMENTS = [
    ("SuperMart Retail Nigeria (12 stores, 144 SKUs and 35,712 POS lines of trading history), Lifecare General Hospital in Ikeja and Crestview University in Abuja",
     "E-commerce Nigeria (2,500 customers, 800 products, 11,000 transaction lines), Superstore US Retail (9,994 lines), Student Performance (6,607 students) and Healthcare Admissions (55,500 admissions)"),
    ("SuperMart Retail", "E-commerce & Superstore Retail"),
    ("Supermart Retail", "E-commerce & Superstore Retail"),
    ("SuperMart", "Superstore"),
    ("Supermart", "Superstore"),
    ("Crestview University", "Student Performance"),
    ("Crestview", "Student Performance"),
    ("Lifecare General Hospital", "Healthcare"),
    ("Lifecare Hospital", "Healthcare"),
    ("Lifecare", "Healthcare"),
    # Capstone titles (Vol5) — more specific, applied after generic? Order matters: do specific first.
    # (Generic above already changed them; fix capstone titles explicitly below.)
    # External mentions in remaining slides
    ("Wikipedia population source: https://en.wikipedia.org/wiki/List_of_Nigerian_states_by_population",
     "E-commerce Nigeria source: e-commerce.xlsx (Transactions, 11,000 lines, 2020–2025) — see Datasets/CLEAN/"),
    ("Cite Wikipedia and preserve CC BY-SA attribution; review the licence for redistributed adaptations.",
     "Cite the attached dataset file and preserve its source note; review the licence for redistributed adaptations."),
    ("CBN FX rate / NBS CPI inflation CSV", "E-commerce monthly revenue CSV (attached)"),
    ("SuperMart_POS.csv , Clinic_Visits.txt", "ecommerce_transactions_CLEAN.csv , healthcare_ANONYMISED.csv"),
    ("Admissions_Portal.csv , LMS_Logins.xlsx , Bursary_Fee_Status.csv", "student_performance_CLEAN.csv , social_media_addiction_CLEAN.csv"),
    ("Monthly_Branch_CSVs folder; dropping 2026-04_April.csv", "Datasets/CLEAN folder; dropping a new ecommerce_transactions export"),
    ("legacy exports and PDF extractions", "legacy exports"),
    ("Datasets: SuperMart Retail, Crestview University, Lifecare General Hospital.",
     "Datasets (attached, standardised): E-commerce Nigeria, Superstore Retail, Student Performance, Healthcare, Salary, Hotel, Diabetes, Social Media, Student Wellbeing."),
]

def replace_text_in_shape(shape, replacements):
    if shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                for old, new in replacements:
                    if old in run.text:
                        run.text = run.text.replace(old, new)
    if shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                for para in cell.text_frame.paragraphs:
                    for run in para.runs:
                        for old, new in replacements:
                            if old in run.text:
                                run.text = run.text.replace(old, new)

def delete_slides(prs, delete_ranges):
    to_delete = set()
    for a,b in delete_ranges:
        to_delete.update(range(a,b+1))
    # delete in reverse order (1-based)
    sldIdLst = prs.slides._sldIdLst
    # map slide index -> element
    slides = list(prs.slides._sldIdLst)
    for idx in sorted(to_delete, reverse=True):
        # idx is 1-based original position
        el = slides[idx-1]
        # remove relationship
        rId = el.rId
        prs.part.drop_rel(rId)
        sldIdLst.remove(el)
    return to_delete

def renumber_vol2(prs):
    """Vol2: map by divider order. Old dividers in order (after deletions):
    2.1->2.1, 2.2->2.2, 2.3->2.3, 2.4(Cleaning)->2.4, 2.5->2.5, 2.6->2.6,
    2.7a->2.7, 2.7b->2.8, 2.8->2.9, 2.9->2.10, 2.10->2.11
    We identify each slide's old module by its 'Module X.Y' label + slide position.
    Simpler: iterate slides in order, track current old module by divider, assign new sequentially.
    """
    new_seq = ["2.1","2.2","2.3","2.4","2.5","2.6","2.7","2.8","2.9","2.10","2.11"]
    # First pass: find divider slides in current (post-delete) deck and record old->new per divider occurrence
    divider_new = {}  # slide index (1-based current) -> new module
    div_order = 0
    for i, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text.strip().startswith("MODULE"):
                if div_order < len(new_seq):
                    divider_new[i] = new_seq[div_order]
                div_order += 1
                break
    print(f"  Vol2 dividers found: {div_order}, mapping {divider_new}")
    # Second pass: for each slide, determine its new module (last divider <= slide), and old module (from label)
    # Build old->new mapping per occurrence: need old label per divider
    old_labels = []
    for i, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text.strip().startswith("MODULE"):
                old_labels.append(shape.text.strip())  # e.g. MODULE 2.4 · 12 TOPICS
                break
    # old_labels order corresponds to new_seq order
    # Now replace in each slide: MODULE old -> MODULE new, Module old -> Module new, topic codes old.x -> new.x
    # We must be careful: old 2.4 appears once (Cleaning) -> new 2.4 (no change). Old 2.7 appears twice -> new 2.7 and 2.8.
    # Approach: process slides sequentially, knowing current new module; replace old prefix with new prefix for that slide only.
    current_new = None
    current_old_prefix = None
    # Precompute per-slide (old_prefix, new_prefix) by scanning dividers
    slide_map = {}
    cur_old = None; cur_new = None; div_idx = 0
    for i, slide in enumerate(prs.slides, 1):
        is_div = i in divider_new
        if is_div:
            # parse old module from divider text
            for shape in slide.shapes:
                if shape.has_text_frame and shape.text.strip().startswith("MODULE"):
                    m = re.search(r"MODULE\s+(\d+\.\d+)", shape.text.strip())
                    cur_old = m.group(1) if m else None
                    break
            cur_new = divider_new[i]
        slide_map[i] = (cur_old, cur_new)
    # Apply replacements slide by slide
    for i, slide in enumerate(prs.slides, 1):
        old, new = slide_map[i]
        if not old or not new or old==new:
            continue
        reps = [
            (f"MODULE {old} ", f"MODULE {new} "),
            (f"Module {old} ", f"Module {new} "),
            (f"Module {old}·", f"Module {new}·"),
            (f"{old}.", f"{new}."),  # topic codes like 2.8.1 -> 2.9.1 (only when old!=new; careful with 2.10 vs 2.1)
        ]
        # For topic codes, use regex to avoid 2.1 matching 2.10: replace f"{old}." only when followed by digit and not part of longer number?
        # Our reps with f"{old}." for old=2.8 -> "2.8." won't match "2.10." — safe. For old=2.9 -> "2.9." safe.
        # But old=2.1 vs new=2.1 (no change, skipped). Good.
        for shape in slide.shapes:
            # custom: replace topic codes via regex
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        # MODULE / Module labels
                        run.text = run.text.replace(f"MODULE {old} ", f"MODULE {new} ")
                        run.text = run.text.replace(f"Module {old} ", f"Module {new} ")
                        # topic code standalone like "2.8.1" or "2.8.1  ·"
                        run.text = re.sub(rf"\b{re.escape(old)}\.(\d+)\b", rf"{new}.\1", run.text)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for para in cell.text_frame.paragraphs:
                            for run in para.runs:
                                run.text = run.text.replace(f"MODULE {old} ", f"MODULE {new} ")
                                run.text = run.text.replace(f"Module {old} ", f"Module {new} ")
                                run.text = re.sub(rf"\b{re.escape(old)}\.(\d+)\b", rf"{new}.\1", run.text)
    print("  Vol2 renumber applied")

def renumber_vol3(prs):
    """Vol3: old 3.1->3.1, 3.2->3.2, 3.3->3.3, 3.4->3.4, 3.5->3.5, 3.6a->3.6, 3.6b->3.7, 3.7->3.8, 3.8->3.9, 3.9->3.10"""
    new_seq = ["3.1","3.2","3.3","3.4","3.5","3.6","3.7","3.8","3.9","3.10"]
    divider_new = {}
    div_order = 0
    for i, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text.strip().startswith("MODULE"):
                if div_order < len(new_seq):
                    divider_new[i] = new_seq[div_order]
                div_order += 1
                break
    print(f"  Vol3 dividers found: {div_order}, mapping {divider_new}")
    slide_map = {}
    cur_old=None; cur_new=None
    for i, slide in enumerate(prs.slides, 1):
        if i in divider_new:
            for shape in slide.shapes:
                if shape.has_text_frame and shape.text.strip().startswith("MODULE"):
                    m = re.search(r"MODULE\s+(\d+\.\d+)", shape.text.strip())
                    cur_old = m.group(1) if m else None
                    break
            cur_new = divider_new[i]
        slide_map[i]=(cur_old,cur_new)
    for i, slide in enumerate(prs.slides, 1):
        old,new = slide_map[i]
        if not old or not new or old==new: continue
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        run.text = run.text.replace(f"MODULE {old} ", f"MODULE {new} ")
                        run.text = run.text.replace(f"Module {old} ", f"Module {new} ")
                        run.text = re.sub(rf"\b{re.escape(old)}\.(\d+)\b", rf"{new}.\1", run.text)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for para in cell.text_frame.paragraphs:
                            for run in para.runs:
                                run.text = run.text.replace(f"MODULE {old} ", f"MODULE {new} ")
                                run.text = run.text.replace(f"Module {old} ", f"Module {new} ")
                                run.text = re.sub(rf"\b{re.escape(old)}\.(\d+)\b", rf"{new}.\1", run.text)
    print("  Vol3 renumber applied")

def update_welcome_roadmap(prs, cfg):
    # Slide 2 (index 1): update counts + datasets
    s2 = prs.slides[1]
    for shape in s2.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if cfg["welcome_old"] in run.text:
                        run.text = run.text.replace(cfg["welcome_old"], cfg["welcome_new"])
                    if "Datasets: SuperMart Retail" in run.text:
                        run.text = run.text.replace("Datasets: SuperMart Retail, Crestview University, Lifecare General Hospital.",
                                                    "Datasets (attached, standardised): E-commerce Nigeria, Superstore Retail, Student Performance, Healthcare, Salary, Hotel, Diabetes, Social Media, Student Wellbeing.")
    # Slide 3 roadmap: remove entries for deleted modules.
    # Each entry = 4 shapes at same top: [empty, number, title, count]. Clear all 4 for deleted titles.
    s3 = prs.slides[2]
    # Group shapes by top position (rounded)
    from collections import defaultdict
    groups = defaultdict(list)
    for shape in s3.shapes:
        if shape.has_text_frame:
            col = 0 if shape.left < 5700000 else 1
            groups[(round(shape.top), col)].append(shape)
    for top, shapes in groups.items():
        # check if any shape in group matches a remove string
        texts = [s.text.strip() for s in shapes]
        if any(any(rem in txt for rem in cfg["roadmap_remove"]) for txt in texts):
            for s in shapes:
                s.text_frame.clear()
    # Also need to renumber roadmap module numbers for Vol2/Vol3 (sequential). Roadmap numbers are separate shapes like "2.4", "2.7" etc.
    # For Vol2/Vol3, rewrite roadmap numbers in order of appearance.
    if cfg["src"].startswith("ETEH_Vol2"):
        # Remaining roadmap entries in order should be 2.1..2.11. Find number shapes (text matches ^2\.\d+$) in top-to-bottom order.
        # python-pptx shapes order is creation order, which matches visual order for roadmap (top to bottom). Use that.
        nums = []
        for shape in s3.shapes:
            if shape.has_text_frame and re.match(r"^2\.\d+$", shape.text.strip()):  # deleted entries already cleared
                nums.append(shape)
        new_nums = ["2.1","2.2","2.3","2.4","2.5","2.6","2.7","2.8","2.9","2.10","2.11"]
        for shape, nn in zip(nums, new_nums):
            shape.text_frame.clear()
            p = shape.text_frame.paragraphs[0]
            run = p.add_run(); run.text = nn
        print(f"  Vol2 roadmap renumbered: {len(nums)} entries")
    if cfg["src"].startswith("ETEH_Vol3"):
        nums = []
        for shape in s3.shapes:
            if shape.has_text_frame and re.match(r"^3\.\d+$", shape.text.strip()):
                nums.append(shape)
        new_nums = ["3.1","3.2","3.3","3.4","3.5","3.6","3.7","3.8","3.9","3.10"]
        for shape, nn in zip(nums, new_nums):
            shape.text_frame.clear()
            p = shape.text_frame.paragraphs[0]
            run = p.add_run(); run.text = nn
        print(f"  Vol3 roadmap renumbered: {len(nums)} entries")

for cfg in CONFIGS:
    print(f"\n=== {cfg['src']} ===")
    prs = Presentation(ROOT/cfg["src"])
    print(f" original slides: {len(prs.slides)}")
    deleted = delete_slides(prs, cfg["delete"])
    print(f" deleted {len(deleted)} slides, remaining: {len(prs.slides)}")
    # Renumber before dataset replacement (order doesn't matter much)
    if cfg["src"].startswith("ETEH_Vol2"):
        renumber_vol2(prs)
    elif cfg["src"].startswith("ETEH_Vol3"):
        renumber_vol3(prs)
    # Dataset replacements on all remaining slides
    for slide in prs.slides:
        for shape in slide.shapes:
            replace_text_in_shape(shape, DATASET_REPLACEMENTS)
    # Fix Vol5 capstone titles that generic replacement may have mangled: ensure correct titles
    if cfg["src"].startswith("ETEH_Vol5"):
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            # Undo generic: "Capstone 1: Superstore Retail Analytics" -> correct
                            run.text = run.text.replace("Capstone 1: Superstore Retail Analytics", "Capstone 1: E-commerce & Retail Analytics")
                            run.text = run.text.replace("Capstone 2: Student Performance Progression", "Capstone 2: Student Performance Analytics")
                            run.text = run.text.replace("Capstone 3: Healthcare Operations Analytics Analytics", "Capstone 3: Healthcare Operations Analytics")
                            run.text = run.text.replace("Capstone 3: Healthcare Operations", "Capstone 3: Healthcare Operations Analytics")
                            # Divider titles: "Capstone 1: Superstore Retail Analytics" etc. already handled
                            # Roadmap titles truncated: "Capstone 1: Superstore Retail A" -> fix
                            run.text = run.text.replace("Capstone 1: Superstore Retail A", "Capstone 1: E-commerce & Retail A")
                            run.text = run.text.replace("Capstone 2: Student Performance Pro", "Capstone 2: Student Performance A")
                            run.text = run.text.replace("Capstone 3: Healthcare Hospital O", "Capstone 3: Healthcare Operations A")
    update_welcome_roadmap(prs, cfg)
    # Update Thank You slide to note cleaned edition + dataset pack location
    thank = prs.slides[-1]
    for shape in thank.shapes:
        if shape.has_text_frame and "Questions, discussion" in shape.text:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if "Questions, discussion" in run.text:
                        run.text = run.text.replace("Questions, discussion and practice tasks next.",
                                                    "Cleaned edition: labs/capstone & external-data modules removed. All examples use the attached standardised datasets (see Datasets/). Questions and practice next.")
    out_path = OUT/cfg["dst"]
    prs.save(str(out_path))
    print(f" saved {out_path.name} ({out_path.stat().st_size/1024:.0f} KB), slides={len(prs.slides)}")

print("\nPresentation cleaning complete.")
