# ETEH 4-Week Excel Data Analytics Bootcamp: Delivery Pack

Eaglet2EagleHub Ltd · Instructor: Abdulrahaman Kehinde

## Edition 3 (in progress): leaner decks, practical work in the workbooks

Edition 2 explained every topic with two fixed slides, which made the decks long and repetitive. Edition 3 is being rolled out one week at a time. **Week 1 is done**; Weeks 2–4 and the Capstone volume are still edition 2 until Week 1 is approved.

### Week 1 deck (`ETEH_Vol1_Week1_Foundations_Slides.pptx`, 146 slides, 111 topics)
- Built from the ETEH deck template (title, welcome, module divider, content frame, closing).
- **One slide per topic, laid out for what the topic needs.** There is no fixed template. Where it helps, a topic slide has, in this order:
  - the **standard definition**;
  - **in simple terms**;
  - **an analogy**;
  - and, only where it is useful, a short **example from our data**.
- Example tables are illustrations only. Each one is labelled with its **table name** (for example `tStores25`, `tDept23`, `tHotelSeg`), not a file name.
- Each module opens with a divider and "What you will learn", and closes with "Key takeaways". The week ends with a summary and a two-slide glossary.
- The Excel companion is introduced **once, at the start of the week** (slides 2 and 4, "How This Week Works"). Modules do not repeat it. After each module the tutor opens that module's sheet in the examples workbook and goes through it.

### Week 1 examples workbook (`01_Section_Examples/ETEH_Week1_Section_Examples.xlsx`)
Rebuilt from scratch as a self-contained practice book.
- **Sheets:**
  - one sheet per module (1.1–1.9);
  - four data sheets: `Data_Stores25` (tStores25), `Data_Sales` (tSales), `Data_DirtySales` (tDirtySales) and `Data_Students` (tStudents);
  - a cover with links to everything.
- **Every one of the 111 mini-sections is covered.** A few closely related ones are grouped, for example 1.6.12–13.
- **Each mini-section has:**
  - a short definition;
  - numbered steps;
  - a **worked solution**: live formulas, with the formula text shown beside each result;
  - **"Your turn" exercises**: about 220 across the nine sheets, with a white answer cell beside each.
- **Checks:** a live check marks each answer ✓ or ✗ as soon as it is typed. Hands-on tasks have their own checks; for example:
  - lock the rate with `$` (checked with FORMULATEXT);
  - convert a range to a table named `tIkejaJan`;
  - replace a typed number with a formula (checked with ISFORMULA);
  - unpivot quarters into a long table.
- **Answer key:** each module sheet ends with one, holding the worked answer to every exercise as a live formula on the module's data.
- **Datasets:** each module starts with a "This module's dataset" panel listing the tables and practice ranges it uses.
- **Verification:**
  - 825 formulas checked with `_build/verify_xlsx.py`: 550 recalculate to their cached values;
  - 7 known pycel limits were checked by hand: TRIM, hidden rows in SUBTOTAL, TEXT format codes, and INDIRECT on a table the learner has not created yet;
  - the remaining formulas use functions pycel cannot evaluate (mostly FORMULATEXT); their cached values were checked separately;
  - `_build/check_xml.py` passes.

### Week 1 assignment
- The tasks and marks are unchanged.
- A new **Function Guide** sheet shows the functions the assignment uses that go beyond the Week 1 workbook: XLOOKUP, AVERAGEIFS, MINIFS, COUNTIFS, EDATE, TEXT "mmm", a SUMIFS date window and the distinct-count pattern.
  - Each one is a live worked example on the assignment's own tables.
  - Each answers a different question from the tasks, so no answers are given away.
- Task briefs and the cover point to the guide.
- The student file and the answer key both verify with 0 mismatches.

### Corrections made while rebuilding Week 1
- tDirtySales holds **335** dates stored as text, not 113. The 335 are:
  - 116 in yyyy.mm.dd form;
  - 106 in dd/mm/yyyy;
  - 113 in dd-Mon-yyyy.
  The deck now says 335.
- The dirty-file duplicate count is stated as **60 repeated records plus 15 empty rows** (3,387 → 3,312 rows).
  - Excel's Remove Duplicates (all columns) reports 74 removed, because 14 of the 15 empty rows are duplicates of the first. Deleting the last empty row gives 3,312.
  - The workbook explains this in 1.8.9.

### Build
```
python3 _build/07_v3_week1_deck.py --preview   # deck (+ PNG previews in _build/_stage/v3_preview)
python3 _build/rev/lint_v3.py                  # deck lint
python3 _build/02_sec_w1.py                    # examples workbook (asserts all 111 topics are covered)
python3 _build/03_asg_w1.py                    # assignment + answer key
```

## Edition 2: beginner-friendly self-study slides

This is **edition 2** of the pack. The five decks were rewritten so a beginner can learn on their own, without an instructor in the room. The workbooks, datasets, assignments and projects are the same as in edition 1, apart from the corrections listed below. Edition 1 is kept unchanged in `ETEH_Bootcamp_Pack_v1.zip` at the repository root. This edition is in `ETEH_Bootcamp_Pack_v2.zip`.

### How each topic is taught
Every topic in all five volumes (663 topics) now has two slides.

1. **Concept slide**
   - **In simple terms**: a plain-language definition with the key words in bold.
   - **Think of it like this**: an everyday Nigerian analogy, such as a market stall, an owambe, WAEC rules or a danfo route.
   - **How to do it**: three numbered steps.
   - **Remember**: one tip.
   - **Avoid**: the most common beginner mistake.
2. **Worked-example slide**
   - **The situation**: a real business question on the pack datasets (SuperMart, Crestview, Lifecare, EkoSuites, the Talent Survey or SuperMart Online).
   - **Step by step**: three steps.
   - **What you should see**: a table of the exact results you should get.
   - **What this tells us**: the business meaning in one line.
   - **Try it yourself**: the workbook and sheet where you can repeat the example.

### Accuracy
- Every number on a worked-example slide was recalculated from the pack datasets, section workbooks or solution workbooks.
  - Hypothetical figures are labelled "example".
  - Training outcomes are labelled "training scenario".
- Edition 1 contained figures with no source in the data. These were removed and replaced with the real values. Examples include:
  - a "Tobacco" category;
  - "41 million rows";
  - a weighted CGPA;
  - "Ward 3 31%";
  - triage bands;
  - a Year-1 Engineering decline;
  - "45 minutes to 6 seconds".
- Where a slide and its workbook disagreed, the workbook was checked against the data and fixed. Each workbook was then rebuilt and re-verified with 0 formula mismatches.
  - Week 4: several insight notes, the Solver note, the zone-growth note, the misleading-axis chart and the AI provenance log.
  - Week 3: a misleading 'Saving' calculation in module 3.1.
  - Capstone (P.6 and P.9): the README limitation and the STAR answer wording.
- The discount-cap recommendation applies to **Furniture only** in every deck, workbook and memo:
  - Furniture's FY2025 margin fell to 1.4%, and 167 of its lines were discounted above 20%. Those lines lost ₦17.5m.
  - A 20% Furniture cap adds ₦17.7m–₦18.0m of profit a year, assuming 0–50% volume loss on the re-priced lines.
  - The chain-wide figures (455 lines lost ₦47.6m; ₦59.6m–₦71.5m for an all-category cap) are shown only as diagnosis or as a possible phase 2.
  - Models: Week 1 sheet 1.2 (1.2.4) and the Capstone 1 SOLUTION sheet '5 What-If'. The category is an input cell.
- Coherence checks run on all 1,406 slides (`_build/rev/lint_decks.py`, `lint_topics.py`, `qa_trace.py`):
  - every concept slide is followed by its worked example;
  - Next links and module footers are correct;
  - no empty table cells;
  - every Try-it workbook and sheet exists;
  - each figure is traced to a workbook, or was recomputed from the raw data.
- The decks use only characters that the slide font can display. Signs such as ≥ and ≤ are written as `>=` / `<=` or in words.

### Slide decks (pack root)

| Deck | Slides | Modules / topics | Companion files |
|---|---|---|---|
| ETEH_Vol1_Week1_Foundations_Slides.pptx | 146 (edition 3) | 1.1–1.9 · 111 topics | ETEH_Week1_Section_Examples.xlsx, ETEH_Week1_Assignment.xlsx |
| ETEH_Vol2_Week2_PowerQuery_Pivots_Slides.pptx | 367 | 2.1–2.11 · 175 topics | ETEH_Week2_Section_Examples.xlsx, Week2_Data_Files, ETEH_Week2_Assignment.xlsx |
| ETEH_Vol3_Week3_SQL_DAX_Dashboards_Slides.pptx | 332 | 3.1–3.10 · 158 topics | ETEH_Week3_Section_Examples.xlsx, Week3_Data_Files (SQLite), ETEH_Week3_Assignment.xlsx |
| ETEH_Vol4_Week4_Stats_Forecast_AI_Slides.pptx | 297 | 4.1–4.14 · 139 topics | ETEH_Week4_Section_Examples.xlsx, ETEH_Week4_Assignment.xlsx |
| ETEH_Vol5_Capstone_Career_Slides.pptx | 174 | P.1–P.9 · 80 topics | ETEH_Capstone_Section_Examples.xlsx, Capstones 1–3 (solution workbooks), Student Projects 1–4 |

**How to study a topic:** read the concept slide first. Then open the workbook sheet named under "Try it yourself" and follow the three steps. Check that your answers match the "What you should see" table. If your numbers differ, look at the **Avoid** box first.

### Edition 1 changes from the original volumes (still apply)
- The final two lab and practice modules of every volume were removed. Roadmaps, topic counts and the "In this session" panels were updated to match.
- Duplicate module codes were renumbered so slide codes match the worksheet names in the section workbooks:
  - Week 2 now runs 2.1–2.7, 2.8A, 2.8B, 2.9–2.11.
  - Week 3 keeps 3.6A and 3.6B.
- External sources were removed: Wikipedia tables, CBN/NBS figures and "live source" links. The modules that depended on them now use files in the pack:
  - 2.4 · Web and PDF sources: the SuperMart store directory HTML page and the monthly sales bulletin PDF.
  - 3.10 · File extracts into SQL, DAX and dashboards: the Lifecare 2023 dirty CSV and lifecare_hospital.db.
  - 4.14 · Internal benchmarks, PDF checks and reconciliation: the SuperMart ledger and bulletin, and Lifecare readmission intervals.
  - P.2 · Source governance: a register of the capstone sources.
- The Capstone 1–3 modules (P.3–P.5) are rewritten around the instructor capstones. Every figure on those slides comes from the solution workbooks.
- Headline worked examples elsewhere now quote real dataset figures, for example:
  - SuperMart FY2025: ₦733.2m sales, 12.7% margin, 3,312 lines.
  - Crestview: 6,607 students, 86.4% progression, 901 at risk.
  - Lifecare FY2023: 10,918 admissions, 89.5% occupancy.
- Week 1 dollar amounts were converted to naira. `$` signs inside cell references such as `$B$2` were kept.

## Folders
1. **01_Section_Examples**: one workbook per week, with one worksheet per module or subsection. It also contains:
   - 00_Core_Datasets: the standardised CSVs and the data dictionary.
   - Week2_Data_Files: dirty and multi-format files (CSV, pipe-delimited TXT, JSON, HTML, PDF, XLSX).
   - Week3_Data_Files: SQLite databases plus their schema and data SQL scripts.
2. **02_Assignments**: one multi-sheet assignment workbook per week, plus their data.
3. **03_Instructor_Projects**:
   - Assignment answer keys.
   - Student project keys.
   - Capstones 1–3, each with a solution workbook and an instructor brief PDF. Capstone 3 also has its extraction SQL.
4. **04_Student_Projects**: four projects for students.
   - P1 SuperMart Online SQL (joins).
   - P2 EkoSuites Hotel.
   - P3 Lifecare Diabetes Screening.
   - P4 Talent Salary Benchmark.

All files are generated by the scripts in `_build/` at the repository root.
- `05_update_slides.py` rebuilds the edition 1 decks from the original volumes.
- `06_revise_slides.py` builds the edition 2 decks from the topic copy in `_build/rev/` (one file per module; the layout rules are in `DESIGN.md`). For example, `python3 _build/06_revise_slides.py --vol=5` rebuilds Volume 5.
