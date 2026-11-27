# ETEH 4-Week Excel Data Analytics Bootcamp: Delivery Pack

Eaglet2EagleHub Ltd · Instructor: Abdulrahaman Kehinde

## Edition 3 (in progress): leaner decks, practical work in the workbooks

Edition 2 explained every topic with two fixed slides, which made the decks long and repetitive. Edition 3 is being rolled out one week at a time. **Weeks 1, 2 and 3 are done**; Week 4 and the Capstone volume are still edition 2.

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
  - 825 formulas checked with `_build/verify_xlsx.py`: 553 recalculate to their cached values;
  - 4 known pycel limits were checked by hand: a deliberately hidden row in SUBTOTAL (row 73 of 1.6), TEXT format codes (twice), and INDIRECT on a table the learner has not created yet. (The TRIM cases flagged earlier now verify: the checker uses Excel's TRIM, see Week 2.)
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

### Week 2 deck (`ETEH_Vol2_Week2_PowerQuery_Pivots_Slides.pptx`, 219 slides, 175 topics)
- The same one-slide-per-topic design as Week 1, for modules 2.1–2.7, 2.8A, 2.8B and 2.9–2.11. Titles match edition 2.
- Every figure is computed from the data when the deck is built (`_build/w2_data.py`, `_build/v3/w2_fx.py`); nothing is typed in by hand.
- Cell addresses quoted on the slides are real cells of `Data_Dirty25` in the workbook. For example R23 holds the text "₦239,240", G31 holds a name with a non-breaking space and a line feed, and C24, C63 and C72 hold the three text-date patterns.

### Week 2 examples workbook (`01_Section_Examples/ETEH_Week2_Section_Examples.xlsx`)
Rebuilt from scratch in the Week 1 format.
- **Sheets:**
  - one sheet per module: 2.1–2.7, 2.8A, 2.8B, 2.9–2.11;
  - `2.10 Live Pivots`: five real PivotTables on tSM, which Excel builds when the file opens;
  - five data sheets, each with one table starting in A1: `Data_Sales25` (tSales25), `Data_Stores` (tStores), `Data_Dirty25` (tDirty25, the raw file as Excel opens it), `Data_SM` (tSM, 9,994 lines 2022–2025) and `Data_Adm23` (tAdm23).
- **All 175 mini-sections are covered** (the build asserts it). Each has Steps, a Worked solution with live formulas, and **281 "Your turn" exercises** in all, with live ✓/✗ checks and an answer key on every sheet.
- **Power Query topics** (2.2–2.4, 2.6, 2.9) give the M script to paste, and ask for the row count or total the query produces. The answer key works these out from the tables in the workbook wherever the same data is there.
- **2.8A/2.8B:** one topic per function (65 in all), using the slide's definition. Each function is worked on tSales25, tStores, tDirty25 or tAdm23 and has at least one exercise.
- **2.10/2.11:** every pivot figure has a formula twin (SUMIFS/COUNTIFS on tSM), so learners can check their pivot. The charts are drawn from the twin tables, so they show in any viewer.
- **Verification:**
  - 1,329 formulas checked with `_build/verify_xlsx.py`: 698 recalculate to their cached values;
  - 3 pycel limits were checked by hand: IFERROR over an array inside SUMPRODUCT (twice), and a TEXT format with `,,` scaling;
  - the remaining formulas use functions pycel cannot evaluate (mostly the FORMULATEXT column);
  - every formula passes a syntax check.
- **Checker fixes found along the way** (`verify_xlsx.py`, `pycel_plugin.py`):
  - pycel's TRIM did not strip leading and trailing spaces; it now matches Excel's;
  - NETWORKDAYS and WORKDAY ignored holidays; they now use them;
  - one failed cell used to make pycel fail every later cell; it is now reset after each error, so many more cells are actually checked.
- **Bug fixed in the shared exercise code (`w1x.py`):** text answers containing a double quote (for example `₦#,##0.0,,"m"` or `"|"`) produced invalid check formulas. Quotes are now escaped. The Week 1 workbook had no such answers and passes the same syntax check.

### Week 2 assignment
- The tasks and marks are unchanged.
- A new **Function Guide** sheet covers WEEKDAY, the only function the tasks use that the Week 2 workbook does not teach. It also shows three patterns the workbook teaches only in simpler forms: a conditional median with AGGREGATE(17,6,…), MAX(COUNTIF(range,{list})) and XLOOKUP(MAX(…)). Each is a live example on tRaw that answers a different question from the tasks.
- The Task 4 brief and the cover point to the guide.
- The student file and the answer key both verify with 0 mismatches.

### Week 3 deck (`ETEH_Vol3_Week3_SQL_DAX_Dashboards_Slides.pptx`, 200 slides, 158 topics)
- The same one-slide-per-topic design, for modules 3.1–3.5, 3.6A, 3.6B and 3.7–3.10. Titles match edition 2.
- Every figure is computed when the deck is built: SQL panels are run against the SQLite databases in `Week3_Data_Files`, and DAX results are worked out from the same tables (`_build/w3_data.py`, `_build/v3/w3_fx.py`).
- Edition 2 year counts that did not match the data were recomputed, and the filter used is stated on the slide.

### Week 3 examples workbook (`01_Section_Examples/ETEH_Week3_Section_Examples.xlsx`)
Rebuilt from scratch in the Week 1 format (`_build/02_sec_w3.py` with `_build/v3/wb3_a.py`, `wb3_b.py`, `wb3_c.py`).
- **Sheets:**
  - one sheet per module: 3.1–3.5, 3.6A, 3.6B, 3.7–3.10;
  - 14 data sheets, each with one table starting in A1: SuperMart Online (tTransactions, tLines, tProducts, tCustomers, tSellers, tDeliveries, tStates, tDimDate), SuperMart Stores FY2025 (tDash, tDashStores), Lifecare (tAdm23, tCensus, and tLcDirty, the raw 2023 extract as received) and Crestview (tCV).
- **All 158 mini-sections are covered** (the build asserts it). Each has Steps, a Worked solution with live formulas, and **243 "Your turn" exercises** in all, with live ✓/✗ checks and an answer key on every sheet.
- **SQL:** every topic whose slide has a SQL panel (23 queries, mostly in 3.2–3.4 and 3.10) gives that query beside a worksheet formula twin, so the learner can check the number SQL returns. `Week3_Data_Files/ETEH_Week3_SQL_Examples.sql` is regenerated from the same queries.
- **DAX topics (3.6A/3.6B):** each measure's result has a worksheet formula twin (mostly SUMIFS/COUNTIFS) that returns the same value for the stated filter. Small helper tables (tCat, tYear) play the part of a PivotTable's rows.
- **Charts (3.7):** summary tables built from formulas (tZoneSum, tMonth, tTier, tPay, tDept, tBins, tRadar, tStoreMonth) feed real Excel charts: column, bar, line, area, pie, doughnut, scatter, radar and combo, plus a sparkline per store. The file format the build uses cannot store bubble charts or the Excel 2016 chart types (waterfall, funnel, treemap, sunburst, histogram, box and whisker). For those, the sheet has the prepared data and the steps, and the learner inserts the chart in Excel.
- **Dashboard (3.8):** a zone selector cell (a drop-down) drives the KPI cards, in place of a slicer.
- **Verification:**
  - 1,653 formulas checked with `_build/verify_xlsx.py`: 1,104 recalculate to their cached values and none disagree;
  - the other 549 are the FORMULATEXT column, which pycel cannot evaluate;
  - every file passes `check_xml.py`;
  - pycel gained array-aware SUMIFS and COUNTIFS, plus COUNTA and UNIQUE, so SUMPRODUCT over a criteria list can be checked.

### Week 3 assignment
- The tasks and marks are unchanged.
- A new **Function Guide** sheet shows the worksheet patterns the tasks use: an XLOOKUP calculated column, year-filtered COUNTIFS/AVERAGEIFS, a month count bounded by EOMONTH, and census rates. Each is a live example on the assignment's own tables that answers a different question from the tasks.
- **Bug fixed in the key:** the Task 3 formula grid held two `=XLOOKUP([@ward_id],…)` formulas outside any table. Excel cannot resolve `[@…]` there, so it would repair the file on opening. They are now shown as text. A new scan (`_build/rev/scan_thisrow.py`) checks every pack workbook for this; no other file has it.
- The reference dashboard in the key now counts months with EOMONTH, which Week 3 teaches, instead of EDATE, which no week teaches.
- An untaught-function scan (`_build/rev/untaught_scan.py`) finds no function in the student file or the key that the Week 1–3 workbooks do not teach.
- Both files verify with 0 mismatches. The student file has 24 formulas (17 recalculated, 7 FORMULATEXT). The key has 43,677 formulas (43,670 recalculated, 7 FORMULATEXT).

### Build
```
python3 _build/07_v3_week1_deck.py --preview   # deck (+ PNG previews in _build/_stage/v3_preview)
python3 _build/rev/lint_v3.py                  # deck lint
python3 _build/02_sec_w1.py                    # examples workbook (asserts all 111 topics are covered)
python3 _build/03_asg_w1.py                    # assignment + answer key
python3 _build/07_v3_week2_deck.py && python3 _build/rev/lint_v3.py 2
cd _build && python3 02_sec_w2.py              # Week 2 workbook (asserts all 175 topics are covered)
python3 _build/03_asg_w2.py                    # Week 2 assignment + answer key
python3 _build/07_v3_week3_deck.py && python3 _build/rev/lint_v3.py 3
cd _build && python3 02_sec_w3.py              # Week 3 workbook + SQL examples file (asserts all 158 topics are covered)
python3 _build/03_asg_w3.py                    # Week 3 assignment + answer key
python3 _build/rev/untaught_scan.py <Week3 file.xlsx>   # functions not taught in the Week 1-3 workbooks
python3 _build/verify_xlsx.py <file.xlsx>      # independent formula check
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
| ETEH_Vol2_Week2_PowerQuery_Pivots_Slides.pptx | 219 (edition 3) | 2.1–2.11 · 175 topics | ETEH_Week2_Section_Examples.xlsx, Week2_Data_Files, ETEH_Week2_Assignment.xlsx |
| ETEH_Vol3_Week3_SQL_DAX_Dashboards_Slides.pptx | 200 | 3.1–3.10 · 158 topics | ETEH_Week3_Section_Examples.xlsx, Week3_Data_Files (SQLite), ETEH_Week3_Assignment.xlsx |
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
