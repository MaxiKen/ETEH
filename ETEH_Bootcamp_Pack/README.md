# ETEH 4-Week Excel Data Analytics Bootcamp: Delivery Pack

Eaglet2EagleHub Ltd · Instructor: Abdulrahaman Kehinde

## Edition 3.1: follow the learning, find every file

Edition 3.1 rebuilds every learner workbook around one rule: **a question never asks for something that has not been
taught yet**. The decks are unchanged apart from the folder paths they quote.

### Which file do I open? (students)

| When | Open | Where |
|---|---|---|
| During the week, after each module's slides | that module's sheet in the week's **Section Examples** workbook | `01_Section_Examples › Week_N › ETEH_WeekN_Section_Examples.xlsx` |
| At the end of the week | the week's **Assignment** | `02_Assignments › Week_N › ETEH_WeekN_Assignment.xlsx` |
| Week 5 (after Week 4) | the **Capstone** workbook, then one of the four **projects** | `01_Section_Examples › Week_5_Capstone`, `04_Student_Projects › Week_5_Capstone › ProjectN_…` |

Tutors: the assignment answer keys are in `03_Instructor_Projects › Week_N`; the capstone solutions and project keys are in
`03_Instructor_Projects › Week_5_Capstone`.

### Folder layout
Each of the four top folders holds **only one sub-folder per week**. Each week's files live in that week's folder, and
every exercise that needs outside files has its **own folder next to the workbook**.

```
01_Section_Examples/
  Week_1/  ETEH_Week1_Section_Examples.xlsx
  Week_2/  ETEH_Week2_Section_Examples.xlsx   Ex2.1_Files/  Ex2.2_Files/  Ex2.3_Files/  Ex2.4_Files/  Ex2.6_Files/
  Week_3/  ETEH_Week3_Section_Examples.xlsx   Ex3.1_Files/  Ex3.2_Files/  Ex3.4_Files/  Ex3.10_Files/
  Week_4/  ETEH_Week4_Section_Examples.xlsx   Ex4.14_Files/
  Week_5_Capstone/  ETEH_Capstone_Section_Examples.xlsx
02_Assignments/
  Week_1/  ETEH_Week1_Assignment.xlsx
  Week_2/  ETEH_Week2_Assignment.xlsx   Assignment_Files/  (Mercyland_EMR_Export_2023.csv)
  Week_3/  ETEH_Week3_Assignment.xlsx   Assignment_Files/  (mercyland_hospital.db + schema)
  Week_4/  ETEH_Week4_Assignment.xlsx
03_Instructor_Projects/
  Week_1 … Week_4/  ETEH_WeekN_Assignment_ANSWER_KEY.xlsx
  Week_5_Capstone/  Capstone1_SuperMart_Stores/  Capstone2_Kingsway_University/  Capstone3_Mercyland_Hospital/
                    Student_Project_Keys/
04_Student_Projects/
  Week_5_Capstone/  Project1_SuperMart_Online_SQL/  Project2_AtlanticPearl_Hotel/
                    Project3_Mercyland_Diabetes_Screening/  Project4_Talent_Salary_Benchmark/   (each with its own Data/ folder)
```

### Naming standard for linked data
- Every sheet states its data as **File › Sheet › Table** in the gold box at the top, for example
  `File: ETEH_Week2_Section_Examples.xlsx › Sheet: Data_Stores › Table: tStores`.
- A data sheet is called `Data_<Name>` and holds exactly one Excel table called `t<Name>`, starting in cell A1
  (one header row, then one row per record). A two-digit suffix is the year, e.g. `tStores25` = 2025.
- Outside files are named as **Folder › File**, e.g. `Ex2.1_Files › SuperMart_Sales_2025_DIRTY.csv`. The folder sits next
  to the workbook. Section Examples use `Ex<module>_Files`, assignments use `Assignment_Files`, projects use `Data`.

### How every workbook works
- `Start_Here` (first tab): five steps, the sheet map in teaching order (with each sheet's data and question types) and
  the list of data sheets.
- One sheet per module (Section Examples) or per task (assignments). The header says what you have just learned and which
  data to use. Questions come in short parts: multiple choice, Yes/No, True/False, choose-from-a-list, short answers,
  numbers, and "do it in Excel, then type the result".
- Answer in the white cell with the gold border (column E). Column F checks the answer straight away (✓ Correct /
  ✗ Try again). Assignments show marks instead (2 for a number or Excel task, 1 for anything else).
- **Solutions** (last tab of every Section Examples workbook and every answer key): every answer, a one-line explanation
  and, for numbers, the live formula that produces it.

### Progression (what each week may ask)
| Week | Starts with | Calculations allowed |
|---|---|---|
| 1 | terms: data, information, data analytics (1.1–1.4 have no numbers at all) | the status bar, then SUM / AVERAGE / COUNT / COUNTA / MIN / MAX (1.7), then filter + Total Row (1.8) |
| 2 | file formats and Power Query | from 2.8A: SUMIFS / COUNTIFS / AVERAGEIFS, ROUND, IF, TEXT; from 2.8B: lookups, dates, FILTER / UNIQUE; then pivots |
| 3 | database words (no SQL in 3.1) | SQL one clause at a time (each question gives the query); every DAX answer has a Week 2 formula twin |
| 4 | descriptive statistics, one function per question | CORREL, SLOPE, INTERCEPT, RSQ, FORECAST.LINEAR, simple what-if models |
| 5 | the project lifecycle | Weeks 1–4 skills only (SUMIFS / COUNTIFS / AVERAGE on the capstone tables) |

The Week 1 assignment no longer has a "Function Guide": it uses Week 1 skills only (the filter and Total Row instead of
SUMIFS). Each assignment's "Do it" questions use the files in its own `Assignment_Files` folder, and the later tasks use a
cleaned copy of the data inside the workbook, so a mistake in an early task does not spoil the rest.

### Size and checks
| Workbook | Sheets | Questions | Data sheets (tables) |
|---|---|---|---|
| Week 1 Section Examples | 9 modules | 105 | tStores25, tStudents, tSales, tDirtySales |
| Week 2 Section Examples | 12 modules | 106 | tStores, tSales25, tDirty25, tSM |
| Week 3 Section Examples | 11 modules | 92 | tTransactions, tProducts, tCustomers, tDimDate |
| Week 4 Section Examples | 14 modules | 95 | tSalary, tStudents, tMonthly, tStoreBench |
| Capstone Section Examples | 9 modules | 52 | tSM, tStudents, tAdmissions, tCensus |
| Week 1 Assignment · SuperMart Q1 2025 store review | 5 tasks | 46 (64 marks) | tSalesQ1, tStores |
| Week 2 Assignment · Mercyland admissions: import, clean, summarise | 4 tasks | 32 (48 marks) | tWardMap, tAdm |
| Week 3 Assignment · Mercyland: SQL, Data Model and DAX | 5 tasks | 35 (52 marks) | tAdm23, tWards, tDoctors, tDate23 |
| Week 4 Assignment · statistics, forecasting, what-if, ethics | 5 tasks | 32 (50 marks) | tSalary, tStudents, tHotelMonthly |

- `_build/learn/selftest.py` types every correct answer and then a wrong answer into each workbook and recalculates the
  checks: **0 problems** in all nine workbooks (one Week 2 FILTER question is a dynamic array that the checker cannot
  spill; Excel 365 can).
- `_build/verify_xlsx.py`: every formula matches its cached value (0 mismatches) in all Section Examples, answer keys,
  capstone solutions and project keys.
- Deck lint (`_build/rev/lint_v3.py 1…5`): 0 issues. `_build/rev/check_facts.py`: 20 facts, 0 missing.

### Build (edition 3.1)
```
python3 _build/01_build_datasets.py          # source data -> _build/_stage/source (not shipped in the pack)
python3 _build/10_learner_workbooks.py       # Section Examples, Weeks 1-4 and Capstone (or: ... 1 2 3 4 5)
python3 _build/11_learner_assignments.py     # assignments + answer keys, Weeks 1-4
python3 _build/04_proj_i1.py  (i2, i3, s1-s4)  # capstone solutions and the four student projects
python3 _build/07_v3_week1_deck.py  (week2 … week4, capstone)  && python3 _build/rev/lint_v3.py N
python3 _build/learn/selftest.py <workbook.xlsx>   # answer checks
python3 _build/verify_xlsx.py <workbook.xlsx>      # formula check
```
The edition 3.0 workbook and assignment builders are kept in `_build/legacy/` for reference only.

## Edition 3.0 (history): leaner decks, practical work in the workbooks

> The workbook, assignment and folder descriptions in this section describe edition 3.0. They are replaced by edition 3.1 above; the deck notes still apply.


Edition 2 explained every topic with two fixed slides, which made the decks long and repetitive. Edition 3 was rolled out one week at a time and is now **complete for all five volumes**: Weeks 1–4 (deck, examples workbook and assignment for each) and the Capstone deck (Volume 5).

**Download:** this edition is packaged as `ETEH_Bootcamp_Pack_v3.zip` at the repository root. Editions 1 and 2 are kept unchanged in `ETEH_Bootcamp_Pack_v1.zip` and `ETEH_Bootcamp_Pack_v2.zip`.

### New organisations and datasets (edition 3)
Three datasets that were reworked from public Kaggle files are replaced by **original data from new organisations**, generated from scratch by `_build/neworgs.py` (fixed seed, no third-party rows). Every deck, workbook, assignment, project and capstone is rebuilt on them, and every figure on a slide is computed from the files.

| Old (editions 1-2) | New (edition 3) | Files (`01_Section_Examples/00_Core_Datasets`) |
|---|---|---|
| Lifecare General Hospital | **Mercyland Specialist Hospital, Abuja**: 200 staffed beds in four wards (W1 Iroko, W2 Mahogany, W3 Obeche, W4 Baobab) | `Mercyland_Inpatient_Episodes_2019_2024.csv` (58,856 episodes; 12,626 in 2023), `Mercyland_Patients.csv`, `Mercyland_Consultants.csv`, `Mercyland_Wards.csv`, `Mercyland_Daily_Bed_Census.csv`, `Mercyland_Diabetes_Screening.csv` (60,000) |
| Crestview University | **Kingsway University, Ilorin** | `Kingsway_Student_Results.csv` (5,840 students), `Kingsway_Student_Welfare_Survey_SENSITIVE.csv` (9,600), `Kingsway_Digital_Habits_Survey.csv` (640) |
| Eko Suites Hotel, Victoria Island | **Atlantic Pearl Hotel, Lekki**: 320 rooms | `AtlanticPearl_Reservations_2023_2024.csv` (38,400 reservations) |

- SuperMart Stores, SuperMart Online and the ETEH Talent Salary Survey are unchanged.
- **Mercyland 2023:** average stay 5.02 days, 30-day readmission 8.7%, bed occupancy 86.9% (June peak 98.1%; 60 days above 95%). The Week 2 dirty extract (`Week2_Data_Files/Mercyland_Episodes_2023_DIRTY.csv`, 12,772 rows) has duplicate rows, ward spelling variants, mixed gender codes, bills stored as text, negative bills and impossible dates. The documented cleaning rules turn it back into the clean file exactly, and the Capstone builders check this.
- **Kingsway:** mean final score 66.3; 1,591 students (27.2%) are at risk. Attendance (r = 0.47) and study hours (r = 0.31) are the strongest drivers. Engineering has the highest at-risk rate (36.8%), Arts & Education the lowest (20.4%).
- **Atlantic Pearl:** 33.5% of reservations are cancelled (OTA 37.8%, Corporate 4.5%). The room-type table shows Simpson's paradox: Deluxe cancels less than Twin within each year, but more overall.
- The SQLite database is now `Week3_Data_Files/mercyland_hospital.db`. Its tables are `wards`, `doctors`, `patients`, `admissions` and `bed_census`, plus the view `v_admissions`.
- ID formats: `ADM2023-00001` (episode), `MSH-P000001` (patient), `CON001` (consultant), `KU/2024/EEE/0001` (matric number), `APH-00001` (reservation).

### Week 1 deck (`ETEH_Vol1_Week1_Foundations_Slides.pptx`, 146 slides, 111 topics)
- Built from the ETEH deck template (title, welcome, module divider, content frame, closing).
- **One slide per topic, laid out for what the topic needs.** There is no fixed template. Where it helps, a topic slide has, in this order:
  - the **standard definition**;
  - **in simple terms**;
  - **an analogy**;
  - and, only where it is useful, a short **example from our data**.
- Example tables are illustrations only. Each one is labelled with its **table name** (for example `tStores25`, `tDept23`, `tHotelSeg`), not a file name.
- Each module opens with a divider and "What you will learn", and closes with "Key takeaways". The week ends with a summary and a two-slide glossary.
- **Data vs information (1.1.1–1.1.3), clarified:** data was defined as "raw, *unorganised* facts", yet the 1.1.1 example is a neat table with headers, so learners could fairly call it information. The definitions now separate **structure** from **processing**. A table of till lines with headers is still *structured data*, because nothing has been summed, compared or interpreted. The 1.1.2 store-totals table is information because 3,312 FY2025 lines were summed and compared with targets. Slide 1.1.3 adds the test ("has anything been summarised, calculated or compared to answer a question?"). The glossary, the takeaway and the workbook's 1.1 concept note use the same wording.
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
  - 825 formulas checked with `_build/verify_xlsx.py`: 567 recalculate to their cached values;
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
  - 14 data sheets, each with one table starting in A1: SuperMart Online (tTransactions, tLines, tProducts, tCustomers, tSellers, tDeliveries, tStates, tDimDate), SuperMart Stores FY2025 (tDash, tDashStores), Mercyland (tAdm23, tCensus, and tMsDirty, the raw 2023 extract as received) and Kingsway (tCV).
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

### Week 4 deck (`ETEH_Vol4_Week4_Stats_Forecast_AI_Slides.pptx`, 188 slides, 139 topics)
- The same one-slide-per-topic design, for modules 4.1–4.14: descriptive statistics, distributions, correlation, regression, forecasting, what-if analysis, storytelling, report writing, presentations, auditing, the AI co-pilot, AI verification, data ethics and benchmarking.
- Every figure is computed from the edition 3 datasets when the deck is built (`_build/w4_data.py`, `_build/v3/w4_fx.py`); `_build/rev/check_facts.py` confirms the headline facts appear on the slides.

### Week 4 examples workbook (`01_Section_Examples/ETEH_Week4_Section_Examples.xlsx`)
Rebuilt from scratch in the Week 1 format (`_build/02_sec_w4.py` with `_build/v3/wb4_common.py`, `wb4_a.py` for 4.1–4.4, `wb4_b.py` for 4.5–4.9 and `wb4_c.py` for 4.10–4.14). The edition 2 builder is kept as `_build/rev/02_sec_w4_edition2.py`.
- **Sheets:**
  - one sheet per module (4.1–4.14), plus **4.10 Audit Case**, a zone summary with the planted errors the 4.10 slides describe (a SUM that misses the last zone, a typed share, and a share divided by itself) beside the corrected column and an OK/ERROR check;
  - 10 data sheets, each with one table starting in A1: tSal (ETEH Talent Salary Survey), tOrders (SuperMart Online lines), tCV (Kingsway), tLOS and tCensus23 (Mercyland), tHotel (Atlantic Pearl), tDash25, tMonthly, tCustRev and tStores (SuperMart Stores).
- **All 139 mini-sections are covered** (the build asserts it). Each has Steps, a Worked solution with live formulas, and "Your turn" work: **171 exercises with live ✓/✗ checks**, plus written and do-it-in-Excel tasks, with an answer key on every sheet.
- **Statistics (4.1–4.4):** AVERAGE/MEDIAN/MODE, STDEV.S, QUARTILE.INC, SKEW and KURT; STANDARDIZE, NORM.DIST and confidence intervals; CORREL, RSQ and a rank-based (Spearman) check; SLOPE, INTERCEPT, FORECAST.LINEAR and LINEST with residuals.
- **Forecasting (4.5):** moving averages, a linear trend, seasonal indices, a simple exponential smoothing column on Data_Monthly, FORECAST.ETS with its confidence band, and MAE/MAPE/RMSE to compare same-month-last-year, straight-line and decomposition forecasts.
- **What-if (4.6):** an Atlantic Pearl profit model driven by input cells, with Goal Seek, a scenario table and a Solver layout (the inputs and constraints are on the sheet; the learner runs Solver in Excel).
- **Storytelling, reports and slides (4.7–4.9):** formula-driven chart titles that state the finding (for example the SuperMart store furthest behind its 2025 target), and report/slide checklists tied to the numbers on the sheet.
- **Auditing, AI and ethics (4.10–4.14):** error-finding with Trace Precedents, Show Formulas and FORMULATEXT on the audit case; AI claims checked against live formulas (tClaims); data ethics on the Mercyland extract (column minimisation, small-cell suppression for patients aged 95+) and a Talent Survey pay-gap test by seniority; and a benchmark lab that reconciles the SuperMart monthly sales bulletin PDF with the sales ledger, store by store (tBulletin).
- **Charts:** 15 real Excel charts (9 column/bar, 4 line, 2 scatter), each fed by a formula-built table. Histograms, box plots and waterfalls use Excel 2016 chart types the build cannot write, so the sheet has the prepared data and the steps.
- **Verification:**
  - 1,567 formulas checked with `_build/verify_xlsx.py`: 1,106 recalculate to their cached values and none disagree;
  - of the other 461, 395 are the FORMULATEXT column, 54 are FORECAST.ETS cells and the band built on them, and 12 use things pycel cannot read (such as `tSal[#Headers]`);
  - every file passes `check_xml.py` and `_build/rev/scan_thisrow.py`;
  - pycel gained MODE (and MODE.SNGL) for the 4.1 checks.

### Week 4 assignment
- The 11 tasks follow the Week 4 modules in order.
- A new **Function Guide** sheet shows the patterns the tasks need (PERCENTILE.INC, STANDARDIZE, NORM.DIST, CORREL, RSQ, SLOPE, FORECAST.LINEAR, INDEX(LINEST(…)), APE and a date-window SUMIFS). Each is a live example on the assignment's own tables that answers a different question from the tasks.
- **Updated for the new data:**
  - **Task 5:** the Atlantic Pearl series now runs a full 18 months (July 2023 to December 2024). It peaks every December and has no opening ramp-up. Training is now July 2023 to June 2024, the test is July to December 2024, and the forecast is for January 2025 (Month_No 19). Q5.1 now asks about the December seasonal peak, replacing the old "incomplete first month" question.
  - **Claims that now come from the data:** the coefficient in Q4.6, the Q7.1 headline (Q4 room-nights fell 14% year on year), the Q7.5 model summary and the Task 9 AI claims. The Task 9 answers are now FALSE, TRUE, TRUE, FALSE.
- The untaught-function scan now checks each assignment against the examples workbooks up to its own week. It reports functions taught only on an assignment's Function Guide separately: EDATE in Week 1, WEEKDAY in Week 2, none in Weeks 3–4.
- Both files pass `check_xml.py` and the `[@]` scan, and verify with 0 mismatches. The student file has 54 formulas (39 recalculated) and the key has 138 (123 recalculated). In each, the other 15 are 14 FORMULATEXT cells and one cell that uses a defined name pycel cannot read.

### Capstone deck (`ETEH_Vol5_Capstone_Career_Slides.pptx`, 120 slides, 80 topics)
- Rebuilt in the edition 3 one-slide-per-topic design for P.1–P.9 (`_build/07_v3_capstone_deck.py`; copy in `_build/v3/w5_m1_2.py` … `w5_m8_9.py`). The edition 2 copy is kept in `_build/rev/t5_*.py`.
  - P.1–P.2: the project lifecycle and source governance, with each capstone's brief, data inventory and source register.
  - P.3 Capstone 1 SuperMart: FY2025 sales ₦733.2m (+20.4%) but margin 13.4% › 12.7%. Discount bands, a rate/mix margin bridge (rate −1.19 pts, mix +0.51 pts), the Q4 2025 margin of 9.8%, SKU and basket analysis, and the Furniture 20% discount cap (+₦17.7m–₦18.0m a year).
  - P.4 Capstone 2 Kingsway: 5,885 raw rows cleaned to 5,840 students; progression 72.8% against a 90% target; 1,591 at risk (27.2%) by trigger; at-risk rate by faculty and attendance band; the tutorial-programme estimate (80.7%, +462 students, R² 0.34, labelled as an estimate).
  - P.5 Capstone 3 Mercyland FY2023: 12,626 admissions, bed-day occupancy 86.9% (June 98.1%; 60 days above 95%, 24 above 100%; about 5 extra beds for 85%); LOS mean 5.0, median 4 (8+ day stays are 18.2% of admissions but 42.2% of bed-days); readmission 8.7%; the Q4 escalation list (444 of 3,125 Q4 admissions: 275 readmitted within 30 days, 119 stays of 14+ days, 50 aged 75+).
  - P.6 workbook QA, P.7 executive defence, P.8 portfolio and CV, P.9 interviews (live Excel and SQL tests on the pack's own data, STAR) and the roadmap to Power BI, SQL and Python.
- **Charts:** 23 native charts on 17 slides: column, bar, line and combo charts, a margin-bridge waterfall, and a one-screen dashboard (KPI cards plus three charts) for each capstone.
- **Data:** every figure is computed from the edition 3 datasets when the deck is built (`_build/v3/w5_fx.py` reuses the capstone builders `04_proj_i1.py`–`04_proj_i3.py`). The table names and columns quoted on the slides (tSM, tCV, tAtRiskTop20, tAdm23, tCensus, …) are the ones in `ETEH_Capstone_Section_Examples.xlsx`.
- **Checks:** `_build/rev/lint_v3.py 5` reports 0 issues (topic order and counts, footers, glyphs, formula spacing, workbook references). `_build/rev/check_facts.py` now also checks the Capstone deck: it recomputes the Kingsway and Mercyland headline facts straight from the CSVs, independently of the deck pipeline (20 facts, 0 missing).
- **Wording fixes:** "days above 95%" (60) and Volume 1's "95% or higher" (63) are both correct; three 2023 days sit at exactly 95.0%. Kingsway's at-risk students are described by rate (Engineering 36.8%, the highest) instead of "most in Engineering", since Engineering holds only 28% of them.

### Data standards
- **SuperMart Online `sellers.store_name` = the registered trading name.**
  - Rules: unique across the marketplace, enforced by `UNIQUE` in `supermart_online.db` and its schema script; no numeric suffixes; any place in a name is the store's own city.
  - Before: 200 sellers shared 86 names (e.g. 11 "Trade Fair Online", 8 "Alaba International Digital"). Many were named after cities they are not in, e.g. "Sokoto Suppliers" in Oshogbo.
  - Names now follow common Nigerian SME patterns: "Yusuf Ventures", "Nnaji Stores", "Makurdi Trading Company". They are generated deterministically by `seller_trading_names()` in `_build/01_build_datasets.py`.
  - Only the names changed: IDs, cities, sales and every total are as before.
  - The core CSV, the database, the SQL scripts, the Week 2 JSON and the raw export were regenerated.
  - **Effect on answers:** Project 1's "seller with the highest 2025 net revenue" was "Alaba International Digital", which in fact merged eight different sellers. The answer is now Kaduna Merchants (SELL0200), and the key groups by `seller_id, store_name`. The Capstone P.9 window-function query groups the same way.
  - **Teaching notes:**
    - Vol3 3.1.8 uses the new `UNIQUE` constraint as its example.
    - The 3.4.10 warning now uses owner names, which still repeat (191 names for 200 sellers), to show why names make poor keys.
    - 3.6.28 says to rank by `seller_id` because stores can be renamed.
    - The 3.4.8 fuzzy-matching example now types 3 of 8 names in a different word form ("Yusuf Venture", "Trading Co."), so 5 of 8 match after ignoring case and spaces, and 3 need fuzzy matching.
- **Capstone 3 escalation rule: long stay = 14+ days**, NHS England's long-stay measure (it tracks 14+ and 21+ day stays). The old 28-day cut-off flagged only 8 Q4 cases. The new one flags 119, a workable weekly review list. The rule order is unchanged: readmission first, then long stay, then age 75+. The cut-off is one constant (`ESC_LOS`) used by the Capstone workbook (P.5.7 conditional column, Power Query code and counts) and the Vol5 deck.

### Clarity pass (all files)
- **Hotel channels (Vol1 1.3.8; Week 1 workbook 1.4):** "three times the corporate rate" was wrong. Online travel agents cancel 37.8% against Corporate's 4.5%, about **eight times**. The workbook note is now computed from the ratio.
- **Wrong-join control total:** two correct totals were used without saying which. **₦1.32bn** is the 10,430 lines matched on the right key (`delivery_id`); **₦1.39bn** is the ledger revenue across all 11,000 orders. Week 3 compares the two keys, so Vol1 1.3.12 now says "the right key gives ₦1.32bn". The AI zone-query story (Vol4 4.11–4.12, Vol5 P.8–P.9, the Capstone workbook P.8–P.9, the Project 1 key) now uses ₦1.39bn and 5.2× throughout. The fix is the same everywhere: drop the deliveries join, which a revenue query does not need. Before, the Vol5 and Capstone STAR answers said it was "rewritten on delivery_id", which contradicted Vol4.
- **Week 3 workbook 3.4.10:** the row was labelled "Join on order_id"; it now says "Left join on delivery_id". Its note no longer says "one delivery per order line", because the 570 cancelled lines have no delivery.
- **Week 2 answer key 5.8:** the sample headline said all four wards admitted "similar volumes (about 2,600–2,800 each)". In fact Iroko admitted 4,062 and Mahogany 1,588. The headline is now computed from the data.
- **Vol4 4.3.9:** Simpson's paradox is now explained as one group concentrated in the subgroup with the highest rate, not as "groups of very different sizes".

### Build (edition 3.0)
See "Build (edition 3.1)" above; the 3.0 workbook builders now live in `_build/legacy/`.

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
   - **The situation**: a real business question on the pack datasets (SuperMart, Kingsway, Mercyland, Atlantic Pearl, the Talent Survey or SuperMart Online; edition 2 used Crestview, Lifecare and EkoSuites, which edition 3 replaced).
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
| ETEH_Vol2_Week2_PowerQuery_Pivots_Slides.pptx | 220 (edition 3) | 2.1–2.11 · 175 topics | ETEH_Week2_Section_Examples.xlsx, Week2_Data_Files, ETEH_Week2_Assignment.xlsx |
| ETEH_Vol3_Week3_SQL_DAX_Dashboards_Slides.pptx | 203 (edition 3) | 3.1–3.10 · 158 topics | ETEH_Week3_Section_Examples.xlsx, Week3_Data_Files (SQLite), ETEH_Week3_Assignment.xlsx |
| ETEH_Vol4_Week4_Stats_Forecast_AI_Slides.pptx | 188 (edition 3) | 4.1–4.14 · 139 topics | ETEH_Week4_Section_Examples.xlsx, ETEH_Week4_Assignment.xlsx |
| ETEH_Vol5_Capstone_Career_Slides.pptx | 120 (edition 3) | P.1–P.9 · 80 topics | ETEH_Capstone_Section_Examples.xlsx, Capstones 1–3 (solution workbooks), Student Projects 1–4 |

**How to study a topic:** read the concept slide first. Then open the workbook sheet named under "Try it yourself" and follow the three steps. Check that your answers match the "What you should see" table. If your numbers differ, look at the **Avoid** box first.

### Edition 1 changes from the original volumes (still apply)
- The final two lab and practice modules of every volume were removed. Roadmaps, topic counts and the "In this session" panels were updated to match.
- Duplicate module codes were renumbered so slide codes match the worksheet names in the section workbooks:
  - Week 2 now runs 2.1–2.7, 2.8A, 2.8B, 2.9–2.11.
  - Week 3 keeps 3.6A and 3.6B.
- External sources were removed: Wikipedia tables, CBN/NBS figures and "live source" links. The modules that depended on them now use files in the pack:
  - 2.4 · Web and PDF sources: the SuperMart store directory HTML page and the monthly sales bulletin PDF.
  - 3.10 · File extracts into SQL, DAX and dashboards: the Mercyland 2023 dirty CSV (`Week2_Data_Files/Mercyland_Episodes_2023_DIRTY.csv`) and `Week3_Data_Files/mercyland_hospital.db`.
  - 4.14 · Internal benchmarks, PDF checks and reconciliation: the SuperMart ledger and bulletin, and Mercyland readmission intervals.
  - P.2 · Source governance: a register of the capstone sources.
- The Capstone 1–3 modules (P.3–P.5) are rewritten around the instructor capstones. Every figure on those slides comes from the solution workbooks.
- Headline worked examples elsewhere now quote real dataset figures, for example:
  - SuperMart FY2025: ₦733.2m sales, 12.7% margin, 3,312 lines.
  - Kingsway: 5,840 students, 1,591 at risk (27.2%).
  - Mercyland FY2023: 12,626 admissions, 86.9% occupancy.
- Week 1 dollar amounts were converted to naira. `$` signs inside cell references such as `$B$2` were kept.

## Folders
See **Folder layout** under Edition 3.1 at the top. The source CSVs, databases and practice files are generated into
`_build/_stage/source/` by `_build/01_build_datasets.py`; learners get exactly the files each exercise needs, copied into
that exercise's own folder.

All files are generated by the scripts in `_build/` at the repository root.
- `05_update_slides.py` rebuilds the edition 1 decks from the original volumes.
- `06_revise_slides.py` builds the edition 2 decks from the topic copy in `_build/rev/`.
