# ETEH Excel Data Analytics Bootcamp — 4-Week Program

A complete, dataset-driven Excel Data Analytics bootcamp built entirely on the **attached datasets**.
Every example, assignment, project and deliverable in this program uses one of the nine provided
datasets — standardised, intentionally dirtied where cleaning practice is required, and extracted
into the forms the course demands (CSV, XLSX, SQL, PDF).

## Program at a glance

| Volume | Week | Focus | Cleaned deck | Section examples | Assignments |
|---|---|---|---|---|---|
| Vol 1 | 1 | Foundations & Excel | 178 slides | 1.1–1.9 | 1.1–1.9 |
| Vol 2 | 2 | Power Query, Pivots, Functions, M | 218 slides | 2.1–2.11 | 2.1–2.11 |
| Vol 3 | 3 | SQL, DAX, Dashboards | 212 slides | 3.1–3.10 | 3.1–3.10 |
| Vol 4 | 4 | Statistics, Forecasting, AI & Ethics | 192 slides | 4.1–4.13 | 4.1–4.13 |
| Vol 5 | 5 | Capstone & Career | 130 slides | P.1–P.10 | P.1–P.10 |

The last two modules of each volume (labs, external-benchmark and other non-learning content)
have been removed. All external-dataset mentions in the presentations have been replaced with
inline examples drawn from the attached datasets.

## Folder structure

```
ETEH_Bootcamp/
├── README.md                          ← this file
├── Presentations/                     ← 5 cleaned PPTX decks (Vol 1–5)
├── Datasets/
│   ├── CLEAN/                         ← 16 standardised CSVs (the analysis truth)
│   ├── DIRTY_FOR_WEEK2/               ← 5 intentionally dirty CSVs + 16-issue log
│   ├── SQLITE/                        ← 4 relational databases + ecommerce_queries.sql
│   ├── XLSX_EXTRACTS/                 ← 9 multi-sheet extracts (500-row capped)
│   └── PDF_DICTIONARIES/              ← 15 dataset dictionaries + DATASET_GUIDE.pdf
├── 01_Section_Examples/               ← folder 1: one XLSX per week, module sheets
├── 02_Assignments/                    ← folder 2: one XLSX per week, multi-sheet
├── 03_Instructor_Projects/           ← folder 3: lecturer worked solutions
└── 04_Student_Projects/               ← folder 4: student project briefs
```

## The 4 delivery folders

### 1. Section Examples (`01_Section_Examples/`)
One workbook per week. Each module (or subsection) is a worksheet with:
- **Concept tables** — the idea, plainly stated
- **Worked examples** with live model formulas (green cells) and yellow practice cells
- **DATA sheets** — 500-row extracts of the relevant datasets for hands-on practice
- **ANSWERS sheets** — instructor marking guides with verified numbers

Files: `Week1_Foundations_Section_Examples.xlsx`, `Week2_PowerQuery_Pivots_Section_Examples.xlsx`,
`Week3_SQL_DAX_Dashboards_Section_Examples.xlsx`, `Week4_Stats_Forecast_AI_Section_Examples.xlsx`,
`Week5_Capstone_Career_Section_Examples.xlsx`.

### 2. Assignments (`02_Assignments/`)
One workbook per week, one sheet per module. Each task states the instruction and the deliverable;
every module sheet names the exact dataset to use and its mark weight. Week 2 and Week 3 assignments
rely on the DIRTY files and the SQLite databases respectively.

### 3. Instructor Projects (`03_Instructor_Projects/`)
`Instructor_Projects_Worked_Solutions.xlsx` — four complete worked projects (E-commerce analytics,
Student Performance analytics, Healthcare operations, Forecasting & benchmark), each with verified
results, interpretation and deliverable checklists, plus a marking reference sheet.

### 4. Student Projects (`04_Student_Projects/`)
`Student_Projects_Briefs.xlsx` — four capstone project briefs (one per track), each with brief,
milestones, success criteria, rules & ethics, and a submission checklist.

## The datasets (all attached, all used)

| Dataset | Rows | Best use in the bootcamp |
|---|---|---|
| e-commerce.xlsx → SQLite | 11,000 txn + 4 dimension tables | **SQL spine** — joinable keys, star schema, DAX, dashboards, capstone Track A |
| Sample - Superstore.csv | 9,994 | Cleaning, pivots, correlation/regression, forecasting (Track D) |
| Salary_Data.csv | 6,704 → 1,791 clean | Distributions, outliers, goal seek, ethics |
| StudentPerformanceFactors.csv | 6,607 | Education analytics (Track B), statistics |
| healthcare_dataset.csv | 55,500 | Operations analytics (Track C), normalization, governance |
| diabetes_prediction_dataset.csv | 100,000 | Distributions, probability, regression, AI tasks |
| Hotel Reservations.csv | 36,275 | Cancellation analysis, forecasting, report/PDF |
| Student Depression Dataset.csv | 27,901 | Correlation, ethics (anonymised only) |
| Students Social Media Addiction.csv | 705 | Statistics practice, correlation |

### Standardisation & intentional dirt
- **CLEAN/** — 16 standardised CSVs: ISO dates, consistent types, ordinal bands, derived flags,
  anonymised versions for sensitive files, dedup and outlier flags.
- **DIRTY_FOR_WEEK2/** — 5 files intentionally corrupted (blanks, text-in-number, mixed casing,
  broken keys, spacer rows) with a 16-issue log. Students clean these and reconcile to CLEAN.
- **SQLITE/** — relational build of the E-commerce data (the only dataset with proper keys) plus
  Superstore, Education and Healthcare databases. `ecommerce_queries.sql` provides the 7 core queries.

## How to use this program

1. **Read** the README, then open the relevant Section Examples workbook alongside the presentation.
2. **Work** the module sheets top to bottom: concept → worked example → your turn (yellow).
3. **Practise** with the DATA extracts; never clean the raw CLEAN files — work on copies.
4. **Assignments** per week are marked against the ANSWERS sheets and the Instructor Projects.
5. **Capstone (Week 5):** pick a track from the Student Projects briefs, build the pack, QA it,
   defend it. Marking uses the rubric in `Week5_Capstone_Career_Section_Examples.xlsx` (10_P.10_Rubric).

## Standards enforced throughout

- **Reconciliation is binary** — a KPI that disagrees with the dataset fails, regardless of presentation.
- **No external datasets** — inline examples only; all figures come from the attached files.
- **Live formulas** — every computed number is a formula, query or verified value; no typed results.
- **Ethics first** — sensitive data (depression, healthcare) used anonymised; purpose and consent noted.
- **Professional output** — consistent navy/teal palette, labelled axes, sources on every data page.
