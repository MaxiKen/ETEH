# ETEH Four-Week Excel Data Analytics Bootcamp

A complete, self-contained data-analytics bootcamp built from the original five
presentation volumes (Weeks 1–4 + Capstone & Career). The original decks were
**cleaned**, the supplied raw datasets were **standardised, dirtied and reformatted**,
and a full set of **working files** was generated so the programme can be delivered
as-is.

---

## What changed vs. the source decks
- **Removed the non-learning "lab" modules** (the user-confirmed true labs):
  `1.10 Practical Labs…Capstone`, `2.12 Week 2 Practical Exercises…`,
  `3.11 Labs, Capstone & Desk Reference`, `4.15 Labs, Capstone & Desk Reference`.
  Adjacent teaching modules (`1.11`, `2.13`, `3.12`, `4.16`) were kept because they
  contain real curriculum. Vol. 5 (Capstone & Career) was kept in full.
- **Retargeted the three "external" datasets** named in the decks
  (*SuperMart Retail, Crestview University, Lifecare General Hospital*) to the
  **attached, standardised datasets** — inline worked-example tables were kept and
  only the dataset name was swapped:
  - SuperMart Retail → **ETEH Retail** (`Sample - Superstore`)
  - Crestview University → **ETEH University** (Student Performance / Depression / Social-Media)
  - Lifecare General Hospital → **ETEH Lifecare Hospital** (`healthcare_dataset`)
  - The relational **e-commerce** workbook became **ETEH Commerce** (SQL).
- Generated a complete delivery kit (datasets + SQL + 4 operation folders).

## Folder layout
```
ETEH_Bootcamp/
├── README.md
├── presentations/      # 5 cleaned decks (*_CLEAN.pptx); originals untouched at repo root
├── datasets/          # the heart of the bootcamp
│   ├── DATASET_GUIDE.md
│   ├── clean/         # standardised "gold" copies (CSV + XLSX)
│   ├── dirty/         # intentionally dirtied working copies for cleaning labs
│   ├── sql/           # eteh_commerce.db (SQLite, FKs) + queries.sql + expected outputs
│   └── pdf/           # printable data dictionary + sample extracts
├── section_examples/  # one xlsx per week; each module = a worksheet of relatable worked examples
├── assignments/       # one xlsx per week; Brief + Practice_Data + Rubric sheets
├── projects_instructor/ # Instructor_Projects.xlsx — 3 fully worked solutions
├── projects_students/   # Student_Projects.xlsx — 3 capstone briefs
└── build/            # reproducible Python build scripts (run in order)
```

## Datasets used (all attached files; best-of selection)
| ETEH name | Source file | Format | Used in |
|---|---|---|---|
| ETEH Retail | `Sample - Superstore.csv` | CSV, XLSX | W1, W2 Pivots, W3 Dashboards, W4 Forecast |
| ETEH Commerce (SQL) | `e-commerce.xlsx` (5 joined tables) | **SQLite + CSV + XLSX** | W3 SQL/DAX/Dashboards (primary relational set) |
| ETEH Lifecare Hospital | `healthcare_dataset.csv` | CSV, XLSX | W1, W4 Stats, W3 Dashboards |
| ETEH University | `StudentPerformanceFactors` + `Student Depression` + `Students Social Media` | XLSX (3 sheets) | W4 Correlation/Regression |
| ETEH Hotels | `Hotel Reservations.csv` | CSV, XLSX | W2 Cleaning, W4 Cancellation analysis |
| ETEH Salary | `Salary_Data.csv` | CSV, XLSX | W4 Linear Regression |
| ETEH Health (Diabetes) | `diabetes_prediction_dataset.csv` (10k sample) | CSV, XLSX | W4 Distributions/Classification |

Every dataset ships as a **clean** copy and a **deliberately dirty** copy (mixed
date formats, inconsistent casing, duplicate rows, blank cells, text-formatted
numbers, currency strings, broken keys) so learners can practise real cleaning.
The dirty-issue log for each dataset is in `datasets/DATASET_GUIDE.md` and the
printable `datasets/pdf/DATASET_DICTIONARY.pdf`.

## E-commerce → SQL
`datasets/sql/eteh_commerce.db` is a relational SQLite database with enforced
primary/foreign keys across **customers · sellers · products · deliveries ·
transactions**. `queries.sql` demonstrates joins, aggregations and window
functions (monthly revenue, category matrix, top sellers, customer lifetime value,
delivery performance, payment mix, rolling 3-month revenue). Pre-computed expected
outputs (`out_*.csv`) let students reconcile their own queries, and `eteh_commerce_flat.csv`
/ `eteh_commerce_pivot_source.csv` feed the non-SQL Pivot/Dashboard modules.

## How to use
1. Teach from `presentations/*_CLEAN.pptx`.
2. Hand out the matching `section_examples/WeekN_Section_Examples.xlsx` (one worksheet per module).
3. Assign `assignments/WeekN_Assignments.xlsx` (data + rubric included).
4. Run capstones from `projects_students/Student_Projects.xlsx`; model them with `projects_instructor/Instructor_Projects.xlsx`.
5. All data lives in `datasets/` — point learners at `DATASET_GUIDE.md` first.

## Reproducibility
`build/build_datasets.py` → `build/build_sql.py` → `build/build_pdf_guide.py` →
`build/build_workbooks.py` → `build/edit_presentations.py` regenerate the entire
kit from the raw inputs.
