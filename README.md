# ETEH — Four-Week Excel Data Analytics Bootcamp

A complete, fully-formed analytics bootcamp built from the five-volume slide
programme plus nine real-world datasets. Everything has been standardized,
intentionally "dirtied" for practice, converted to multiple formats (CSV, XLSX,
PDF, SQL), and wired into per-week workbooks and projects.

---

## Folder map

```
ETEH/
├── Datasets/                      # the data layer
│   ├── Raw/                      # verbatim copies of the 9 supplied datasets (provenance)
│   ├── Clean/                    # standardized, analysis-ready CSV + XLSX (one per dataset)
│   │                              #   + ecommerce_*_reconciled tables + data_dictionary.xlsx
│   ├── Dirty/                    # intentionally corrupted copies for cleaning/QA practice
│   │                              #   (see dirty_manifest.md for the exact issues)
│   ├── SQL/                      # ecommerce.db (relational, joined) + all_datasets.db
│   │                              #   + ecommerce_schema.sql + ecommerce_sample_queries.sql
│   ├── Reports/                  # data_catalogue.pdf (printed-ready dictionary)
│   └── README.md                 # what each dataset is and where it is used
│
├── Section_Examples/             # FOLDER 1 — training examples (per week)
│   └── Week1..Week5_Section_Examples.xlsx   # one worksheet per module,
│                                             # relatable worked examples w/ live formulas
│
├── Assignments/                  # FOLDER 2 — weekly assignments (per week)
│   └── Week1..Week5_Assignments.xlsx         # one worksheet per module + Answer Key sheet
│
├── Projects_Instructor/          # FOLDER 3 — the one instructor-led capstone
│   └── E-Commerce & Retail Analytics Deep-Dive (brief + worked SQL solution)
│
├── Projects_Students/            # FOLDER 4 — four student projects
│   ├── project_01_retail_analytics/
│   ├── project_02_healthcare_operations/
│   ├── project_03_marketplace_sql/           (uses ecommerce.db)
│   └── project_04_student_success/
│
├── Presentations_Clean/          # the 5 volumes, cleaned:
│                                 #   • last two "lab" modules removed per week
│                                 #   • external example-brand names replaced with the
│                                 #     real curated datasets used in this bootcamp
│
└── build/                        # reproducible Python generators (run in order)
```

---

## What changed from the original slide decks

1. **Last two modules removed each week** — the non-learning "lab / capstone-drill /
   judgement" modules were stripped so every remaining week is pure learning content.
   *Vol1: 1.10–1.11 · Vol2: 2.12–2.13 · Vol3: 3.11–3.12 · Vol4: 4.15–4.16 · Vol5: P.11–P.12.*
2. **External example datasets replaced** — the recurring fictional brands
   *SuperMart Retail*, *Crestview University* and *Lifecare General Hospital* were
   swapped for the real datasets shipped here, so the inline examples now point at
   data the learner actually has:

   | Original (external)        | Now maps to (real dataset)            |
   |----------------------------|--------------------------------------|
   | SuperMart Retail           | Sample – Superstore (`superstore_sales`) |
   | Crestview University       | Student Performance Factors (`student_performance`) |
   | Lifecare General Hospital  | Healthcare Records (`healthcare_records`) |
   | (new) E-Commerce Marketplace | `e-commerce` → `ecommerce.db` (SQL)  |

---

## The nine datasets (all used)

| Dataset | Rows | Best use in bootcamp | Format |
|---|---|---|---|
| **e-commerce** (5 joined tables) | 37k txns / 2.5k customers | **SQL joins, relational modelling, DAX star schema** | SQLite + CSV + XLSX |
| Healthcare Records | 55.5k | Tidy-data, cleaning, stats, hospital ops capstone | CSV/XLSX |
| Superstore Sales | 10k | PivotTables, dashboards, forecasting, retail capstone | CSV/XLSX |
| Student Performance | 6.6k | Regression, correlation, education capstone | CSV/XLSX |
| Salary Survey | 6.7k | Regression, compensation analysis | CSV/XLSX |
| Student Depression | 27.9k | Classification, group comparison | CSV/XLSX |
| Diabetes Prediction | 100k | Classification, feature analysis | CSV/XLSX |
| Hotel Reservations | 36.3k | Cancellation classification, cleaning | CSV/XLSX |
| Social-Media Wellbeing | 705 | Small, simple relatable examples | CSV/XLSX |

Every dataset exists in three forms:
- **Clean** — standardized types, labels, ISO dates, sentinel fixes (faithful reference).
- **Dirty** — seeded, documented issues (missing, duplicates, mixed case, currency
  text, mixed date formats, impossible values) for analysis practice.
- **SQL** — loaded into `ecommerce.db` / `all_datasets.db`.

The **e-commerce** dataset is the relational centrepiece: five tables
(`customers`, `sellers`, `products`, `deliveries`, `transactions`) joined by
primary/foreign keys, with indexes and a worked set of JOIN / aggregation / window
queries in `Datasets/SQL/ecommerce_sample_queries.sql`. Its keys were reconciled
so 100% of transactions resolve (source had ~70% blank foreign keys and a
`DEL` date-format mismatch).

---

## How to use the workbooks

- **Section_Examples** — open `WeekN_Section_Examples.xlsx`; each module is a tab with
  1–2 worked examples drawn from a real dataset, including **live Excel formulas**
  (e.g. `=IF(D8>30000,"High","Standard")`, `=SUM(D8:D15)`, margin `%`).
- **Assignments** — `WeekN_Assignments.xlsx` mirrors the modules with tasks + a
  **Answer Key** tab of sample solutions.
- **Projects** — see `Projects_Instructor/README.md` and
  `Projects_Students/README.md`.

## Reproduce everything

```
cd build
python3 clean_datasets.py      # Clean/ (standardized)
python3 reconcile_ecommerce.py # fix e-commerce foreign keys
python3 dirty_datasets.py      # Dirty/ (intentional issues)
python3 build_sql.py           # SQL/ (ecommerce.db + all_datasets.db)
python3 build_reports.py       # Reports/ (PDF catalogue) + data_dictionary.xlsx
python3 make_workbooks.py      # Section_Examples/ + Assignments/
python3 build_projects.py      # Projects_Instructor/ + Projects_Students/
python3 clean_presentations.py # Presentations_Clean/ (remove labs + relabel data)
```

All generators are deterministic (fixed seeds) so the outputs are reproducible.
