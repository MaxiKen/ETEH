# Datasets — ETEH Bootcamp data layer

Nine real-world datasets, each delivered in three states plus SQL, so learners can
practise cleaning, analysis and relational modelling on consistent, traceable data.

## Sub-folders
- **Raw/** — the nine supplied files, copied verbatim for provenance.
- **Clean/** — standardized, analysis-ready CSV + XLSX. Types fixed, categories
  canonicalised, dates → ISO, sentinel values handled. This is the *gold* reference.
- **Dirty/** — intentionally corrupted copies (`*_dirty.csv`) for cleaning/QA drills.
  See **dirty_manifest.md** for the exact, seeded issues in each file.
- **SQL/** — `ecommerce.db` (relational marketplace) and `all_datasets.db`
  (every dataset as a flat table). Plus `ecommerce_schema.sql` and
  `ecommerce_sample_queries.sql`.
- **Reports/** — `data_catalogue.pdf`, a printed-ready dictionary of every dataset.

## The datasets

| File (Clean/) | Rows | Notes |
|---|---|---|
| `ecommerce_*.csv` (5 tables) | 37k txns | Relational; keys reconciled; primary SQL dataset |
| `healthcare_records.csv` | 55,500 | Names/doctors normal-cased; billing absolute; LOS added |
| `superstore_sales.csv` | 9,994 | ISO dates; `ship_days`, `profit_margin`, `year` derived |
| `student_performance.csv` | 6,607 | `exam_score` capped at 100; impossible flagged |
| `salary_data.csv` | 6,704 | Education labels unified; job titles title-cased |
| `student_depression.csv` | 27,901 | `suicidal_thoughts` etc. canonicalised |
| `diabetes_prediction.csv` | 100,000 | `smoking_history` unified |
| `hotel_reservations.csv` | 36,275 | `booking_status` spacing fixed; arrival date added |
| `social_media_addiction.csv` | 705 | Small, friendly examples |

## E-Commerce → SQL
`ecommerce.db` contains five joined tables:

```
customers 1──< transactions >──1 products >──1 sellers
                     │
                     └──1 deliveries >──1 customers
```

Open it and try:
```
sqlite3 ecommerce.db < ecommerce_sample_queries.sql
```
or connect from Excel via **Get Data → From Database → SQLite/ODBC**.

## Using the dirty files
Pick a `*_dirty.csv`, clean it back to the Clean version, and reconcile your row
counts and totals. The `dirty_manifest.md` lists every injected defect so you can
build a marking checklist.
