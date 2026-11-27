# v2 revision progress

Topic copy files in `_build/rev/` (one per module). Build with `python3 _build/06_revise_slides.py`.
A module is "done" when its copy file exists, every figure has been checked against the data with pandas,
and the build prints no OVERFLOW / CLIP / TABLE / TITLE / NOTE warnings for it.

| Vol | Module | Topics | Status |
|-----|--------|--------|--------|
| 1 | 1.1 Introduction to Data Analytics | 14 | done |
| 1 | 1.2 The Four Types of Analytics | 6 | done |
| 1 | 1.3 The Data Analytics Process | 12 | done |
| 1 | 1.4 Applications Across Industries | 10 | done (1.4.9 uses clearly labelled illustrative figures: no mining dataset in the pack) |
| 1 | 1.5 Introduction to Microsoft Excel | 17 | done |
| 1 | 1.6 Working with Data in Excel | 14 | done |
| 1 | 1.7 Excel Formulas and Core Functions | 14 | done |
| 1 | 1.8 Organising Data with Excel Tables | 12 | done |
| 1 | 1.9 Tidy Data Laws | 12 | done |
| 2 | 2.1 Understanding External Data Sources | 8 | done |
| 2 | 2.2 Introduction to Power Query | 10 | done |
| 2 | 2.3 Connecting and Importing with Power Query | 8 | done |
| 2 | 2.4 Web and PDF Data Sources | 5 | done |
| 2 | 2.5 Data Cleaning with Excel | 12 | done |
| 2 | 2.6 Cleaning and Transforming in Power Query | 20 | done |
| 2 | 2.7 Excel Cleaning versus Power Query Cleaning | 6 | done |
| 2 | 2.8A The 65 Analytical Excel Functions (1–33) | 33 | done |
| 2 | 2.8B Text, Date, Lookup and Dynamic Array Functions (34–65) | 32 | done |
| 2 | 2.9 Introduction to the Power Query M Language | 11 | done |
| 2 | 2.10 Introduction to PivotTables | 19 | done |
| 2 | 2.11 Introduction to PivotCharts | 11 | done |
| 3 | 3.1 Database Fundamentals | 15 | done |
| 3 | 3.2 Introduction to SQL and SQLite | 17 | done |
| 3 | 3.3 ODBC Connections | 9 | done |
| 3 | 3.4 Joins, Merge and Append | 16 | done |
| 3 | 3.5 Data Model and Star Schema | 7 | done |
| 3 | 3.6A Power Pivot and DAX Foundations | 21 | done |
| 3 | 3.6B DAX Time Intelligence, KPIs and Performance | 20 | done |
| 3 | 3.7 Chart Anatomy and Chart Types | 26 | done |
| 3 | 3.8 Dashboard Development | 19 | done |
| 3 | 3.9 Dashboard Blueprints | 3 | done |
| 3 | 3.10 File Extracts into SQL, DAX and Dashboards | 5 | done |
| V4 | 4.1 | 14 | done (workbook note fixed: mean >4x median; rounding example uses real mean) |
| V4 | 4.2 | 9 | done (workbook 4.2.7 insight corrected) |
| V4 | 4.3 | 10 | done (workbook 4.3.8 rebuilt: seniority ruled out, experience-band confounder table added) |
| V4 | 4.4 | 12 | done (workbook: residual insight corrected, experience range 0-34) |
| V4 | 4.5 | 13 | done |
| V4 | 4.6 | 12 | done (workbook Solver insight corrected) |
| V4 | 4.7 | 9 | done (workbook: zone growth, invented refurbishment caveat, misleading-axis chart fixed) |
| V4 | 4.8 | 10 | done (workbook recommendation now matches data: Q2 peak, Oncology) |
| V4 | 4.9 | 10 | done |
| V4 | 4.10 | 8 | done |
| V4 | 4.11 | 10 | done (workbook provenance log: inflation corrected to ~5.2x) |
| V4 | 4.12 | 8 | done |
| V4 | 4.13 | 9 | done |
| V4 | 4.14 | 5 | done (workbook 4.14.2 insight: months-vs-FY rounding gap up to 0.2m) |
| V3 | fix | - | 3.2 HAVING slide: 85 store names with delivered sales (was 86) |
| V5 | P.1 | 12 | done (v1 '₦48,000 cash freed' / 'nine visuals' replaced with workbook figures) |
| V5 | P.2 | 5 | done |
| V5 | P.3 | 10 | done (points at Capstone 1 SOLUTION sheets) |
| V5 | P.4 | 10 | done (Crestview: all figures re-derived from the student CSV) |
| V5 | P.5 | 10 | done (P.5.7 retitled: v1 'Triage Bands' was not in the data; 06 Next-link honours title=) |
| V5 | P.6 | 8 | done (v1 '41-record exclusion', '3 min 20 s', 'J. Adeyemi' removed) |
| V5 | P.7 | 8 | done (v1 'Year-1 Engineering fell 3 points' and 'Ward 3' removed) |
| V5 | P.8 | 9 | done (v1 '45 min to 6 s' removed; CV bullets from workbook) |
| V5 | P.9 | 8 | done (STAR uses the ₦7.19bn wrong join; v1 'mean CGPA', '8% missing', '96 SKUs' removed) |
| V5 | fix | - | Capstone workbook P.6 limitation + P.9 STAR wording; rebuilt, verify 0 mismatches |
| V2-V4 | fix | - | Glyphs missing from Carlito (≥ ≤ ≈ √ α ▲▼) replaced in 2.8.10/24/25, 3.6.37, 3.8.7, 4.1.13, 4.2.5, 4.3.2, 4.4.1, 4.5.5, 4.6.3/9/11 |
