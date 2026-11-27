# ETEH Bootcamp — Dataset Guide

_Generated 2026-11-27 · all datasets standardised, with deliberately dirtied working copies for cleaning labs._

## Folder layout

```
datasets/
  clean/   # standardised 'gold' copies (CSV + XLSX)
  dirty/   # intentionally messy working copies to clean
  sql/     # eteh_commerce.db (SQLite) + queries.sql + expected outputs
  pdf/     # printable data dictionary & sample extracts
```

## ETEH Retail  (`eteh_retail`)

- **Domain:** Retail sales (Superstore)
- **Grain:** One row per product line on an order
- **Used in:** W1.6 data entry · W2 Pivots · W3 Dashboards · W4 Forecasting
- **Clean files:** eteh_retail.csv, eteh_retail.xlsx
- **Dirty files:** eteh_retail.csv, eteh_retail.xlsx
- **Intentional dirty issues (cleaning challenges):**
  - Mixed date formats in order_date / ship_date (ISO, US m/d/Y, 'Mon D, Y').
  - Category / sub_category have stray upper/lower-case and trailing spaces.
  - Ship Mode has typos ('Stan dard Class', 'SameDay').
  - City / State padded with leading/trailing whitespace.
  - ~60 fully duplicated rows.
  - Blank discount and profit cells.

## ETEH Commerce (SQL)  (`eteh_commerce`)

- **Domain:** Relational e-commerce (5 joined tables)
- **Grain:** One row per transaction line; keys: customer_id, seller_id, product_id, delivery_id
- **Used in:** W3 SQL, DAX & Dashboards (primary relational dataset)
- **Clean files:** eteh_commerce.db (SQLite), eteh_commerce.xlsx (5 sheets), eteh_commerce_flat.csv
- **Dirty files:** eteh_commerce_transactions.csv, eteh_commerce_products.csv
- **Intentional dirty issues (cleaning challenges):**
  - transaction_id duplicated (~26k repeats) in the raw file.
  - delivery_id key spaces do not align between tables (re-link via order_id).
  - A few orphan customer/product/seller keys.
  - order_status mixed case; commission_amount stored as text ('NGN 12,345.67').
  - order_date in mixed d/m/Y vs Y-m-d; 40 appended duplicate rows.

## ETEH Lifecare Hospital  (`eteh_lifecare`)

- **Domain:** Healthcare operations & billing
- **Grain:** One row per hospital admission
- **Used in:** W1 examples · W4 Descriptive stats & distributions · W3 Dashboards
- **Clean files:** eteh_lifecare.csv, eteh_lifecare.xlsx
- **Dirty files:** eteh_lifecare.csv, eteh_lifecare.xlsx
- **Intentional dirty issues (cleaning challenges):**
  - Patient names in random case ('Bobby JacksOn').
  - Admission / discharge dates in mixed formats.
  - Billing amount stored as text with currency symbol ('₦18,856.28').
  - Medical condition / test results mixed case; some blank test_results.
  - ~50 duplicated rows.

## ETEH University – Student Analytics  (`eteh_university`)

- **Domain:** Education / student outcomes
- **Grain:** One row per student (three related surveys)
- **Used in:** W4 Correlation, Regression & distributions (Crestview replacement)
- **Clean files:** eteh_university.xlsx (Performance, Wellbeing, Social_Media)
- **Dirty files:** eteh_university.xlsx, eteh_uni_performance.csv, eteh_uni_wellbeing.csv, eteh_uni_social.csv
- **Intentional dirty issues (cleaning challenges):**
  - Performance: gender / parental_involvement mixed case; blank attendance & scores.
  - Wellbeing: sleep_duration recoded inconsistently ('5 – 6 hrs').
  - Social: platform names mixed case; duplicate rows.

## ETEH Hotels  (`eteh_hotels`)

- **Domain:** Hospitality / bookings
- **Grain:** One row per room booking
- **Used in:** W4 Logistic-style cancellation analysis & pivot practice
- **Clean files:** eteh_hotels.csv, eteh_hotels.xlsx
- **Dirty files:** eteh_hotels.csv, eteh_hotels.xlsx
- **Intentional dirty issues (cleaning challenges):**
  - meal_plan / room_type mixed case and spacing ('meal plan 1', 'room type 4').
  - booking_status mixed case ('canceled' vs 'Canceled').
  - Blank price / special-request cells; ~45 duplicate rows.

## ETEH Salary  (`eteh_salary`)

- **Domain:** HR / compensation
- **Grain:** One row per employee
- **Used in:** W4 Linear Regression (simple worked example)
- **Clean files:** eteh_salary.csv, eteh_salary.xlsx
- **Dirty files:** eteh_salary.csv, eteh_salary.xlsx
- **Intentional dirty issues (cleaning challenges):**
  - Education Level inconsistent ('Bachelor's' vs "Bachelor's Degree", 'phD' vs 'PhD').
  - Blank age / gender / salary cells.
  - Duplicate rows (kept intentionally as the teaching point).

## ETEH Health (Diabetes)  (`eteh_health_diabetes`)

- **Domain:** Public-health screening
- **Grain:** One row per patient screening (10,000-row sample)
- **Used in:** W4 Distributions, correlation & classification practice
- **Clean files:** eteh_health_diabetes.csv, eteh_health_diabetes.xlsx
- **Dirty files:** eteh_health_diabetes.csv, eteh_health_diabetes.xlsx
- **Intentional dirty issues (cleaning challenges):**
  - smoking_history mixed case; gender lower-cased in places.
  - Duplicate rows; large original (100k) sampled to 10k for performance.

## Suggested module → dataset map

- **Week 1 (Foundations):** eteh_retail, eteh_lifecare (tidy-data & entry examples)
- **Week 2 (Power Query & Pivots):** eteh_retail (pivots), eteh_hotels (cleaning), eteh_commerce_flat (transforms)
- **Week 3 (SQL, DAX & Dashboards):** eteh_commerce.db (SQL joins), eteh_retail & eteh_lifecare (dashboards)
- **Week 4 (Stats, Forecast & AI):** eteh_salary (regression), eteh_university (correlation), eteh_health_diabetes & eteh_lifecare (distributions), eteh_hotels (cancellation analysis)
