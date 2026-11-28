-- ETEH Week 3 - SQL examples (SQLite), edition 3: every query shown on the Week 3 slides.
-- Run each against the database named above it (both databases are in this Ex3.2_Files folder) in DB Browser for SQLite or sqlite3.
-- The result printed on the slide is what the query returns on these files.

-- 3.1.3 Primary Keys
-- database: supermart_online.db
SELECT COUNT(*) AS rows_,
       COUNT(DISTINCT transaction_id) AS distinct_ids
FROM transactions;

-- 3.1.4 Foreign Keys
-- database: supermart_online.db
SELECT COUNT(*) AS orphan_lines
FROM transactions t
LEFT JOIN products p
  ON p.product_id = t.product_id
WHERE p.product_id IS NULL;

-- 3.1.7 Normalisation: Second and Third Normal Form
-- database: supermart_online.db
SELECT COUNT(*) AS customers_zone_differs
FROM customers c
JOIN nigeria_states s ON s.state = c.state
WHERE s.zone <> c.zone;

-- 3.1.8 Database Data Types
-- database: supermart_online.db
SELECT typeof(quantity)    AS quantity,
       typeof(total_price) AS total_price,
       typeof(order_date)  AS order_date
FROM transactions LIMIT 1;

-- 3.1.10 Views
-- database: supermart_online.db
SELECT customer_zone, COUNT(*) AS lines
FROM v_order_lines
GROUP BY customer_zone
ORDER BY lines DESC;

-- 3.2.1 What Is SQL?
-- database: supermart_online.db
SELECT COUNT(*)         AS order_lines,
       SUM(total_price) AS gross_value
FROM transactions;

-- 3.2.5 SELECT and FROM
-- database: supermart_online.db
SELECT product_id, product_name,
       category, price
FROM products
LIMIT 4;

-- 3.2.6 Filtering Rows with WHERE
-- database: supermart_online.db
SELECT 'delivery_id IS NULL' AS test, COUNT(*) AS lines
FROM transactions WHERE delivery_id IS NULL
UNION ALL
SELECT 'delivery_id = NULL', COUNT(*)
FROM transactions WHERE delivery_id = NULL;

-- 3.2.7 Combining Conditions: AND, OR, NOT
-- database: supermart_online.db
SELECT COUNT(*) AS customers
FROM customers
WHERE (membership_tier = 'Gold'
       OR membership_tier = 'Platinum')
  AND zone = 'South-West';

-- 3.2.8 Sorting with ORDER BY
-- database: supermart_online.db
SELECT order_id, order_date, total_price
FROM transactions
ORDER BY total_price DESC
LIMIT 3;

-- 3.2.10 Aggregating with GROUP BY
-- database: supermart_online.db
SELECT p.category,
       COUNT(*)           AS lines,
       SUM(t.total_price) AS revenue
FROM transactions t
JOIN products p ON p.product_id = t.product_id
WHERE t.order_status IN ('Delivered','Completed')
GROUP BY p.category
ORDER BY revenue DESC
LIMIT 4;

-- 3.2.11 Filtering Groups with HAVING
-- database: supermart_online.db
SELECT seller_id, SUM(total_price) AS net_rev
FROM transactions
WHERE order_status IN ('Delivered','Completed')
GROUP BY seller_id
HAVING SUM(total_price) > 15000000
ORDER BY net_rev DESC
LIMIT 3;

-- 3.2.12 Aliases and Calculated Columns
-- database: supermart_online.db
SELECT COUNT(*) AS lines_checked,
       SUM(ABS(total_price * commission_percentage
               / 100.0 - commission_amount) > 0.01)
         AS mismatches
FROM transactions;

-- 3.2.13 Removing Duplicates with DISTINCT
-- database: supermart_online.db
SELECT COUNT(*)                     AS deliveries,
       COUNT(DISTINCT courier_name) AS couriers,
       COUNT(DISTINCT state)        AS states
FROM deliveries;

-- 3.2.14 Conditional Logic with CASE
-- database: supermart_online.db
SELECT CASE WHEN total_price < 50000  THEN '1 <50k'
            WHEN total_price < 250000 THEN '2 50k-250k'
            WHEN total_price < 1000000 THEN '3 250k-1m'
            ELSE '4 1m+' END AS band,
       COUNT(*) AS lines, SUM(total_price) AS value
FROM transactions
GROUP BY band;

-- 3.2.15 Subqueries
-- database: supermart_online.db
SELECT COUNT(*) AS above_average
FROM (SELECT customer_id, SUM(total_price) AS spend
      FROM transactions GROUP BY customer_id)
WHERE spend > (SELECT AVG(spend) FROM
      (SELECT SUM(total_price) AS spend
       FROM transactions GROUP BY customer_id));

-- 3.2.16 Useful SQL String and Date Functions
-- database: supermart_online.db
SELECT strftime('%Y', order_date) AS yr,
       COUNT(*)         AS lines,
       SUM(total_price) AS value
FROM transactions
GROUP BY yr
ORDER BY yr;

-- 3.2.17 Writing a Complete Analytical Query
-- database: supermart_online.db
SELECT customer_zone,
       COUNT(*)                   AS orders,
       SUM(total_price)           AS net_revenue,
       ROUND(AVG(days_to_deliver), 1) AS avg_days
FROM v_order_lines
WHERE order_status IN ('Delivered','Completed')
GROUP BY customer_zone
ORDER BY net_revenue DESC
LIMIT 4;

-- 3.4.2 Inner Join
-- database: supermart_online.db
SELECT COUNT(*)           AS lines,
       SUM(t.total_price) AS value
FROM transactions t
JOIN deliveries d
  ON d.delivery_id = t.delivery_id;

-- 3.4.3 Left Outer Join
-- database: supermart_online.db
SELECT COUNT(*)                AS all_orders,
       COUNT(d.courier_name)   AS with_courier,
       SUM(d.delivery_id IS NULL) AS courier_null
FROM transactions t
LEFT JOIN deliveries d
  ON d.delivery_id = t.delivery_id;

-- 3.4.6 Left Anti Join
-- database: supermart_online.db
SELECT t.order_status, t.payment_status,
       COUNT(*) AS orders, SUM(t.total_price) AS value
FROM transactions t
LEFT JOIN deliveries d
  ON d.delivery_id = t.delivery_id
WHERE d.delivery_id IS NULL
GROUP BY 1, 2;

-- 3.4.15 Verifying Join Cardinality
-- database: supermart_online.db
SELECT 'transactions' AS tbl, COUNT(*) AS rows_,
       COUNT(DISTINCT order_id) AS keys
FROM transactions
UNION ALL
SELECT 'deliveries', COUNT(*), COUNT(DISTINCT order_id)
FROM deliveries;

-- 3.10.3 Intermediate: SQL joins and an exception report
-- database: a staging copy - load Mercyland_Episodes_2023_DIRTY.csv as stg_admissions_2023 (3.10.2), then
--           ATTACH DATABASE 'mercyland_hospital.db' AS prod;
SELECT 'Doctor not in doctors' AS rule, COUNT(*) AS n
FROM stg_admissions_2023 s
LEFT JOIN prod.doctors d ON d.consultant_name = s.Consultant_Name
WHERE d.consultant_id IS NULL
UNION ALL
SELECT 'Discharge before admission', COUNT(*)
FROM stg_admissions_2023
WHERE Discharge_Date <> '' AND Discharge_Date < Admit_Date
UNION ALL
SELECT 'Bill below zero', COUNT(*)
FROM stg_admissions_2023
WHERE CAST(REPLACE(Bill_Amount_NGN, ',', '') AS REAL) < 0;
