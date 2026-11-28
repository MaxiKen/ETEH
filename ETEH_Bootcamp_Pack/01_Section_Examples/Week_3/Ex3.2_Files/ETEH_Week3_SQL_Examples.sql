-- ETEH Week 3 - SQL examples (SQLite), edition 3.3: every query shown on the Week 3 slides.
-- Run each against the database named above it (both databases are in this Ex3.2_Files folder) in DB Browser for SQLite (Execute SQL tab).
-- The result printed on the slide is what the query returns on these files.

-- 3.1.3 Primary Keys
-- database: supermart_online.db
SELECT COUNT(*) AS rows_,
       COUNT(DISTINCT transaction_id) AS distinct_ids
FROM transactions;

-- 3.1.9 Database Data Types
-- database: supermart_online.db
SELECT typeof(quantity)    AS quantity,
       typeof(total_price) AS total_price,
       typeof(order_date)  AS order_date
FROM transactions LIMIT 1;

-- 3.1.11 Views
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

-- 3.2.6 SELECT and FROM
-- database: supermart_online.db
SELECT product_id, product_name,
       category, price
FROM products
LIMIT 4;

-- 3.2.7 Filtering Rows with WHERE
-- database: supermart_online.db
SELECT COUNT(*) AS ussd_lines
FROM transactions
WHERE payment_method = 'USSD';

-- 3.2.8 Combining Conditions: AND, OR, NOT
-- database: supermart_online.db
SELECT COUNT(*) AS customers
FROM customers
WHERE (membership_tier = 'Gold'
       OR membership_tier = 'Platinum')
  AND zone = 'South-West';

-- 3.2.9 Sorting with ORDER BY
-- database: supermart_online.db
SELECT order_id, order_date, total_price
FROM transactions
ORDER BY total_price DESC
LIMIT 3;

-- 3.2.11 Aggregating with GROUP BY
-- database: supermart_online.db
SELECT payment_method,
       COUNT(*)         AS lines,
       SUM(total_price) AS value
FROM transactions
GROUP BY payment_method
ORDER BY lines DESC
LIMIT 4;

-- 3.2.12 Joining Two Tables on a Key
-- database: supermart_online.db
SELECT t.order_id, p.product_name, p.category
FROM transactions t
JOIN products p ON p.product_id = t.product_id
LIMIT 3;

-- 3.2.13 Removing Duplicates with DISTINCT
-- database: supermart_online.db
SELECT COUNT(*)                     AS deliveries,
       COUNT(DISTINCT courier_name) AS couriers,
       COUNT(DISTINCT state)        AS states
FROM deliveries;

-- 3.2.14 Writing a Multi-Line Query
-- database: supermart_online.db
SELECT customer_zone,
       COUNT(*)         AS orders,
       SUM(total_price) AS net_revenue
FROM v_order_lines
WHERE order_status IN ('Delivered','Completed')
GROUP BY customer_zone
ORDER BY net_revenue DESC
LIMIT 4;

-- 3.10.3 Intermediate: simple SQL checks, then clean in Power Query
-- database: the staging table stg_admissions_2023 (load it as in 3.10.2)
SELECT COUNT(*) FROM stg_admissions_2023
WHERE Discharge_Date = '';

SELECT Ward, COUNT(*) AS n
FROM stg_admissions_2023
GROUP BY Ward;
