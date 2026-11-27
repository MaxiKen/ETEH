-- ETEH Week 3 - SQL examples (SQLite). Run against the database named above each query.
-- Files: Week3_Data_Files/supermart_online.db, Week3_Data_Files/lifecare_hospital.db

-- 3.1.3 Primary key check - is every transaction_id unique?
-- database: supermart_online.db
SELECT COUNT(*) AS rows_, COUNT(DISTINCT transaction_id) AS distinct_ids,
       COUNT(*) - COUNT(DISTINCT transaction_id) AS duplicates
FROM transactions;

-- 3.1.4 Foreign key check - orders pointing at a product that does not exist (orphans)
-- database: supermart_online.db
SELECT COUNT(*) AS orphan_lines
FROM transactions t
LEFT JOIN products p ON p.product_id = t.product_id
WHERE p.product_id IS NULL;

-- 3.2.5 SELECT and FROM - first rows of products
-- database: supermart_online.db
SELECT product_id, product_name, category, price
FROM products
LIMIT 5;

-- 3.2.6 WHERE - Electronics priced above ₦500,000
-- database: supermart_online.db
SELECT product_id, product_name, brand_name, price
FROM products
WHERE category = 'Electronics' AND price > 500000
ORDER BY price DESC;

-- 3.2.7 AND / OR / NOT - Lagos or Abuja customers who are not Bronze
-- database: supermart_online.db
SELECT membership_tier, COUNT(*) AS customers
FROM customers
WHERE (state = 'Lagos' OR state = 'FCT') AND NOT membership_tier = 'Bronze'
GROUP BY membership_tier;

-- 3.2.8-9 ORDER BY + LIMIT - the 10 biggest orders
-- database: supermart_online.db
SELECT order_id, order_date, product_id, quantity, total_price
FROM transactions
ORDER BY total_price DESC
LIMIT 10;

-- 3.2.10 GROUP BY - revenue by category (completed and delivered orders only)
-- database: supermart_online.db
SELECT p.category,
       COUNT(*)            AS orders,
       SUM(t.total_price)  AS revenue
FROM transactions t
JOIN products p ON p.product_id = t.product_id
WHERE t.order_status IN ('Delivered', 'Completed')
GROUP BY p.category
ORDER BY revenue DESC;

-- 3.2.11 HAVING - sellers with more than ₦15m of delivered revenue
-- database: supermart_online.db
SELECT s.store_name, COUNT(*) AS orders, SUM(t.total_price) AS revenue
FROM transactions t
JOIN sellers s ON s.seller_id = t.seller_id
WHERE t.order_status = 'Delivered'
GROUP BY s.store_name
HAVING SUM(t.total_price) > 15000000
ORDER BY revenue DESC;

-- 3.2.12 Aliases and calculated columns - seller commission rate check
-- database: supermart_online.db
SELECT order_id, total_price, commission_percentage AS comm_pct,
       commission_amount,
       ROUND(total_price * commission_percentage / 100.0, 2) AS recalculated,
       commission_amount - ROUND(total_price * commission_percentage / 100.0, 2) AS diff
FROM transactions
LIMIT 6;

-- 3.2.13 DISTINCT - which couriers deliver for us?
-- database: supermart_online.db
SELECT DISTINCT courier_name
FROM deliveries
ORDER BY courier_name;

-- 3.2.14 CASE - order-value bands
-- database: supermart_online.db
SELECT CASE WHEN total_price < 50000 THEN '1 Under ₦50k'
            WHEN total_price < 250000 THEN '2 ₦50k-250k'
            WHEN total_price < 1000000 THEN '3 ₦250k-1m'
            ELSE '4 ₦1m+' END AS value_band,
       COUNT(*) AS orders, SUM(total_price) AS revenue
FROM transactions
GROUP BY value_band
ORDER BY value_band;

-- 3.2.15 Subquery - customers who spent more than the average customer
-- database: supermart_online.db
SELECT COUNT(*) AS above_average_customers
FROM (SELECT customer_id, SUM(total_price) AS spend
      FROM transactions GROUP BY customer_id) AS c
WHERE spend > (SELECT AVG(spend) FROM
                 (SELECT SUM(total_price) AS spend FROM transactions GROUP BY customer_id));

-- 3.2.16 String and date functions - revenue by year
-- database: supermart_online.db
SELECT strftime('%Y', order_date) AS order_year,
       COUNT(*) AS orders,
       SUM(total_price) AS revenue,
       ROUND(AVG(total_price), 0) AS avg_order
FROM transactions
WHERE order_status <> 'Cancelled'
GROUP BY order_year
ORDER BY order_year;

-- 3.2.17 A complete analytical query - 2025 zone scorecard
-- database: supermart_online.db
-- Business question: which customer zones drove 2025 net revenue, and how fast did we deliver?
SELECT c.zone                                         AS customer_zone,
       COUNT(DISTINCT t.customer_id)                  AS customers,
       COUNT(*)                                       AS orders,
       SUM(t.total_price)                             AS net_revenue,
       ROUND(SUM(t.total_price) * 1.0 / COUNT(*), 0)  AS avg_order_value,
       ROUND(AVG(julianday(d.delivered_date) - julianday(t.order_date)), 1) AS avg_days_to_deliver
FROM transactions t
JOIN customers  c ON c.customer_id = t.customer_id
LEFT JOIN deliveries d ON d.delivery_id = t.delivery_id
WHERE t.order_date BETWEEN '2025-01-01' AND '2025-12-31'
  AND t.order_status IN ('Delivered', 'Completed')
GROUP BY c.zone
ORDER BY net_revenue DESC;

-- 3.2 (hospital) Same skills on lifecare_hospital.db - ALOS and readmission by department
-- database: lifecare_hospital.db
SELECT department,
       COUNT(*) AS admissions,
       ROUND(AVG(length_of_stay), 1) AS alos_days,
       ROUND(100.0 * SUM(readmit_30d = 'Yes') / COUNT(*), 1) AS readmit_pct
FROM admissions
GROUP BY department
ORDER BY readmit_pct DESC;

-- 3.4 Inner join row count
SELECT COUNT(*) FROM transactions t JOIN deliveries d ON d.delivery_id = t.delivery_id;

-- 3.4 Left outer join row count
SELECT COUNT(*) FROM transactions t LEFT JOIN deliveries d ON d.delivery_id = t.delivery_id;

-- 3.4 Right outer join row count
SELECT COUNT(*) FROM deliveries d LEFT JOIN transactions t ON t.delivery_id = d.delivery_id;

-- 3.4 Full outer join row count
SELECT COUNT(*) FROM (SELECT t.order_id FROM transactions t LEFT JOIN deliveries d ON d.delivery_id = t.delivery_id UNION ALL SELECT d.order_id FROM deliveries d LEFT JOIN transactions t ON t.delivery_id = d.delivery_id WHERE t.transaction_id IS NULL);

-- 3.4 Left anti join row count
SELECT COUNT(*) FROM transactions t LEFT JOIN deliveries d ON d.delivery_id = t.delivery_id WHERE d.delivery_id IS NULL;

-- 3.4 Right anti join row count
SELECT COUNT(*) FROM deliveries d LEFT JOIN transactions t ON t.delivery_id = d.delivery_id WHERE t.transaction_id IS NULL;

-- 3.4.6 Left anti in practice - what happened to the orders that never shipped?
-- database: supermart_online.db
SELECT t.order_status, t.payment_status, COUNT(*) AS orders, SUM(t.total_price) AS value
FROM transactions t
LEFT JOIN deliveries d ON d.delivery_id = t.delivery_id
WHERE d.delivery_id IS NULL
GROUP BY t.order_status, t.payment_status;

-- 3.4.6b Left anti on customers - registered but never ordered
-- database: supermart_online.db
SELECT c.membership_tier, COUNT(*) AS customers_without_orders
FROM customers c
LEFT JOIN transactions t ON t.customer_id = c.customer_id
WHERE t.transaction_id IS NULL
GROUP BY c.membership_tier;

-- 3.4.10 Cartesian trap: wrong key
SELECT COUNT(*) FROM transactions t JOIN deliveries d ON d.customer_id = t.customer_id;

-- 3.4.15 Cardinality check - is order_id unique on BOTH sides before a one-to-one merge?
-- database: supermart_online.db
SELECT 'transactions' AS table_name, COUNT(*) AS rows_, COUNT(DISTINCT order_id) AS distinct_keys FROM transactions
UNION ALL
SELECT 'deliveries', COUNT(*), COUNT(DISTINCT order_id) FROM deliveries;

-- 3.10.3a Admissions whose doctor is not in the doctors table (orphans)
-- database: lifecare_hospital.db
SELECT COUNT(*) AS orphan_admissions
FROM admissions a
LEFT JOIN doctors d ON d.doctor_id = a.doctor_id
WHERE d.doctor_id IS NULL;

-- 3.10.3b Exception report - one row per rule, with counts
-- database: lifecare_hospital.db
SELECT 'Discharge before admission' AS rule, COUNT(*) AS rows_ FROM admissions WHERE discharge_date < admission_date
UNION ALL
SELECT 'Bill below zero', COUNT(*) FROM admissions WHERE billing_amount < 0
UNION ALL
SELECT 'Stay over 30 days', COUNT(*) FROM admissions WHERE length_of_stay > 30
UNION ALL
SELECT 'Room outside ward range', COUNT(*) FROM admissions a JOIN wards w ON w.ward_id = a.ward_id
       WHERE a.room_number NOT BETWEEN w.first_room AND w.last_room;
